use super::*;

impl ContentXml {
    pub fn set_cell_value_preserving_styles_raw(
        original_content: &str,
        sheet_index: usize,
        target_row: usize,
        target_col: usize,
        value: &CellValue,
    ) -> Result<String, AppError> {
        let mut reader = Reader::from_str(original_content);
        reader.config_mut().trim_text(false);
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        let mut current_sheet: usize = 0;
        let mut in_target_sheet = false;
        let mut current_row: usize = 0;
        let mut in_target_row = false;
        let mut current_col: usize = 0;
        let mut current_row_repeat: usize = 1;
        let mut value_written = false;

        let mut skip_cell_depth = 0usize;
        let mut repeat_row_capture: Option<RepeatRowCapture> = None;

        loop {
            let event = reader
                .read_event()
                .map_err(|e| AppError::XmlParseError(e.to_string()))?;

            if let Some(capture) = repeat_row_capture.as_mut() {
                let is_row_end = matches!(
                    &event,
                    Event::End(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row")
                );
                if is_row_end && capture.depth == 1 {
                    Self::emit_repeated_row_split(
                        &mut writer,
                        capture,
                        target_col,
                        value,
                        &mut value_written,
                    )?;
                    current_row += capture.row_repeat;
                    current_row_repeat = 1;
                    in_target_row = false;
                    repeat_row_capture = None;
                    continue;
                }

                if matches!(&event, Event::Start(_)) {
                    capture.depth += 1;
                } else if matches!(&event, Event::End(_)) && capture.depth > 1 {
                    capture.depth -= 1;
                }
                capture.inner_events.push(event.into_owned());
                continue;
            }

            match event {
                Event::Start(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    in_target_sheet = current_sheet == sheet_index;
                    current_sheet += 1;
                    writer
                        .write_event(Event::Start(e.to_owned()))
                        .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                }
                Event::End(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    if in_target_sheet && !value_written {
                        if target_row > current_row {
                            let rows_before = target_row - current_row;
                            let mut repeated_row = BytesStart::new("table:table-row");
                            let rows_text = rows_before.to_string();
                            repeated_row
                                .push_attribute(("table:number-rows-repeated", rows_text.as_str()));
                            writer
                                .write_event(Event::Empty(repeated_row))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                        writer
                            .write_event(Event::Start(BytesStart::new("table:table-row")))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        if target_col > 0 {
                            let gap = Self::default_gap_cell(target_col);
                            writer
                                .write_event(Event::Empty(gap))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                        Self::write_value_cell(&mut writer, value, None)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("table:table-row")))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        value_written = true;
                    }
                    writer
                        .write_event(Event::End(e.to_owned()))
                        .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                    in_target_sheet = false;
                    current_row = 0;
                }
                Event::Empty(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    in_target_sheet = current_sheet == sheet_index;
                    current_sheet += 1;
                    if in_target_sheet {
                        writer
                            .write_event(Event::Start(e.to_owned()))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        if target_row > 0 {
                            let mut repeated_row = BytesStart::new("table:table-row");
                            let rows_text = target_row.to_string();
                            repeated_row
                                .push_attribute(("table:number-rows-repeated", rows_text.as_str()));
                            writer
                                .write_event(Event::Empty(repeated_row))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                        writer
                            .write_event(Event::Start(BytesStart::new("table:table-row")))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        if target_col > 0 {
                            let gap = Self::default_gap_cell(target_col);
                            writer
                                .write_event(Event::Empty(gap))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                        Self::write_value_cell(&mut writer, value, None)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("table:table-row")))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        writer
                            .write_event(Event::End(BytesEnd::new("table:table")))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        value_written = true;
                    } else {
                        writer
                            .write_event(Event::Empty(e.to_owned()))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                    }
                    in_target_sheet = false;
                }
                Event::Start(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") => {
                    let row_repeat =
                        Self::attr_repeat(&e, b"number-rows-repeated", reader.decoder());
                    let in_range = in_target_sheet
                        && !value_written
                        && target_row >= current_row
                        && target_row < (current_row + row_repeat);
                    if in_range && row_repeat > 1 {
                        let before = target_row - current_row;
                        let after = row_repeat - before - 1;
                        repeat_row_capture = Some(RepeatRowCapture {
                            row_start: e.to_owned(),
                            row_repeat,
                            before,
                            after,
                            inner_events: Vec::new(),
                            depth: 1,
                        });
                        current_row_repeat = row_repeat;
                        in_target_row = false;
                        current_col = 0;
                        continue;
                    }
                    current_row_repeat = row_repeat;
                    in_target_row = in_range;
                    current_col = 0;
                    writer
                        .write_event(Event::Start(e.to_owned()))
                        .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                }
                Event::Empty(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") => {
                    let row_repeat =
                        Self::attr_repeat(&e, b"number-rows-repeated", reader.decoder());
                    let in_range = in_target_sheet
                        && !value_written
                        && target_row >= current_row
                        && target_row < (current_row + row_repeat);
                    if in_range {
                        let before = target_row - current_row;
                        let after = row_repeat - before - 1;
                        if before > 0 {
                            let before_row = Self::clone_row_with_repeat(&e, Some(before));
                            writer
                                .write_event(Event::Empty(before_row))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }

                        let target_row_tag = Self::clone_row_with_repeat(&e, None);
                        writer
                            .write_event(Event::Start(target_row_tag))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        if target_col > 0 {
                            let gap = Self::default_gap_cell(target_col);
                            writer
                                .write_event(Event::Empty(gap))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                        Self::write_value_cell(&mut writer, value, None)?;
                        writer
                            .write_event(Event::End(BytesEnd::new("table:table-row")))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        value_written = true;

                        if after > 0 {
                            let after_row = Self::clone_row_with_repeat(&e, Some(after));
                            writer
                                .write_event(Event::Empty(after_row))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                    } else {
                        writer
                            .write_event(Event::Empty(e.to_owned()))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                    }
                    current_row += row_repeat;
                    current_row_repeat = 1;
                    in_target_row = false;
                }
                Event::End(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") => {
                    if in_target_row && !value_written && target_col >= current_col {
                        let gap_count = target_col - current_col;
                        if gap_count > 0 {
                            let gap = Self::default_gap_cell(gap_count);
                            writer
                                .write_event(Event::Empty(gap))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                        Self::write_value_cell(&mut writer, value, None)?;
                        value_written = true;
                    }
                    writer
                        .write_event(Event::End(e.to_owned()))
                        .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                    current_row += current_row_repeat;
                    current_row_repeat = 1;
                    in_target_row = false;
                }
                Event::Empty(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") => {
                    if in_target_row && !value_written {
                        let repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        let range_end = current_col + repeat;
                        if target_col >= current_col && target_col < range_end {
                            let before = target_col.saturating_sub(current_col);
                            let after = range_end.saturating_sub(target_col + 1);
                            if before > 0 {
                                let before_tag = Self::clone_cell_with_repeat(&e, before);
                                writer
                                    .write_event(Event::Empty(before_tag))
                                    .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                            }
                            Self::write_value_cell(&mut writer, value, Some(&e))?;
                            if after > 0 {
                                let after_tag = Self::clone_cell_with_repeat(&e, after);
                                writer
                                    .write_event(Event::Empty(after_tag))
                                    .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                            }
                            value_written = true;
                        } else {
                            writer
                                .write_event(Event::Empty(e.to_owned()))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                        current_col = range_end;
                    } else {
                        writer
                            .write_event(Event::Empty(e.to_owned()))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                    }
                }
                Event::Start(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") => {
                    if skip_cell_depth > 0 {
                        skip_cell_depth += 1;
                        continue;
                    }

                    if in_target_row && !value_written {
                        let repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        let range_end = current_col + repeat;
                        if target_col >= current_col && target_col < range_end {
                            if repeat > 1 {
                                return Err(AppError::InvalidOdsFormat(
                                    "cannot safely edit repeated non-empty cell".to_string(),
                                ));
                            }
                            Self::write_value_cell(&mut writer, value, Some(&e))?;
                            value_written = true;
                            skip_cell_depth = 1;
                        } else {
                            writer
                                .write_event(Event::Start(e.to_owned()))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
                        current_col = range_end;
                    } else {
                        writer
                            .write_event(Event::Start(e.to_owned()))
                            .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                    }
                }
                Event::Start(e)
                    if Self::is_local_name_bytes(e.name().as_ref(), b"covered-table-cell") =>
                {
                    if in_target_row && !value_written {
                        let repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        current_col += repeat;
                    }
                    writer
                        .write_event(Event::Start(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::Empty(e)
                    if Self::is_local_name_bytes(e.name().as_ref(), b"covered-table-cell") =>
                {
                    if in_target_row && !value_written {
                        let repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        current_col += repeat;
                    }
                    writer
                        .write_event(Event::Empty(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::End(e)
                    if Self::is_local_name_bytes(e.name().as_ref(), b"covered-table-cell") =>
                {
                    writer
                        .write_event(Event::End(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::End(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") => {
                    if skip_cell_depth > 0 {
                        skip_cell_depth -= 1;
                        continue;
                    }
                    writer
                        .write_event(Event::End(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::Eof => break,
                other => {
                    if skip_cell_depth == 0 {
                        writer
                            .write_event(other.into_owned())
                            .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                    }
                }
            }
        }

        if !value_written {
            return Err(AppError::InvalidInput(
                "target cell could not be written in source xml".to_string(),
            ));
        }

        let bytes = writer.into_inner().into_inner();
        String::from_utf8(bytes).map_err(|e| AppError::XmlParseError(e.to_string()))
    }
    fn write_value_cell(
        writer: &mut Writer<Cursor<Vec<u8>>>,
        value: &CellValue,
        existing: Option<&BytesStart<'_>>,
    ) -> Result<(), AppError> {
        let mut cell = BytesStart::new("table:table-cell");
        if let Some(existing) = existing {
            for attr in existing.attributes().flatten() {
                let key = attr.key.as_ref();
                if Self::is_local_name_bytes(key, b"value-type")
                    || Self::is_local_name_bytes(key, b"value")
                    || Self::is_local_name_bytes(key, b"boolean-value")
                    || Self::is_local_name_bytes(key, b"number-columns-repeated")
                {
                    continue;
                }
                cell.push_attribute(attr);
            }
        }

        match value {
            CellValue::String(v) => {
                cell.push_attribute(("office:value-type", "string"));
                writer
                    .write_event(Event::Start(cell))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::Start(BytesStart::new("text:p")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::Text(BytesText::new(v)))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::End(BytesEnd::new("text:p")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::End(BytesEnd::new("table:table-cell")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
            }
            CellValue::Number(v) => {
                let s = v.to_string();
                cell.push_attribute(("office:value-type", "float"));
                cell.push_attribute(("office:value", s.as_str()));
                writer
                    .write_event(Event::Start(cell))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::Start(BytesStart::new("text:p")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::Text(BytesText::new(&s)))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::End(BytesEnd::new("text:p")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::End(BytesEnd::new("table:table-cell")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
            }
            CellValue::Boolean(v) => {
                let s = if *v { "true" } else { "false" };
                cell.push_attribute(("office:value-type", "boolean"));
                cell.push_attribute(("office:boolean-value", s));
                writer
                    .write_event(Event::Start(cell))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::Start(BytesStart::new("text:p")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::Text(BytesText::new(s)))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::End(BytesEnd::new("text:p")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                writer
                    .write_event(Event::End(BytesEnd::new("table:table-cell")))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
            }
            CellValue::Empty => {
                writer
                    .write_event(Event::Empty(cell))
                    .map_err(|e| AppError::XmlParseError(e.to_string()))?;
            }
        }
        Ok(())
    }

    fn clone_cell_with_repeat(src: &BytesStart<'_>, repeat: usize) -> BytesStart<'static> {
        let mut out = BytesStart::new("table:table-cell");
        for attr in src.attributes().flatten() {
            if Self::is_local_name_bytes(attr.key.as_ref(), b"number-columns-repeated") {
                continue;
            }
            out.push_attribute(attr);
        }
        if repeat > 1 {
            let n_text = repeat.to_string();
            out.push_attribute(("table:number-columns-repeated", n_text.as_str()));
        }
        out
    }

    fn default_gap_cell(repeat: usize) -> BytesStart<'static> {
        let mut gap = BytesStart::new("table:table-cell");
        if repeat > 1 {
            let cols_text = repeat.to_string();
            gap.push_attribute(("table:number-columns-repeated", cols_text.as_str()));
        }
        gap
    }

    fn clone_row_with_repeat(src: &BytesStart<'_>, repeat: Option<usize>) -> BytesStart<'static> {
        let mut out = BytesStart::new("table:table-row");
        for attr in src.attributes().flatten() {
            if Self::is_local_name_bytes(attr.key.as_ref(), b"number-rows-repeated") {
                continue;
            }
            out.push_attribute(attr);
        }
        if let Some(n) = repeat {
            if n > 1 {
                let n_text = n.to_string();
                out.push_attribute(("table:number-rows-repeated", n_text.as_str()));
            }
        }
        out
    }

    fn emit_repeated_row_split(
        writer: &mut Writer<Cursor<Vec<u8>>>,
        capture: &RepeatRowCapture,
        target_col: usize,
        value: &CellValue,
        value_written: &mut bool,
    ) -> Result<(), AppError> {
        if capture.before > 0 {
            let before_start =
                Self::clone_row_with_repeat(&capture.row_start, Some(capture.before));
            writer
                .write_event(Event::Start(before_start))
                .map_err(|e| AppError::XmlParseError(e.to_string()))?;
            Self::replay_row_inner(writer, &capture.inner_events, None, value)?;
            writer
                .write_event(Event::End(BytesEnd::new("table:table-row")))
                .map_err(|e| AppError::XmlParseError(e.to_string()))?;
        }

        let target_start = Self::clone_row_with_repeat(&capture.row_start, None);
        writer
            .write_event(Event::Start(target_start))
            .map_err(|e| AppError::XmlParseError(e.to_string()))?;
        let wrote = Self::replay_row_inner(writer, &capture.inner_events, Some(target_col), value)?;
        writer
            .write_event(Event::End(BytesEnd::new("table:table-row")))
            .map_err(|e| AppError::XmlParseError(e.to_string()))?;
        if !wrote {
            return Err(AppError::InvalidInput(
                "target cell could not be written in repeated row".to_string(),
            ));
        }
        *value_written = true;

        if capture.after > 0 {
            let after_start = Self::clone_row_with_repeat(&capture.row_start, Some(capture.after));
            writer
                .write_event(Event::Start(after_start))
                .map_err(|e| AppError::XmlParseError(e.to_string()))?;
            Self::replay_row_inner(writer, &capture.inner_events, None, value)?;
            writer
                .write_event(Event::End(BytesEnd::new("table:table-row")))
                .map_err(|e| AppError::XmlParseError(e.to_string()))?;
        }
        Ok(())
    }

    fn replay_row_inner(
        writer: &mut Writer<Cursor<Vec<u8>>>,
        inner: &[Event<'static>],
        target_col: Option<usize>,
        value: &CellValue,
    ) -> Result<bool, AppError> {
        let mut current_col = 0usize;
        let mut wrote = false;
        let mut skip_cell_depth = 0usize;

        for event in inner {
            match event {
                Event::Empty(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") => {
                    if let Some(tc) = target_col {
                        if !wrote {
                            let repeat = Self::attr_repeat_owned(e, b"number-columns-repeated");
                            let range_end = current_col + repeat;
                            if tc >= current_col && tc < range_end {
                                let before = tc.saturating_sub(current_col);
                                let after = range_end.saturating_sub(tc + 1);
                                if before > 0 {
                                    let before_tag = Self::clone_cell_with_repeat(e, before);
                                    writer
                                        .write_event(Event::Empty(before_tag))
                                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                                }
                                Self::write_value_cell(writer, value, Some(e))?;
                                if after > 0 {
                                    let after_tag = Self::clone_cell_with_repeat(e, after);
                                    writer
                                        .write_event(Event::Empty(after_tag))
                                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                                }
                                wrote = true;
                            } else {
                                writer
                                    .write_event(Event::Empty(e.to_owned()))
                                    .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                            }
                            current_col = range_end;
                            continue;
                        }
                    }

                    writer
                        .write_event(Event::Empty(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::Start(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") => {
                    if skip_cell_depth > 0 {
                        skip_cell_depth += 1;
                        continue;
                    }

                    if let Some(tc) = target_col {
                        if !wrote {
                            let repeat = Self::attr_repeat_owned(e, b"number-columns-repeated");
                            let range_end = current_col + repeat;
                            if tc >= current_col && tc < range_end {
                                if repeat > 1 {
                                    return Err(AppError::InvalidOdsFormat(
                                        "cannot safely edit repeated non-empty cell".to_string(),
                                    ));
                                }
                                Self::write_value_cell(writer, value, Some(e))?;
                                wrote = true;
                                skip_cell_depth = 1;
                                current_col = range_end;
                                continue;
                            }
                            current_col = range_end;
                        }
                    }

                    writer
                        .write_event(Event::Start(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::Start(e)
                    if Self::is_local_name_bytes(e.name().as_ref(), b"covered-table-cell") =>
                {
                    if let Some(tc) = target_col {
                        if !wrote {
                            let repeat = Self::attr_repeat_owned(e, b"number-columns-repeated");
                            current_col += repeat;
                            if tc < current_col {
                                return Err(AppError::InvalidInput(
                                    "target is a covered cell in merged range".to_string(),
                                ));
                            }
                        }
                    }
                    writer
                        .write_event(Event::Start(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::Empty(e)
                    if Self::is_local_name_bytes(e.name().as_ref(), b"covered-table-cell") =>
                {
                    if let Some(tc) = target_col {
                        if !wrote {
                            let repeat = Self::attr_repeat_owned(e, b"number-columns-repeated");
                            current_col += repeat;
                            if tc < current_col {
                                return Err(AppError::InvalidInput(
                                    "target is a covered cell in merged range".to_string(),
                                ));
                            }
                        }
                    }
                    writer
                        .write_event(Event::Empty(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::End(e)
                    if Self::is_local_name_bytes(e.name().as_ref(), b"covered-table-cell") =>
                {
                    writer
                        .write_event(Event::End(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Event::End(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") => {
                    if skip_cell_depth > 0 {
                        skip_cell_depth -= 1;
                        continue;
                    }
                    writer
                        .write_event(Event::End(e.to_owned()))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                other => {
                    if skip_cell_depth == 0 {
                        writer
                            .write_event(other.clone())
                            .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                    }
                }
            }
        }

        if let Some(tc) = target_col {
            if !wrote && tc >= current_col {
                let gap_count = tc - current_col;
                if gap_count > 0 {
                    let gap = Self::default_gap_cell(gap_count);
                    writer
                        .write_event(Event::Empty(gap))
                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                }
                Self::write_value_cell(writer, value, None)?;
                wrote = true;
            }
        }
        Ok(wrote)
    }

    fn attr_repeat_owned(e: &BytesStart<'_>, key: &[u8]) -> usize {
        for attr in e.attributes().flatten() {
            if Self::is_local_name_bytes(attr.key.as_ref(), key) {
                if let Ok(v) = std::str::from_utf8(attr.value.as_ref()) {
                    if let Ok(n) = v.parse::<usize>() {
                        return n.max(1);
                    }
                }
            }
        }
        1
    }
}

struct RepeatRowCapture {
    row_start: BytesStart<'static>,
    row_repeat: usize,
    before: usize,
    after: usize,
    inner_events: Vec<Event<'static>>,
    depth: usize,
}
