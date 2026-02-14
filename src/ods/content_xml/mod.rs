use crate::common::errors::AppError;
use crate::ods::sheet_model::{Cell, CellValue, Sheet, Workbook};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;

pub struct ContentXml;
mod workbook_xml;

impl ContentXml {
    pub fn resolve_merged_anchor_raw(
        original_content: &str,
        sheet_index: usize,
        target_row: usize,
        target_col: usize,
    ) -> Result<(usize, usize), AppError> {
        let mut reader = Reader::from_str(original_content);
        reader.config_mut().trim_text(false);

        let mut current_sheet: usize = 0;
        let mut in_target_sheet = false;
        let mut current_row: usize = 0;
        let mut current_row_repeat: usize = 1;
        let mut in_row = false;
        let mut current_col: usize = 0;

        loop {
            let event = reader
                .read_event()
                .map_err(|e| AppError::XmlParseError(e.to_string()))?;
            match event {
                Event::Start(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    in_target_sheet = current_sheet == sheet_index;
                    current_sheet += 1;
                }
                Event::Empty(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    in_target_sheet = current_sheet == sheet_index;
                    current_sheet += 1;
                    if in_target_sheet {
                        break;
                    }
                    in_target_sheet = false;
                }
                Event::End(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    if in_target_sheet {
                        break;
                    }
                }
                Event::Start(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") => {
                    if in_target_sheet {
                        current_row_repeat =
                            Self::attr_repeat(&e, b"number-rows-repeated", reader.decoder());
                        in_row = true;
                        current_col = 0;
                    }
                }
                Event::Empty(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") => {
                    if in_target_sheet {
                        let repeat =
                            Self::attr_repeat(&e, b"number-rows-repeated", reader.decoder());
                        if target_row >= current_row && target_row < current_row + repeat {
                            return Ok((target_row, target_col));
                        }
                        current_row += repeat;
                    }
                }
                Event::End(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") => {
                    if in_target_sheet && in_row {
                        current_row += current_row_repeat;
                        current_row_repeat = 1;
                        in_row = false;
                    }
                }
                Event::Start(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") => {
                    if in_target_sheet && in_row {
                        let col_repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        let row_span =
                            Self::attr_repeat(&e, b"number-rows-spanned", reader.decoder());
                        let col_span =
                            Self::attr_repeat(&e, b"number-columns-spanned", reader.decoder());
                        if let Some(anchor_col) = Self::find_spanned_anchor_col(
                            current_row,
                            current_col,
                            row_span,
                            col_span,
                            col_repeat,
                            target_row,
                            target_col,
                        ) {
                            return Ok((current_row, anchor_col));
                        }
                        current_col += col_repeat * col_span;
                    }
                }
                Event::Empty(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") => {
                    if in_target_sheet && in_row {
                        let col_repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        let row_span =
                            Self::attr_repeat(&e, b"number-rows-spanned", reader.decoder());
                        let col_span =
                            Self::attr_repeat(&e, b"number-columns-spanned", reader.decoder());
                        if let Some(anchor_col) = Self::find_spanned_anchor_col(
                            current_row,
                            current_col,
                            row_span,
                            col_span,
                            col_repeat,
                            target_row,
                            target_col,
                        ) {
                            return Ok((current_row, anchor_col));
                        }
                        current_col += col_repeat * col_span;
                    }
                }
                Event::Start(e)
                    if Self::is_local_name_bytes(e.name().as_ref(), b"covered-table-cell") =>
                {
                    if in_target_sheet && in_row {
                        let repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        current_col += repeat;
                    }
                }
                Event::Empty(e)
                    if Self::is_local_name_bytes(e.name().as_ref(), b"covered-table-cell") =>
                {
                    if in_target_sheet && in_row {
                        let repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        current_col += repeat;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        Ok((target_row, target_col))
    }

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
    pub fn sheet_names_from_content_raw(original_content: &str) -> Result<Vec<String>, AppError> {
        let tables = Self::find_table_blocks(original_content)?;
        Ok(tables.into_iter().map(|t| t.name).collect())
    }

    pub fn duplicate_sheet_preserving_styles_raw(
        original_content: &str,
        source_name: Option<&str>,
        source_index: Option<usize>,
        new_sheet_name: &str,
    ) -> Result<String, AppError> {
        let tables = Self::find_table_blocks(original_content)?;
        if tables.is_empty() {
            return Err(AppError::InvalidOdsFormat(
                "no table:table blocks found".to_string(),
            ));
        }
        if tables.iter().any(|t| t.name == new_sheet_name) {
            return Err(AppError::SheetNameAlreadyExists(new_sheet_name.to_string()));
        }

        let table_pos = if let Some(name) = source_name {
            tables
                .iter()
                .position(|t| t.name == name)
                .ok_or_else(|| AppError::SheetNotFound(name.to_string()))?
        } else if let Some(index) = source_index {
            if index >= tables.len() {
                return Err(AppError::SheetNotFound(index.to_string()));
            }
            index
        } else {
            return Err(AppError::InvalidInput(
                "missing source sheet selector".to_string(),
            ));
        };

        let source = &tables[table_pos];
        let source_xml = &original_content[source.start..source.end];
        let cloned = Self::rename_first_table_name(source_xml, new_sheet_name)?;

        let mut out = String::with_capacity(original_content.len() + cloned.len() + 8);
        out.push_str(&original_content[..source.end]);
        out.push_str(&cloned);
        out.push_str(&original_content[source.end..]);
        Ok(out)
    }

    pub fn rename_first_sheet_name_raw(
        original_content: &str,
        new_sheet_name: &str,
    ) -> Result<String, AppError> {
        let tables = Self::find_table_blocks(original_content)?;
        let first = tables
            .first()
            .ok_or_else(|| AppError::InvalidOdsFormat("no table:table blocks found".to_string()))?;
        let table_xml = &original_content[first.start..first.end];
        let renamed = Self::rename_first_table_name(table_xml, new_sheet_name)?;

        let mut out = String::with_capacity(original_content.len() + 32);
        out.push_str(&original_content[..first.start]);
        out.push_str(&renamed);
        out.push_str(&original_content[first.end..]);
        Ok(out)
    }

    fn is_local_name_bytes(full_name: &[u8], local_name: &[u8]) -> bool {
        if full_name == local_name {
            return true;
        }
        if let Some(pos) = full_name.iter().rposition(|b| *b == b':') {
            return &full_name[pos + 1..] == local_name;
        }
        false
    }

    fn find_spanned_anchor_col(
        current_row: usize,
        current_col: usize,
        row_span: usize,
        col_span: usize,
        col_repeat: usize,
        target_row: usize,
        target_col: usize,
    ) -> Option<usize> {
        for rep in 0..col_repeat {
            let start_col = current_col + (rep * col_span);
            let end_col = start_col + col_span - 1;
            if target_row >= current_row
                && target_row < current_row + row_span
                && target_col >= start_col
                && target_col <= end_col
            {
                return Some(start_col);
            }
        }
        None
    }

    fn find_table_blocks(content: &str) -> Result<Vec<TableBlock>, AppError> {
        let mut result = Vec::new();
        let mut pos = 0usize;
        while let Some(start) = Self::find_next_table_open(content, pos) {
            let start_tag_end = Self::find_tag_end(content, start)?;
            let start_tag = &content[start..start_tag_end];
            let name = Self::extract_table_name(start_tag)
                .ok_or_else(|| AppError::InvalidOdsFormat("table without name".to_string()))?;

            let self_closing = start_tag.trim_end().ends_with("/>");
            let end = if self_closing {
                start_tag_end
            } else {
                Self::find_matching_table_end(content, start_tag_end)?
            };

            result.push(TableBlock { start, end, name });
            pos = end;
        }
        Ok(result)
    }

    fn find_tag_end(content: &str, tag_start: usize) -> Result<usize, AppError> {
        let bytes = content.as_bytes();
        let mut i = tag_start;
        let mut in_quote = false;
        while i < bytes.len() {
            let b = bytes[i];
            if b == b'"' {
                in_quote = !in_quote;
            } else if b == b'>' && !in_quote {
                return Ok(i + 1);
            }
            i += 1;
        }
        Err(AppError::InvalidOdsFormat(
            "unterminated start tag".to_string(),
        ))
    }

    fn find_matching_table_end(content: &str, from: usize) -> Result<usize, AppError> {
        let mut depth = 1usize;
        let mut pos = from;
        while pos < content.len() {
            let next_open = Self::find_next_table_open(content, pos);
            let next_close = content[pos..].find("</table:table>").map(|i| pos + i);
            match (next_open, next_close) {
                (Some(o), Some(c)) if o < c => {
                    let end = Self::find_tag_end(content, o)?;
                    let tag = &content[o..end];
                    if !tag.trim_end().ends_with("/>") {
                        depth += 1;
                    }
                    pos = end;
                }
                (_, Some(c)) => {
                    depth -= 1;
                    let close_end = c + "</table:table>".len();
                    pos = close_end;
                    if depth == 0 {
                        return Ok(close_end);
                    }
                }
                _ => {
                    return Err(AppError::InvalidOdsFormat(
                        "unterminated table block".to_string(),
                    ))
                }
            }
        }
        Err(AppError::InvalidOdsFormat(
            "unterminated table block".to_string(),
        ))
    }

    fn extract_table_name(start_tag: &str) -> Option<String> {
        Self::extract_attr_value(start_tag, "table:name")
            .or_else(|| Self::extract_attr_value(start_tag, "name"))
    }

    fn extract_attr_value(tag: &str, key: &str) -> Option<String> {
        let pattern = format!("{key}=\"");
        let start = tag.find(&pattern)? + pattern.len();
        let end_rel = tag[start..].find('"')?;
        Some(tag[start..start + end_rel].to_string())
    }

    fn rename_first_table_name(table_xml: &str, new_name: &str) -> Result<String, AppError> {
        let tag_end = Self::find_tag_end(table_xml, 0)?;
        let start_tag = &table_xml[..tag_end];
        let rest = &table_xml[tag_end..];

        let replaced = if start_tag.contains("table:name=\"") {
            Self::replace_attr_value(start_tag, "table:name", new_name)?
        } else if start_tag.contains("name=\"") {
            Self::replace_attr_value(start_tag, "name", new_name)?
        } else {
            return Err(AppError::InvalidOdsFormat(
                "source table missing name attribute".to_string(),
            ));
        };

        let mut out = String::with_capacity(table_xml.len() + 32);
        out.push_str(&replaced);
        out.push_str(rest);
        Ok(out)
    }

    fn replace_attr_value(tag: &str, key: &str, new_value: &str) -> Result<String, AppError> {
        let pattern = format!("{key}=\"");
        let start = tag.find(&pattern).ok_or_else(|| {
            AppError::InvalidOdsFormat(format!("missing attribute in table start tag: {key}"))
        })?;
        let value_start = start + pattern.len();
        let end_rel = tag[value_start..].find('"').ok_or_else(|| {
            AppError::InvalidOdsFormat("unterminated attribute in table start tag".to_string())
        })?;
        let value_end = value_start + end_rel;

        let mut out = String::with_capacity(tag.len() + new_value.len());
        out.push_str(&tag[..value_start]);
        out.push_str(new_value);
        out.push_str(&tag[value_end..]);
        Ok(out)
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
        } else {
            // Synthetic cells should not inherit unrelated column defaults.
            cell.push_attribute(("table:style-name", "Default"));
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
        gap.push_attribute(("table:style-name", "Default"));
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

    fn find_next_table_open(content: &str, from: usize) -> Option<usize> {
        let needle = "<table:table";
        let bytes = content.as_bytes();
        let mut pos = from;
        while let Some(rel) = content[pos..].find(needle) {
            let idx = pos + rel;
            let after = idx + needle.len();
            if after >= bytes.len() {
                return Some(idx);
            }
            let next = bytes[after];
            if next == b'>' || next == b'/' || next.is_ascii_whitespace() {
                return Some(idx);
            }
            pos = after;
        }
        None
    }

    fn attr_repeat(e: &BytesStart<'_>, key: &[u8], decoder: quick_xml::encoding::Decoder) -> usize {
        // ODS can compress repeated rows/columns using repeat attributes.
        for attr in e.attributes().flatten() {
            if Self::is_local_name_bytes(attr.key.as_ref(), key) {
                if let Ok(v) = attr.decode_and_unescape_value(decoder) {
                    if let Ok(n) = v.parse::<usize>() {
                        return n.max(1);
                    }
                }
            }
        }
        1
    }
}

struct TableBlock {
    start: usize,
    end: usize,
    name: String,
}

struct RepeatRowCapture {
    row_start: BytesStart<'static>,
    row_repeat: usize,
    before: usize,
    after: usize,
    inner_events: Vec<Event<'static>>,
    depth: usize,
}
