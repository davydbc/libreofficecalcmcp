use super::*;

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
}
