use super::*;

impl ContentXml {
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

        let table_pos = Self::resolve_table_index(
            &tables,
            source_name,
            source_index,
            "missing source sheet selector",
        )?;

        let source = &tables[table_pos];
        let source_xml = &original_content[source.start..source.end];
        let cloned = Self::rename_first_table_name(source_xml, new_sheet_name)?;

        let mut out = String::with_capacity(original_content.len() + cloned.len() + 8);
        out.push_str(&original_content[..source.end]);
        out.push_str(&cloned);
        out.push_str(&original_content[source.end..]);
        Ok(out)
    }

    pub fn rename_sheet_preserving_styles_raw(
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

        let table_pos = Self::resolve_table_index(
            &tables,
            source_name,
            source_index,
            "missing source sheet selector",
        )?;
        let source = &tables[table_pos];
        let source_xml = &original_content[source.start..source.end];
        let renamed = Self::rename_first_table_name(source_xml, new_sheet_name)?;

        let mut out = String::with_capacity(original_content.len() + 32);
        out.push_str(&original_content[..source.start]);
        out.push_str(&renamed);
        out.push_str(&original_content[source.end..]);
        Ok(out)
    }

    pub fn delete_sheet_preserving_styles_raw(
        original_content: &str,
        source_name: Option<&str>,
        source_index: Option<usize>,
    ) -> Result<String, AppError> {
        let tables = Self::find_table_blocks(original_content)?;
        if tables.is_empty() {
            return Err(AppError::InvalidOdsFormat(
                "no table:table blocks found".to_string(),
            ));
        }
        if tables.len() == 1 {
            return Err(AppError::InvalidInput(
                "cannot delete the last remaining sheet".to_string(),
            ));
        }

        let table_pos = Self::resolve_table_index(
            &tables,
            source_name,
            source_index,
            "missing source sheet selector",
        )?;
        let source = &tables[table_pos];

        let mut out = String::with_capacity(original_content.len());
        out.push_str(&original_content[..source.start]);
        out.push_str(&original_content[source.end..]);
        Ok(out)
    }

    pub fn add_sheet_preserving_styles_raw(
        original_content: &str,
        sheet_name: &str,
        position: &str,
    ) -> Result<String, AppError> {
        let tables = Self::find_table_blocks(original_content)?;
        if tables.is_empty() {
            return Err(AppError::InvalidOdsFormat(
                "no table:table blocks found".to_string(),
            ));
        }
        if tables.iter().any(|t| t.name == sheet_name) {
            return Err(AppError::SheetNameAlreadyExists(sheet_name.to_string()));
        }

        let escaped_name = Self::escape_xml_attr(sheet_name);
        let new_table = format!("<table:table table:name=\"{escaped_name}\"/>");

        let insert_at = if position.eq_ignore_ascii_case("start") {
            tables[0].start
        } else {
            tables[tables.len() - 1].end
        };

        let mut out = String::with_capacity(original_content.len() + new_table.len() + 8);
        out.push_str(&original_content[..insert_at]);
        out.push_str(&new_table);
        out.push_str(&original_content[insert_at..]);
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

    fn resolve_table_index(
        tables: &[TableBlock],
        source_name: Option<&str>,
        source_index: Option<usize>,
        missing_selector_message: &str,
    ) -> Result<usize, AppError> {
        if let Some(name) = source_name {
            return tables
                .iter()
                .position(|t| t.name == name)
                .ok_or_else(|| AppError::SheetNotFound(name.to_string()));
        }
        if let Some(index) = source_index {
            if index >= tables.len() {
                return Err(AppError::SheetNotFound(index.to_string()));
            }
            return Ok(index);
        }
        Err(AppError::InvalidInput(missing_selector_message.to_string()))
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
        let bytes = tag.as_bytes();
        let mut from = 0usize;

        while let Some(rel) = tag[from..].find(&pattern) {
            let attr_start = from + rel;
            let prev_ok = if attr_start == 0 {
                true
            } else {
                matches!(
                    bytes[attr_start - 1],
                    b' ' | b'\t' | b'\n' | b'\r' | b'<'
                )
            };
            if prev_ok {
                let value_start = attr_start + pattern.len();
                let end_rel = tag[value_start..].find('"')?;
                return Some(tag[value_start..value_start + end_rel].to_string());
            }
            from = attr_start + key.len();
        }

        None
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
        out.push_str(&Self::escape_xml_attr(new_value));
        out.push_str(&tag[value_end..]);
        Ok(out)
    }

    fn escape_xml_attr(value: &str) -> String {
        value
            .replace('&', "&amp;")
            .replace('"', "&quot;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
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
}

struct TableBlock {
    start: usize,
    end: usize,
    name: String,
}
