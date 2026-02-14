use crate::common::errors::AppError;
use crate::ods::sheet_model::{Cell, CellValue, Sheet, Workbook};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;
use xmltree::{Element, EmitterConfig, XMLNode};

pub struct ContentXml;

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

        loop {
            let event = reader
                .read_event()
                .map_err(|e| AppError::XmlParseError(e.to_string()))?;
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
                        for row_idx in current_row..=target_row {
                            writer
                                .write_event(Event::Start(BytesStart::new("table:table-row")))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                            if row_idx == target_row {
                                for _ in 0..target_col {
                                    writer
                                        .write_event(Event::Empty(BytesStart::new(
                                            "table:table-cell",
                                        )))
                                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                                }
                                Self::write_value_cell(&mut writer, value, None)?;
                            } else {
                                writer
                                    .write_event(Event::Empty(BytesStart::new("table:table-cell")))
                                    .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                            }
                            writer
                                .write_event(Event::End(BytesEnd::new("table:table-row")))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
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
                        for row_idx in 0..=target_row {
                            writer
                                .write_event(Event::Start(BytesStart::new("table:table-row")))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                            if row_idx == target_row {
                                for _ in 0..target_col {
                                    writer
                                        .write_event(Event::Empty(BytesStart::new(
                                            "table:table-cell",
                                        )))
                                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                                }
                                Self::write_value_cell(&mut writer, value, None)?;
                            } else {
                                writer
                                    .write_event(Event::Empty(BytesStart::new("table:table-cell")))
                                    .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                            }
                            writer
                                .write_event(Event::End(BytesEnd::new("table:table-row")))
                                .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                        }
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
                    current_row_repeat = row_repeat;
                    in_target_row = in_target_sheet
                        && !value_written
                        && target_row >= current_row
                        && target_row < (current_row + row_repeat);
                    current_col = 0;
                    writer
                        .write_event(Event::Start(e.to_owned()))
                        .map_err(|e| AppError::XmlParseError(e.to_string()))?;
                }
                Event::End(e) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") => {
                    if in_target_row && !value_written && target_col >= current_col {
                        for _ in current_col..target_col {
                            writer
                                .write_event(Event::Empty(BytesStart::new("table:table-cell")))
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
                                for _ in 0..before {
                                    let before_tag = Self::clone_cell_without_repeat(&e);
                                    writer
                                        .write_event(Event::Empty(before_tag))
                                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                                }
                            }
                            Self::write_value_cell(&mut writer, value, Some(&e))?;
                            if after > 0 {
                                for _ in 0..after {
                                    let after_tag = Self::clone_cell_without_repeat(&e);
                                    writer
                                        .write_event(Event::Empty(after_tag))
                                        .map_err(|er| AppError::XmlParseError(er.to_string()))?;
                                }
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

    pub fn parse(content: &str) -> Result<Workbook, AppError> {
        // Parseamos solo el subconjunto necesario de ODS (tablas, filas, celdas y texto).
        let mut reader = Reader::from_str(content);
        reader.config_mut().trim_text(false);
        let mut sheets: Vec<Sheet> = Vec::new();

        let mut current_sheet: Option<Sheet> = None;
        let mut current_row: Option<Vec<Cell>> = None;
        let mut current_cell_value = CellValue::Empty;
        let mut row_repeat = 1usize;
        let mut cell_repeat = 1usize;
        let mut in_text_p = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    let mut name = "Sheet1".to_string();
                    for attr in e.attributes().flatten() {
                        if Self::is_local_name_bytes(attr.key.as_ref(), b"name") {
                            name = attr
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(|x| AppError::XmlParseError(x.to_string()))?
                                .to_string();
                        }
                    }
                    current_sheet = Some(Sheet::new(name));
                }
                Ok(Event::Empty(e)) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    let mut name = "Sheet1".to_string();
                    for attr in e.attributes().flatten() {
                        if Self::is_local_name_bytes(attr.key.as_ref(), b"name") {
                            name = attr
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(|x| AppError::XmlParseError(x.to_string()))?
                                .to_string();
                        }
                    }
                    sheets.push(Sheet::new(name));
                }
                Ok(Event::End(e)) if Self::is_local_name_bytes(e.name().as_ref(), b"table") => {
                    if let Some(sheet) = current_sheet.take() {
                        sheets.push(sheet);
                    }
                }
                Ok(Event::Start(e))
                    if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") =>
                {
                    row_repeat = Self::attr_repeat(&e, b"number-rows-repeated", reader.decoder());
                    current_row = Some(Vec::new());
                }
                Ok(Event::End(e)) if Self::is_local_name_bytes(e.name().as_ref(), b"table-row") => {
                    if let (Some(sheet), Some(row)) = (current_sheet.as_mut(), current_row.take()) {
                        for _ in 0..row_repeat {
                            sheet.rows.push(row.clone());
                        }
                    }
                    row_repeat = 1;
                }
                Ok(Event::Empty(e))
                    if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") =>
                {
                    if let Some(row) = current_row.as_mut() {
                        let value = Self::value_from_attrs(&e, reader.decoder());
                        let repeat =
                            Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                        for _ in 0..repeat {
                            row.push(Cell {
                                value: value.clone(),
                            });
                        }
                    }
                }
                Ok(Event::Start(e))
                    if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") =>
                {
                    current_cell_value = Self::value_from_attrs(&e, reader.decoder());
                    cell_repeat =
                        Self::attr_repeat(&e, b"number-columns-repeated", reader.decoder());
                }
                Ok(Event::Start(e)) if Self::is_local_name_bytes(e.name().as_ref(), b"p") => {
                    in_text_p = true;
                }
                Ok(Event::End(e)) if Self::is_local_name_bytes(e.name().as_ref(), b"p") => {
                    in_text_p = false;
                }
                Ok(Event::Text(text)) if in_text_p => {
                    let t = text
                        .unescape()
                        .map_err(|x| AppError::XmlParseError(x.to_string()))?
                        .into_owned();
                    match &mut current_cell_value {
                        CellValue::String(existing) => existing.push_str(&t),
                        CellValue::Empty => current_cell_value = CellValue::String(t),
                        CellValue::Number(_) | CellValue::Boolean(_) => {}
                    }
                }
                Ok(Event::End(e))
                    if Self::is_local_name_bytes(e.name().as_ref(), b"table-cell") =>
                {
                    // Expandimos celdas repetidas para exponer siempre una matriz explícita.
                    if let Some(row) = current_row.as_mut() {
                        for _ in 0..cell_repeat {
                            row.push(Cell {
                                value: current_cell_value.clone(),
                            });
                        }
                    }
                    current_cell_value = CellValue::Empty;
                    cell_repeat = 1;
                    in_text_p = false;
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => return Err(AppError::XmlParseError(e.to_string())),
            }
        }

        Ok(Workbook { sheets })
    }

    pub fn render(workbook: &Workbook) -> Result<String, AppError> {
        // Writes the workbook model back to ODS content.xml syntax.
        let mut writer = Writer::new_with_indent(Cursor::new(Vec::new()), b' ', 2);
        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), None)))?;

        let mut root = BytesStart::new("office:document-content");
        root.push_attribute((
            "xmlns:office",
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        ));
        root.push_attribute((
            "xmlns:table",
            "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        ));
        root.push_attribute((
            "xmlns:text",
            "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
        ));
        root.push_attribute((
            "xmlns:calcext",
            "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
        ));
        root.push_attribute(("office:version", "1.2"));
        writer.write_event(Event::Start(root))?;

        writer.write_event(Event::Start(BytesStart::new("office:body")))?;
        writer.write_event(Event::Start(BytesStart::new("office:spreadsheet")))?;

        for sheet in &workbook.sheets {
            let mut table = BytesStart::new("table:table");
            table.push_attribute(("table:name", sheet.name.as_str()));
            writer.write_event(Event::Start(table))?;

            for row in &sheet.rows {
                writer.write_event(Event::Start(BytesStart::new("table:table-row")))?;
                for cell in row {
                    let mut cell_tag = BytesStart::new("table:table-cell");
                    let maybe_text = match &cell.value {
                        CellValue::String(v) => {
                            cell_tag.push_attribute(("office:value-type", "string"));
                            Some(v.clone())
                        }
                        CellValue::Number(v) => {
                            let n = v.to_string();
                            cell_tag.push_attribute(("office:value-type", "float"));
                            cell_tag.push_attribute(("office:value", n.as_str()));
                            Some(n)
                        }
                        CellValue::Boolean(v) => {
                            let b = if *v { "true" } else { "false" };
                            cell_tag.push_attribute(("office:value-type", "boolean"));
                            cell_tag.push_attribute(("office:boolean-value", b));
                            Some(b.to_string())
                        }
                        CellValue::Empty => None,
                    };

                    if let Some(text) = maybe_text {
                        writer.write_event(Event::Start(cell_tag))?;
                        writer.write_event(Event::Start(BytesStart::new("text:p")))?;
                        writer.write_event(Event::Text(BytesText::new(&text)))?;
                        writer.write_event(Event::End(BytesEnd::new("text:p")))?;
                        writer.write_event(Event::End(BytesEnd::new("table:table-cell")))?;
                    } else {
                        writer.write_event(Event::Empty(cell_tag))?;
                    }
                }
                writer.write_event(Event::End(BytesEnd::new("table:table-row")))?;
            }

            writer.write_event(Event::End(BytesEnd::new("table:table")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("office:spreadsheet")))?;
        writer.write_event(Event::End(BytesEnd::new("office:body")))?;
        writer.write_event(Event::End(BytesEnd::new("office:document-content")))?;

        let bytes = writer.into_inner().into_inner();
        String::from_utf8(bytes).map_err(|e| AppError::XmlParseError(e.to_string()))
    }

    pub fn render_preserving_original(
        workbook: &Workbook,
        original_content: &str,
    ) -> Result<String, AppError> {
        let mut root = match Element::parse(original_content.as_bytes()) {
            Ok(r) => r,
            Err(_) => return Self::render(workbook),
        };

        let Some(body) = Self::child_mut_by_local_name(&mut root, "body") else {
            return Self::render(workbook);
        };
        let Some(spreadsheet) = Self::child_mut_by_local_name(body, "spreadsheet") else {
            return Self::render(workbook);
        };

        // Keep non-table children (for example calculation settings, named expressions, etc.).
        spreadsheet
            .children
            .retain(|n| !matches!(n, XMLNode::Element(e) if Self::is_local_name(&e.name, "table")));

        for table in Self::render_table_elements(workbook) {
            spreadsheet.children.push(XMLNode::Element(table));
        }

        let mut out = Vec::new();
        root.write_with_config(
            &mut out,
            EmitterConfig::new()
                .perform_indent(true)
                .write_document_declaration(true),
        )
        .map_err(|e| AppError::XmlParseError(e.to_string()))?;
        String::from_utf8(out).map_err(|e| AppError::XmlParseError(e.to_string()))
    }

    pub fn sheet_names_from_content(original_content: &str) -> Result<Vec<String>, AppError> {
        let root = Element::parse(original_content.as_bytes())
            .map_err(|e| AppError::XmlParseError(e.to_string()))?;
        let body = Self::child_by_local_name(&root, "body")
            .ok_or_else(|| AppError::InvalidOdsFormat("missing office:body".to_string()))?;
        let spreadsheet = Self::child_by_local_name(body, "spreadsheet")
            .ok_or_else(|| AppError::InvalidOdsFormat("missing office:spreadsheet".to_string()))?;

        let mut names = Vec::new();
        for child in &spreadsheet.children {
            if let XMLNode::Element(table) = child {
                if !Self::is_local_name(&table.name, "table") {
                    continue;
                }
                if let Some(name) = Self::table_name(table) {
                    names.push(name);
                }
            }
        }
        Ok(names)
    }

    pub fn duplicate_sheet_preserving_styles(
        original_content: &str,
        source_name: Option<&str>,
        source_index: Option<usize>,
        new_sheet_name: &str,
    ) -> Result<String, AppError> {
        let mut root = Element::parse(original_content.as_bytes())
            .map_err(|e| AppError::XmlParseError(e.to_string()))?;
        let body = Self::child_mut_by_local_name(&mut root, "body")
            .ok_or_else(|| AppError::InvalidOdsFormat("missing office:body".to_string()))?;
        let spreadsheet = Self::child_mut_by_local_name(body, "spreadsheet")
            .ok_or_else(|| AppError::InvalidOdsFormat("missing office:spreadsheet".to_string()))?;

        let mut table_node_indices = Vec::new();
        let mut table_names = Vec::new();
        for (i, child) in spreadsheet.children.iter().enumerate() {
            if let XMLNode::Element(table) = child {
                if !Self::is_local_name(&table.name, "table") {
                    continue;
                }
                table_node_indices.push(i);
                table_names.push(Self::table_name(table).unwrap_or_else(|| "Sheet".to_string()));
            }
        }

        if table_names.iter().any(|n| n == new_sheet_name) {
            return Err(AppError::SheetNameAlreadyExists(new_sheet_name.to_string()));
        }

        let table_pos = if let Some(name) = source_name {
            table_names
                .iter()
                .position(|n| n == name)
                .ok_or_else(|| AppError::SheetNotFound(name.to_string()))?
        } else if let Some(index) = source_index {
            if index >= table_names.len() {
                return Err(AppError::SheetNotFound(index.to_string()));
            }
            index
        } else {
            return Err(AppError::InvalidInput(
                "missing source sheet selector".to_string(),
            ));
        };

        let source_node_index = table_node_indices[table_pos];
        let cloned_table = match &spreadsheet.children[source_node_index] {
            XMLNode::Element(table) => {
                let mut copy = table.clone();
                Self::set_table_name(&mut copy, new_sheet_name.to_string());
                XMLNode::Element(copy)
            }
            _ => {
                return Err(AppError::InvalidOdsFormat(
                    "source table node is not an element".to_string(),
                ))
            }
        };

        spreadsheet
            .children
            .insert(source_node_index + 1, cloned_table);

        let mut out = Vec::new();
        root.write_with_config(
            &mut out,
            EmitterConfig::new()
                .perform_indent(true)
                .write_document_declaration(true),
        )
        .map_err(|e| AppError::XmlParseError(e.to_string()))?;
        String::from_utf8(out).map_err(|e| AppError::XmlParseError(e.to_string()))
    }

    fn render_table_elements(workbook: &Workbook) -> Vec<Element> {
        let mut result = Vec::new();
        for sheet in &workbook.sheets {
            let mut table = Element::new("table:table");
            table
                .attributes
                .insert("table:name".to_string(), sheet.name.clone());

            for row in &sheet.rows {
                let mut row_el = Element::new("table:table-row");
                for cell in row {
                    let mut cell_el = Element::new("table:table-cell");
                    match &cell.value {
                        CellValue::String(v) => {
                            cell_el
                                .attributes
                                .insert("office:value-type".to_string(), "string".to_string());
                            let mut p = Element::new("text:p");
                            p.children.push(XMLNode::Text(v.clone()));
                            cell_el.children.push(XMLNode::Element(p));
                        }
                        CellValue::Number(v) => {
                            let n = v.to_string();
                            cell_el
                                .attributes
                                .insert("office:value-type".to_string(), "float".to_string());
                            cell_el
                                .attributes
                                .insert("office:value".to_string(), n.clone());
                            let mut p = Element::new("text:p");
                            p.children.push(XMLNode::Text(n));
                            cell_el.children.push(XMLNode::Element(p));
                        }
                        CellValue::Boolean(v) => {
                            let b = if *v { "true" } else { "false" }.to_string();
                            cell_el
                                .attributes
                                .insert("office:value-type".to_string(), "boolean".to_string());
                            cell_el
                                .attributes
                                .insert("office:boolean-value".to_string(), b.clone());
                            let mut p = Element::new("text:p");
                            p.children.push(XMLNode::Text(b));
                            cell_el.children.push(XMLNode::Element(p));
                        }
                        CellValue::Empty => {}
                    }
                    row_el.children.push(XMLNode::Element(cell_el));
                }
                table.children.push(XMLNode::Element(row_el));
            }
            result.push(table);
        }
        result
    }

    fn child_mut_by_local_name<'a>(
        element: &'a mut Element,
        local_name: &str,
    ) -> Option<&'a mut Element> {
        for child in &mut element.children {
            if let XMLNode::Element(e) = child {
                if Self::is_local_name(&e.name, local_name) {
                    return Some(e);
                }
            }
        }
        None
    }

    fn child_by_local_name<'a>(element: &'a Element, local_name: &str) -> Option<&'a Element> {
        for child in &element.children {
            if let XMLNode::Element(e) = child {
                if Self::is_local_name(&e.name, local_name) {
                    return Some(e);
                }
            }
        }
        None
    }

    fn table_name(table: &Element) -> Option<String> {
        for (key, value) in &table.attributes {
            if Self::is_local_name(key, "name") {
                return Some(value.clone());
            }
        }
        None
    }

    fn set_table_name(table: &mut Element, name: String) {
        let key = table
            .attributes
            .keys()
            .find(|k| Self::is_local_name(k, "name"))
            .cloned()
            .unwrap_or_else(|| "table:name".to_string());
        table.attributes.insert(key, name);
    }

    fn is_local_name(full_name: &str, local_name: &str) -> bool {
        full_name == local_name || full_name.rsplit(':').next() == Some(local_name)
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

    fn clone_cell_without_repeat(src: &BytesStart<'_>) -> BytesStart<'static> {
        let mut out = BytesStart::new("table:table-cell");
        for attr in src.attributes().flatten() {
            if Self::is_local_name_bytes(attr.key.as_ref(), b"number-columns-repeated") {
                continue;
            }
            out.push_attribute(attr);
        }
        out
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

    fn value_from_attrs(e: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder) -> CellValue {
        // Value type is represented by attributes; text is optional for numbers/booleans.
        let mut value_type: Option<String> = None;
        let mut value: Option<String> = None;
        let mut boolean_value: Option<String> = None;

        for attr in e.attributes().flatten() {
            let key = attr.key.as_ref();
            let decoded = match attr.decode_and_unescape_value(decoder) {
                Ok(v) => v.to_string(),
                Err(_) => continue,
            };
            if Self::is_local_name_bytes(key, b"value-type") {
                value_type = Some(decoded);
            } else if Self::is_local_name_bytes(key, b"value") {
                value = Some(decoded);
            } else if Self::is_local_name_bytes(key, b"boolean-value") {
                boolean_value = Some(decoded);
            }
        }

        match value_type.as_deref() {
            Some("float") => value
                .and_then(|v| v.parse::<f64>().ok())
                .map(CellValue::Number)
                .unwrap_or(CellValue::Empty),
            Some("boolean") => boolean_value
                .map(|v| v.eq_ignore_ascii_case("true"))
                .map(CellValue::Boolean)
                .unwrap_or(CellValue::Empty),
            Some("string") => value.map(CellValue::String).unwrap_or(CellValue::Empty),
            _ => CellValue::Empty,
        }
    }
}

struct TableBlock {
    start: usize,
    end: usize,
    name: String,
}
