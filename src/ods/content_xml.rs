use crate::common::errors::AppError;
use crate::ods::sheet_model::{Cell, CellValue, Sheet, Workbook};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;
use xmltree::{Element, EmitterConfig, XMLNode};

pub struct ContentXml;

impl ContentXml {
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
