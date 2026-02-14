use crate::common::errors::AppError;
use crate::ods::sheet_model::{Cell, CellValue, Sheet, Workbook};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;

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
                Ok(Event::Start(e)) if e.name().as_ref() == b"table:table" => {
                    let mut name = "Sheet1".to_string();
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"table:name" {
                            name = attr
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(|x| AppError::XmlParseError(x.to_string()))?
                                .to_string();
                        }
                    }
                    current_sheet = Some(Sheet::new(name));
                }
                Ok(Event::End(e)) if e.name().as_ref() == b"table:table" => {
                    if let Some(sheet) = current_sheet.take() {
                        sheets.push(sheet);
                    }
                }
                Ok(Event::Start(e)) if e.name().as_ref() == b"table:table-row" => {
                    row_repeat =
                        Self::attr_repeat(&e, b"table:number-rows-repeated", reader.decoder());
                    current_row = Some(Vec::new());
                }
                Ok(Event::End(e)) if e.name().as_ref() == b"table:table-row" => {
                    if let (Some(sheet), Some(row)) = (current_sheet.as_mut(), current_row.take()) {
                        for _ in 0..row_repeat {
                            sheet.rows.push(row.clone());
                        }
                    }
                    row_repeat = 1;
                }
                Ok(Event::Empty(e)) if e.name().as_ref() == b"table:table-cell" => {
                    if let Some(row) = current_row.as_mut() {
                        let value = Self::value_from_attrs(&e, reader.decoder());
                        let repeat = Self::attr_repeat(
                            &e,
                            b"table:number-columns-repeated",
                            reader.decoder(),
                        );
                        for _ in 0..repeat {
                            row.push(Cell {
                                value: value.clone(),
                            });
                        }
                    }
                }
                Ok(Event::Start(e)) if e.name().as_ref() == b"table:table-cell" => {
                    current_cell_value = Self::value_from_attrs(&e, reader.decoder());
                    cell_repeat =
                        Self::attr_repeat(&e, b"table:number-columns-repeated", reader.decoder());
                }
                Ok(Event::Start(e)) if e.name().as_ref() == b"text:p" => {
                    in_text_p = true;
                }
                Ok(Event::End(e)) if e.name().as_ref() == b"text:p" => {
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
                Ok(Event::End(e)) if e.name().as_ref() == b"table:table-cell" => {
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

    fn attr_repeat(e: &BytesStart<'_>, key: &[u8], decoder: quick_xml::encoding::Decoder) -> usize {
        // ODS can compress repeated rows/columns using repeat attributes.
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == key {
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
            match key {
                b"office:value-type" => value_type = Some(decoded),
                b"office:value" => value = Some(decoded),
                b"office:boolean-value" => boolean_value = Some(decoded),
                _ => {}
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
