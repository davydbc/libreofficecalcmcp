use crate::common::errors::AppError;
use crate::ods::sheet_model::{Cell, CellValue, Sheet, Workbook};
use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::io::Cursor;

pub struct ContentXml;
mod cell_edit;
mod merged_anchor;
mod table_blocks;
mod workbook_xml;

impl ContentXml {
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
}
