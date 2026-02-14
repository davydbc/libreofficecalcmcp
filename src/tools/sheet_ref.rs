use crate::common::errors::AppError;
use crate::ods::sheet_model::Workbook;
use serde::de::Error as DeError;
use serde::{Deserialize, Deserializer};
use serde_json::Value;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetRef {
    Name { name: String },
    Index { index: usize },
}

impl SheetRef {
    pub fn as_name(&self) -> Option<&str> {
        match self {
            SheetRef::Name { name } => Some(name.as_str()),
            SheetRef::Index { .. } => None,
        }
    }

    pub fn as_index(&self) -> Option<usize> {
        match self {
            SheetRef::Name { .. } => None,
            SheetRef::Index { index } => Some(*index),
        }
    }

    pub fn resolve_in_names(&self, sheet_names: &[String]) -> Result<(usize, String), AppError> {
        match self {
            SheetRef::Name { name } => sheet_names
                .iter()
                .position(|n| n == name)
                .map(|idx| (idx, sheet_names[idx].clone()))
                .ok_or_else(|| AppError::SheetNotFound(name.clone())),
            SheetRef::Index { index } => {
                if *index >= sheet_names.len() {
                    Err(AppError::SheetNotFound(index.to_string()))
                } else {
                    Ok((*index, sheet_names[*index].clone()))
                }
            }
        }
    }

    pub fn resolve_in_workbook(&self, workbook: &Workbook) -> Result<(usize, String), AppError> {
        match self {
            SheetRef::Name { name } => workbook
                .sheet_index_by_name(name)
                .map(|idx| (idx, workbook.sheets[idx].name.clone()))
                .ok_or_else(|| AppError::SheetNotFound(name.clone())),
            SheetRef::Index { index } => {
                if *index >= workbook.sheets.len() {
                    Err(AppError::SheetNotFound(index.to_string()))
                } else {
                    Ok((*index, workbook.sheets[*index].name.clone()))
                }
            }
        }
    }
}

impl<'de> Deserialize<'de> for SheetRef {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let value = Value::deserialize(deserializer)?;
        parse_sheet_ref_value(value).map_err(D::Error::custom)
    }
}

fn parse_sheet_ref_value(value: Value) -> Result<SheetRef, String> {
    match value {
        Value::Object(map) => {
            if let Some(name_value) = map.get("name") {
                if let Some(name) = name_value.as_str() {
                    return Ok(SheetRef::Name {
                        name: name.to_string(),
                    });
                }
                return Err("sheet.name must be a string".to_string());
            }

            if let Some(index_value) = map.get("index") {
                return parse_index_value(index_value);
            }

            Err("sheet must include either name or index".to_string())
        }
        Value::String(text) => parse_sheet_ref_string(&text),
        Value::Number(_) => parse_index_value(&value),
        _ => Err("sheet must be an object or string".to_string()),
    }
}

fn parse_sheet_ref_string(text: &str) -> Result<SheetRef, String> {
    let trimmed = text.trim();
    if trimmed.starts_with('{') {
        let nested_value: Value = serde_json::from_str(trimmed)
            .map_err(|_| "sheet JSON string is not a valid object".to_string())?;
        return parse_sheet_ref_value(nested_value);
    }
    Ok(SheetRef::Name {
        name: text.to_string(),
    })
}

fn parse_index_value(value: &Value) -> Result<SheetRef, String> {
    match value {
        Value::Number(number) => {
            if let Some(index) = number.as_u64() {
                return usize::try_from(index)
                    .map(|idx| SheetRef::Index { index: idx })
                    .map_err(|_| "sheet.index is too large".to_string());
            }
            Err("sheet.index must be a non-negative integer".to_string())
        }
        Value::String(text) => text
            .trim()
            .parse::<usize>()
            .map(|index| SheetRef::Index { index })
            .map_err(|_| "sheet.index must be a non-negative integer".to_string()),
        _ => Err("sheet.index must be a non-negative integer".to_string()),
    }
}
