use crate::common::errors::AppError;
use serde::de::DeserializeOwned;
use serde::Serialize;
use serde_json::Value;

pub struct JsonUtil;

impl JsonUtil {
    pub fn from_value<T: DeserializeOwned>(value: Value) -> Result<T, AppError> {
        serde_json::from_value(value).map_err(|e| AppError::InvalidInput(e.to_string()))
    }

    pub fn to_value<T: Serialize>(value: T) -> Result<Value, AppError> {
        serde_json::to_value(value).map_err(|e| AppError::InvalidInput(e.to_string()))
    }
}
