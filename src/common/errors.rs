use thiserror::Error;

#[derive(Debug, Error)]
pub enum AppError {
    #[error("invalid path: {0}")]
    InvalidPath(String),
    #[error("file not found: {0}")]
    FileNotFound(String),
    #[error("file already exists: {0}")]
    AlreadyExists(String),
    #[error("invalid ods format: {0}")]
    InvalidOdsFormat(String),
    #[error("sheet not found: {0}")]
    SheetNotFound(String),
    #[error("sheet name already exists: {0}")]
    SheetNameAlreadyExists(String),
    #[error("invalid cell address: {0}")]
    InvalidCellAddress(String),
    #[error("xml parse error: {0}")]
    XmlParseError(String),
    #[error("zip error: {0}")]
    ZipError(String),
    #[error("io error: {0}")]
    IoError(String),
    #[error("invalid input: {0}")]
    InvalidInput(String),
}

impl AppError {
    pub fn code(&self) -> i32 {
        match self {
            AppError::InvalidPath(_) => 1001,
            AppError::FileNotFound(_) => 1002,
            AppError::AlreadyExists(_) => 1003,
            AppError::InvalidOdsFormat(_) => 1004,
            AppError::SheetNotFound(_) => 1005,
            AppError::SheetNameAlreadyExists(_) => 1006,
            AppError::InvalidCellAddress(_) => 1007,
            AppError::XmlParseError(_) => 1008,
            AppError::ZipError(_) => 1009,
            AppError::IoError(_) => 1010,
            AppError::InvalidInput(_) => 1011,
        }
    }
}

impl From<std::io::Error> for AppError {
    fn from(value: std::io::Error) -> Self {
        Self::IoError(value.to_string())
    }
}

impl From<zip::result::ZipError> for AppError {
    fn from(value: zip::result::ZipError) -> Self {
        Self::ZipError(value.to_string())
    }
}

impl From<quick_xml::Error> for AppError {
    fn from(value: quick_xml::Error) -> Self {
        Self::XmlParseError(value.to_string())
    }
}
