use crate::common::errors::AppError;
use std::path::{Path, PathBuf};

pub struct FsUtil;

impl FsUtil {
    // Normalizes relative paths and enforces the .ods extension contract.
    pub fn resolve_ods_path(path: &str) -> Result<PathBuf, AppError> {
        if path.trim().is_empty() {
            return Err(AppError::InvalidPath("path is empty".to_string()));
        }
        let input = Path::new(path);
        let abs = if input.is_absolute() {
            input.to_path_buf()
        } else {
            std::env::current_dir()?.join(input)
        };
        let ext = abs
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if ext != "ods" {
            return Err(AppError::InvalidPath(format!(
                "expected .ods extension: {}",
                abs.display()
            )));
        }
        Ok(abs)
    }
}
