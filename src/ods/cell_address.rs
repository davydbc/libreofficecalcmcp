use crate::common::errors::AppError;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellAddress {
    pub row: usize,
    pub col: usize,
}

impl CellAddress {
    // Parses A1 notation (for example: B3, AA10) into zero-based indexes.
    pub fn parse(input: &str) -> Result<Self, AppError> {
        if input.trim().is_empty() {
            return Err(AppError::InvalidCellAddress("address is empty".to_string()));
        }

        let mut letters = String::new();
        let mut digits = String::new();
        for ch in input.chars() {
            if ch.is_ascii_alphabetic() {
                if !digits.is_empty() {
                    return Err(AppError::InvalidCellAddress(input.to_string()));
                }
                letters.push(ch.to_ascii_uppercase());
            } else if ch.is_ascii_digit() {
                digits.push(ch);
            } else {
                return Err(AppError::InvalidCellAddress(input.to_string()));
            }
        }

        if letters.is_empty() || digits.is_empty() {
            return Err(AppError::InvalidCellAddress(input.to_string()));
        }

        let col = letters
            .chars()
            .fold(0usize, |acc, c| acc * 26 + ((c as u8 - b'A') as usize + 1));
        let row_num: usize = digits
            .parse()
            .map_err(|_| AppError::InvalidCellAddress(input.to_string()))?;

        if row_num == 0 || col == 0 {
            return Err(AppError::InvalidCellAddress(input.to_string()));
        }

        Ok(Self {
            row: row_num - 1,
            col: col - 1,
        })
    }

    pub fn to_a1(self) -> String {
        // Converts zero-based column index back to base-26 spreadsheet letters.
        let mut col = self.col + 1;
        let mut letters = String::new();
        while col > 0 {
            let rem = (col - 1) % 26;
            letters.insert(0, (b'A' + rem as u8) as char);
            col = (col - 1) / 26;
        }
        format!("{}{}", letters, self.row + 1)
    }
}
