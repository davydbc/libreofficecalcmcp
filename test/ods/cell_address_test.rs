use mcp_ods::common::errors::AppError;
use mcp_ods::ods::cell_address::CellAddress;

#[test]
fn parse_and_format_cell_address() {
    let a1 = CellAddress::parse("C12").expect("valid address");
    assert_eq!(a1.row, 11);
    assert_eq!(a1.col, 2);
    assert_eq!(a1.to_a1(), "C12");
}

#[test]
fn invalid_cell_address_is_rejected() {
    let err = CellAddress::parse("12A").expect_err("invalid");
    assert!(matches!(err, AppError::InvalidCellAddress(_)));
}
