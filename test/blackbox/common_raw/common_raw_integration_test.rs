use mcp_ods::common::fs::FsUtil;
use mcp_ods::ods::cell_address::CellAddress;

#[test]
fn common_raw_fs_util_resolve_ods_path_covers_relative_absolute_and_errors() {
    let rel = FsUtil::resolve_ods_path("./tmp/test.ods").expect("relative");
    assert!(rel.is_absolute());
    assert!(rel.to_string_lossy().ends_with("tmp\\test.ods") || rel.to_string_lossy().ends_with("tmp/test.ods"));

    let abs = std::env::current_dir().expect("cwd").join("x.ods");
    let abs_resolved = FsUtil::resolve_ods_path(abs.to_string_lossy().as_ref()).expect("absolute");
    assert_eq!(abs_resolved, abs);

    let empty = FsUtil::resolve_ods_path("").expect_err("empty");
    assert!(empty.to_string().contains("path is empty"));

    let bad_ext = FsUtil::resolve_ods_path("bad.txt").expect_err("ext");
    assert!(bad_ext.to_string().contains("expected .ods extension"));
}

#[test]
fn common_raw_cell_address_parse_and_to_a1_cover_edge_cases() {
    let a1 = CellAddress::parse("A1").expect("a1");
    assert_eq!(a1.row, 0);
    assert_eq!(a1.col, 0);
    assert_eq!(a1.to_a1(), "A1");

    let aa10 = CellAddress::parse("AA10").expect("aa10");
    assert_eq!(aa10.row, 9);
    assert_eq!(aa10.col, 26);
    assert_eq!(aa10.to_a1(), "AA10");

    let zzz999 = CellAddress::parse("ZZZ999").expect("zzz999");
    assert_eq!(zzz999.to_a1(), "ZZZ999");

    for invalid in ["", "1A", "A0", "A-1", "A 1", "$A$1", "A_1"] {
        assert!(CellAddress::parse(invalid).is_err(), "{invalid} should fail");
    }
}
