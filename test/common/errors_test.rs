use mcp_ods::common::errors::AppError;

#[test]
fn app_error_codes_are_stable() {
    let cases = vec![
        (AppError::InvalidPath("x".to_string()), 1001),
        (AppError::FileNotFound("x".to_string()), 1002),
        (AppError::AlreadyExists("x".to_string()), 1003),
        (AppError::InvalidOdsFormat("x".to_string()), 1004),
        (AppError::SheetNotFound("x".to_string()), 1005),
        (AppError::SheetNameAlreadyExists("x".to_string()), 1006),
        (AppError::InvalidCellAddress("x".to_string()), 1007),
        (AppError::XmlParseError("x".to_string()), 1008),
        (AppError::ZipError("x".to_string()), 1009),
        (AppError::IoError("x".to_string()), 1010),
        (AppError::InvalidInput("x".to_string()), 1011),
    ];
    for (err, expected) in cases {
        assert_eq!(err.code(), expected);
    }
}

#[test]
fn app_error_from_conversions_are_mapped() {
    let io_err = std::io::Error::other("io");
    assert!(matches!(AppError::from(io_err), AppError::IoError(_)));

    let zip_err = zip::result::ZipError::FileNotFound;
    assert!(matches!(AppError::from(zip_err), AppError::ZipError(_)));

    let mut reader = quick_xml::Reader::from_str("<root><a></root>");
    let quick_xml_err = loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Eof) => panic!("expected xml parse error"),
            Ok(_) => continue,
            Err(err) => break err,
        }
    };
    assert!(matches!(
        AppError::from(quick_xml_err),
        AppError::XmlParseError(_)
    ));
}
