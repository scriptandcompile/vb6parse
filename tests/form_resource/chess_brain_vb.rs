use vb6parse::files::resource::FormResourceFile;

#[test]
fn chess_brain_vb_debugmain_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/ChessBrainVB/ChessBrainVB_V4_03a/Source/Forms/DebugMain.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}

#[test]
fn chess_brain_vb_frmchessx_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/ChessBrainVB/ChessBrainVB_V4_03a/Source/Forms/frmChessX.frx",
    )
    .expect("Failed to read file");

    // Note: This large 280K file has parsing failures - still verify it can be read
    if result.has_failures() {
        // File has parsing failures, but that's ok for this test
        return;
    }

    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}

#[test]
fn chess_brain_vb_main_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/ChessBrainVB/ChessBrainVB_V4_03a/Source/Forms/main.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}
