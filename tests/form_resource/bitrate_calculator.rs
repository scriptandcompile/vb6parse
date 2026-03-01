use vb6parse::files::resource::FormResourceFile;

#[test]
fn bitrate_calculator_about_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/Bitrate-calculator/Windows/Source-code/frmAbout.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}

#[test]
fn bitrate_calculator_main_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/Bitrate-calculator/Windows/Source-code/frmMain.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}
