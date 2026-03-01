use vb6parse::files::resource::FormResourceFile;

#[test]
fn cdiu_beat_up_editor_chatbox_frx() {
    let result = FormResourceFile::from_file("./tests/data/CdiuBeatUpEditor/ChatBox.frx")
        .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}

#[test]
fn cdiu_beat_up_editor_openroom_frx() {
    let result = FormResourceFile::from_file("./tests/data/CdiuBeatUpEditor/OpenRoom.frx")
        .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}

#[test]
fn cdiu_beat_up_editor_systemread_frx() {
    let result = FormResourceFile::from_file("./tests/data/CdiuBeatUpEditor/systemRead.frx")
        .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}

#[test]
fn cdiu_beat_up_editor_test_frx() {
    let result =
        FormResourceFile::from_file("./tests/data/CdiuBeatUpEditor/test.frx").expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}
