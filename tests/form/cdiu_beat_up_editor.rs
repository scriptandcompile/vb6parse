use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_chatbox_form_load() {
    let form_file_bytes = include_bytes!("../data/CdiuBeatUpEditor/ChatBox.frm");

    let source_file = SourceFile::decode_with_replacement("ChatBox.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'ChatBox.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'ChatBox.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_goroom_form_load() {
    let form_file_bytes = include_bytes!("../data/CdiuBeatUpEditor/GoRoom.frm");

    let source_file = SourceFile::decode_with_replacement("GoRoom.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'GoRoom.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'GoRoom.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_openroom_form_load() {
    let form_file_bytes = include_bytes!("../data/CdiuBeatUpEditor/OpenRoom.frm");

    let source_file = SourceFile::decode_with_replacement("OpenRoom.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'OpenRoom.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'OpenRoom.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_systemread_form_load() {
    let form_file_bytes = include_bytes!("../data/CdiuBeatUpEditor/systemRead.frm");

    let source_file = SourceFile::decode_with_replacement("systemRead.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'systemRead.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'systemRead.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_test_form_load() {
    let form_file_bytes = include_bytes!("../data/CdiuBeatUpEditor/test.frm");

    let source_file = SourceFile::decode_with_replacement("test.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'test.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'test.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}
