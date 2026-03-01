use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

#[test]
fn chess_brain_vb_debugmain_form_load() {
    let form_file_bytes = include_bytes!("../data/ChessBrainVB/ChessBrainVB_V4_03a/Source/Forms/DebugMain.frm");

    let source_file = SourceFile::decode_with_replacement("DebugMain.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'DebugMain.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'DebugMain.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/chess_brain_vb");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn chess_brain_vb_frmchessx_form_load() {
    let form_file_bytes = include_bytes!("../data/ChessBrainVB/ChessBrainVB_V4_03a/Source/Forms/frmChessX.frm");

    let source_file = SourceFile::decode_with_replacement("frmChessX.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'frmChessX.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'frmChessX.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/chess_brain_vb");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn chess_brain_vb_main_form_load() {
    let form_file_bytes = include_bytes!("../data/ChessBrainVB/ChessBrainVB_V4_03a/Source/Forms/Main.frm");

    let source_file = SourceFile::decode_with_replacement("Main.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Main.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Main.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/chess_brain_vb");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}
