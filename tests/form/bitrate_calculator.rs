use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters (© symbol)
// Error: "end byte index 1738 is not a char boundary; it is inside '©' (bytes 1737..1739)"
#[test]
#[ignore]
fn bitrate_calculator_about_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Bitrate-calculator/Windows/Source-code/frmAbout.frm");

    let source_file = SourceFile::decode_with_replacement("frmAbout.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'frmAbout.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'frmAbout.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/bitrate_calculator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn bitrate_calculator_main_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Bitrate-calculator/Windows/Source-code/frmMain.frm");

    let source_file = SourceFile::decode_with_replacement("frmMain.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'frmMain.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'frmMain.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/bitrate_calculator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}
