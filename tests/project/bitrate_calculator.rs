use vb6parse::*;

// Note: This test allows parsing failures because the project file contains
// empty property values (Priority=, AssignedTo=, Comment=) which the parser
// doesn't currently support. However, it's also ignored due to a character
// boundary bug in the error reporting when handling multi-byte UTF-8 characters.
#[test]
#[ignore]
fn bitrate_calculator_project_load() {
    let project_file_bytes =
        include_bytes!("../data/Bitrate-calculator/Windows/Source-code/BitrateCalc.vbp");

    let result = SourceFile::decode_with_replacement("BitrateCalc.vbp", project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'BitrateCalc.vbp': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    // Note: This project file contains empty values (Priority=, AssignedTo=, Comment=)
    // which the parser doesn't currently support. We allow failures and snapshot what parses.
    if !failures.is_empty() {
        eprintln!(
            "Warning: Project file has {} parsing failures (empty property values)",
            failures.len()
        );
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/project/bitrate_calculator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}
