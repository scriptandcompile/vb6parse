use vb6parse::*;

#[test]
fn chess_brain_vb_chessbrainvb_project_load() {
    let project_file_bytes =
        include_bytes!("../data/ChessBrainVB/ChessBrainVB_V4_03a/Source/ChessBrainVB.vbp");

    let result = SourceFile::decode_with_replacement("ChessBrainVB.vbp", project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'ChessBrainVB.vbp': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/project/chess_brain_vb");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn chess_brain_vb_chessbrainvb_debug_project_load() {
    let project_file_bytes =
        include_bytes!("../data/ChessBrainVB/ChessBrainVB_V4_03a/Source/ChessBrainVB_debug.vbp");

    let result = SourceFile::decode_with_replacement("ChessBrainVB_debug.vbp", project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'ChessBrainVB_debug.vbp': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/project/chess_brain_vb");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn chess_brain_vb_chessbrainvb_pcode_project_load() {
    let project_file_bytes =
        include_bytes!("../data/ChessBrainVB/ChessBrainVB_V4_03a/Source/ChessBrainVB_PCode.vbp");

    let result = SourceFile::decode_with_replacement("ChessBrainVB_PCode.vbp", project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'ChessBrainVB_PCode.vbp': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/project/chess_brain_vb");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}
