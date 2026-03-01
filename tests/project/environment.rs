use vb6parse::*;

#[test]
fn environment_m2000_project_load() {
    let project_file_bytes = include_bytes!("../data/Environment/M2000.vbp");

    let result = SourceFile::decode_with_replacement("M2000.vbp", project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'M2000.vbp': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/project/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn environment_mexe_project_load() {
    let project_file_bytes = include_bytes!("../data/Environment/mexe.vbp");

    let result = SourceFile::decode_with_replacement("mexe.vbp", project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'mexe.vbp': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/project/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn environment_testme_project_load() {
    let project_file_bytes = include_bytes!("../data/Environment/testme.vbp");

    let result = SourceFile::decode_with_replacement("testme.vbp", project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'testme.vbp': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/project/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}
