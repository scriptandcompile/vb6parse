use vb6parse::*;

#[test]
fn audiostation_project_load() {
    let file_path = "./tests/data/audiostation/Audiostation/Audiostation.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let result = SourceFile::decode_with_replacement(file_path, &project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/project/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}
