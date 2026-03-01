use vb6parse::*;

#[test]
fn endless_runner_game_vb6_0_game_project_load() {
    let project_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/game.vbp");

    let result = SourceFile::decode_with_replacement("game.vbp", project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'game.vbp': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/project/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}
