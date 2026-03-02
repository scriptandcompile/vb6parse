use vb6parse::*;

#[test]
fn endless_runner_game_vb6_0_game_class_load() {
    let class_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/game.cls");

    let result = SourceFile::decode_with_replacement("game.cls", class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'game.cls': {e:?}"),
    };

    let (class_file_opt, failures) = ClassFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class = class_file_opt.expect("Class should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/class/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn endless_runner_game_vb6_0_game2_class_load() {
    let class_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/game2.cls");

    let result = SourceFile::decode_with_replacement("game2.cls", class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'game2.cls': {e:?}"),
    };

    let (class_file_opt, failures) = ClassFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class = class_file_opt.expect("Class should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/class/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}
