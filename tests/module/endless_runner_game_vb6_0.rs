use vb6parse::*;

#[test]
fn endless_runner_game_vb6_0_module1_module_load() {
    let module_bytes = include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Module1.bas");

    let result = SourceFile::decode_with_replacement("Module1.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'Module1.bas': {e:?}"),
    };

    let (module_file_opt, failures) = ModuleFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.print();
        }

        panic!("Module parse had failures");
    }

    let module = module_file_opt.expect("Module should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}
