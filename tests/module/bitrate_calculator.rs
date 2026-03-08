use vb6parse::*;

#[test]
fn bitrate_calculator_global_module_load() {
    let module_bytes =
        include_bytes!("../data/Bitrate-calculator/Windows/Source-code/modGlobal.bas");

    let result = SourceFile::decode_with_replacement("modGlobal.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'modGlobal.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/bitrate_calculator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn bitrate_calculator_hook_wheel_mouse_module_load() {
    let module_bytes =
        include_bytes!("../data/Bitrate-calculator/Windows/Source-code/modHookWheelMouse.bas");

    let result = SourceFile::decode_with_replacement("modHookWheelMouse.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'modHookWheelMouse.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/bitrate_calculator");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}
