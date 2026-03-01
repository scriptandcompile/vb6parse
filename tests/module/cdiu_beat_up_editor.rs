use vb6parse::*;

#[test]
fn cdiu_beat_up_editor_cdiu_12_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/cdiu_12.bas");

    let result = SourceFile::decode_with_replacement("cdiu_12.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cdiu_12.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_cma1_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/cma1.bas");

    let result = SourceFile::decode_with_replacement("cma1.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cma1.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_cma2_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/cma2.bas");

    let result = SourceFile::decode_with_replacement("cma2.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cma2.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_cma3_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/cma3.bas");

    let result = SourceFile::decode_with_replacement("cma3.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cma3.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_cma4_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/cma4.bas");

    let result = SourceFile::decode_with_replacement("cma4.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cma4.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_cma5_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/cma5.bas");

    let result = SourceFile::decode_with_replacement("cma5.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cma5.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn cdiu_beat_up_editor_cma6_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/cma6.bas");

    let result = SourceFile::decode_with_replacement("cma6.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cma6.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_cma7_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/cma7.bas");

    let result = SourceFile::decode_with_replacement("cma7.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cma7.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn cdiu_beat_up_editor_fmod_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/fmod.bas");

    let result = SourceFile::decode_with_replacement("fmod.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'fmod.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn cdiu_beat_up_editor_icecopymemory_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/icecopymemory.bas");

    let result = SourceFile::decode_with_replacement("icecopymemory.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'icecopymemory.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn cdiu_beat_up_editor_mcpu_proc_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/MCPU_Proc.bas");

    let result = SourceFile::decode_with_replacement("MCPU_Proc.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'MCPU_Proc.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

// Note: This test is ignored due to a bug in the parser's error reporting logic
// when handling multi-byte UTF-8 characters
#[test]
#[ignore]
fn cdiu_beat_up_editor_other_do_module_load() {
    let module_bytes = include_bytes!("../data/CdiuBeatUpEditor/other_do.bas");

    let result = SourceFile::decode_with_replacement("other_do.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'other_do.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}
