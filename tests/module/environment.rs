use vb6parse::*;


#[test]
fn environment_activex_module_load() {
    let module_bytes = include_bytes!("../data/Environment/ActiveX.bas");

    let result = SourceFile::decode_with_replacement("ActiveX.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'ActiveX.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_changeres_module_load() {
    let module_bytes = include_bytes!("../data/Environment/changeres.bas");

    let result = SourceFile::decode_with_replacement("changeres.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'changeres.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_dbx_module_load() {
    let module_bytes = include_bytes!("../data/Environment/dbX.bas");

    let result = SourceFile::decode_with_replacement("dbX.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'dbX.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_dc_module_load() {
    let module_bytes = include_bytes!("../data/Environment/dc.bas");

    let result = SourceFile::decode_with_replacement("dc.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'dc.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_eduardocircle_module_load() {
    let module_bytes = include_bytes!("../data/Environment/EduardoCircle.bas");

    let result = SourceFile::decode_with_replacement("EduardoCircle.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'EduardoCircle.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_fcall_module_load() {
    let module_bytes = include_bytes!("../data/Environment/Fcall.bas");

    let result = SourceFile::decode_with_replacement("Fcall.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'Fcall.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_getadaptsinfo_module_load() {
    let module_bytes = include_bytes!("../data/Environment/GetAdaptsInfo.bas");

    let result = SourceFile::decode_with_replacement("GetAdaptsInfo.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'GetAdaptsInfo.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_globalvars_module_load() {
    let module_bytes = include_bytes!("../data/Environment/globalvars.bas");

    let result = SourceFile::decode_with_replacement("globalvars.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'globalvars.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_gpp1_module_load() {
    let module_bytes = include_bytes!("../data/Environment/gpp1.bas");

    let result = SourceFile::decode_with_replacement("gpp1.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'gpp1.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_helpmod_module_load() {
    let module_bytes = include_bytes!("../data/Environment/helpmod.bas");

    let result = SourceFile::decode_with_replacement("helpmod.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'helpmod.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_isprinter_module_load() {
    let module_bytes = include_bytes!("../data/Environment/IsPrinter.bas");

    let result = SourceFile::decode_with_replacement("IsPrinter.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'IsPrinter.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_jpg_module_load() {
    let module_bytes = include_bytes!("../data/Environment/jpg.bas");

    let result = SourceFile::decode_with_replacement("jpg.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'jpg.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_libmemory_module_load() {
    let module_bytes = include_bytes!("../data/Environment/LibMemory.bas");

    let result = SourceFile::decode_with_replacement("LibMemory.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'LibMemory.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_mdlidispatch_module_load() {
    let module_bytes = include_bytes!("../data/Environment/mdlIDispatch.bas");

    let result = SourceFile::decode_with_replacement("mdlIDispatch.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'mdlIDispatch.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_mdqrcodegen_module_load() {
    let module_bytes = include_bytes!("../data/Environment/mdQRCodegen.bas");

    let result = SourceFile::decode_with_replacement("mdQRCodegen.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'mdQRCodegen.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_mdtlsnative_module_load() {
    let module_bytes = include_bytes!("../data/Environment/mdTlsNative.bas");

    let result = SourceFile::decode_with_replacement("mdTlsNative.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'mdTlsNative.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_mdtlsthunks_module_load() {
    let module_bytes = include_bytes!("../data/Environment/mdTlsThunks.bas");

    let result = SourceFile::decode_with_replacement("mdTlsThunks.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'mdTlsThunks.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_mexe_module_load() {
    let module_bytes = include_bytes!("../data/Environment/mexe.bas");

    let result = SourceFile::decode_with_replacement("mexe.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'mexe.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_mhlsrgb_module_load() {
    let module_bytes = include_bytes!("../data/Environment/mHLSRGB.bas");

    let result = SourceFile::decode_with_replacement("mHLSRGB.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'mHLSRGB.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_modgdiplusresize_module_load() {
    let module_bytes = include_bytes!("../data/Environment/modGDIPlusResize.bas");

    let result = SourceFile::decode_with_replacement("modGDIPlusResize.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'modGDIPlusResize.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_modobjectextender_module_load() {
    let module_bytes = include_bytes!("../data/Environment/modObjectExtender.bas");

    let result = SourceFile::decode_with_replacement("modObjectExtender.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'modObjectExtender.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_mod_text_module_load() {
    let module_bytes = include_bytes!("../data/Environment/Mod_Text.bas");

    let result = SourceFile::decode_with_replacement("Mod_Text.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'Mod_Text.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_modtrickunregcom_module_load() {
    let module_bytes = include_bytes!("../data/Environment/modTrickUnregCOM.bas");

    let result = SourceFile::decode_with_replacement("modTrickUnregCOM.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'modTrickUnregCOM.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_modtypeinfo_module_load() {
    let module_bytes = include_bytes!("../data/Environment/modTypeInfo.bas");

    let result = SourceFile::decode_with_replacement("modTypeInfo.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'modTypeInfo.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_module10_module_load() {
    let module_bytes = include_bytes!("../data/Environment/Module10.bas");

    let result = SourceFile::decode_with_replacement("Module10.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'Module10.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_module12_module_load() {
    let module_bytes = include_bytes!("../data/Environment/Module12.bas");

    let result = SourceFile::decode_with_replacement("Module12.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'Module12.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_mod_util_module_load() {
    let module_bytes = include_bytes!("../data/Environment/Mod_Util.bas");

    let result = SourceFile::decode_with_replacement("Mod_Util.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'Mod_Util.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_monitor_module_load() {
    let module_bytes = include_bytes!("../data/Environment/monitor.bas");

    let result = SourceFile::decode_with_replacement("monitor.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'monitor.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_newrnd_module_load() {
    let module_bytes = include_bytes!("../data/Environment/newRND.bas");

    let result = SourceFile::decode_with_replacement("newRND.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'newRND.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_pic_module_load() {
    let module_bytes = include_bytes!("../data/Environment/pic.bas");

    let result = SourceFile::decode_with_replacement("pic.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'pic.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_pipe_module_load() {
    let module_bytes = include_bytes!("../data/Environment/pipe.bas");

    let result = SourceFile::decode_with_replacement("pipe.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'pipe.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
#[ignore = "Fails due to char boundary parsing bug in source_stream.rs. Parser attempts to slice multi-byte UTF-8 characters at invalid byte boundaries."]
fn environment_servermod_module_load() {
    let module_bytes = include_bytes!("../data/Environment/ServerMod.bas");

    let result = SourceFile::decode_with_replacement("ServerMod.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'ServerMod.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}

#[test]
fn environment_sleep_module_load() {
    let module_bytes = include_bytes!("../data/Environment/sleep.bas");

    let result = SourceFile::decode_with_replacement("sleep.bas", module_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'sleep.bas': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/module/environment");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module);
}
