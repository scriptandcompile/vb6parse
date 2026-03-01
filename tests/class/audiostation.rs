use vb6parse::*;

#[test]
fn audiostation_audio_endpoint_volume_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsAudioEndpointVolume.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_recorder_class_load() {
    let file_path =
        "./tests/data/audiostation/Audiostation/src/Classes/clsAudiostationRecorder.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_steamer_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsAudiostationSteamer.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_bass_time_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsBassTime.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_file_io_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsFileIo.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
#[ignore = "currently an error with sourcestream dealing with non-ascii character boundaries in utf-8 replacements"] // end byte index 7802 is not a char boundary; it is inside 'ä' (bytes 7801..7803)
fn audiostation_iaudio_meter_information_class_load() {
    let file_path =
        "./tests/data/audiostation/Audiostation/src/Classes/clsIAudioMeterInformation.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_local_storage_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsLocalStorage.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_logger_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsLogger.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_mp3_info_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsMp3Info.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_nest_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsNest.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_registry_settings_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsRegistrySettings.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_sibra_soft_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsSibraSoft.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_smart_buffer_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsSmartBuffer.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_string_builder_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsStringBuilder.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_web_client_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/clsWebClient.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn audiostation_volume_channel_class_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Classes/mdlVolumeChannel.cls";
    let class_bytes = std::fs::read(file_path).expect("Failed to read class file");

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}
