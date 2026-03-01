use image::EncodableLayout;
use vb6parse::files::ModuleFile;
use vb6parse::io::SourceFile;

#[test]
fn audiostation_mod_args_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modArgs.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_audiostation_cd_player_module_load() {
    let file_path =
        "./tests/data/audiostation/Audiostation/src/Modules/modAudiostationCDPlayer.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_audiostation_midi_player_module_load() {
    let file_path =
        "./tests/data/audiostation/Audiostation/src/Modules/modAudiostationMIDIPlayer.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_audiostation_mp3_player_module_load() {
    let file_path =
        "./tests/data/audiostation/Audiostation/src/Modules/modAudiostationMP3Player.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_bass_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modBass.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_bass_cd_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modBassCD.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_convert_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modConvert.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_enums_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modEnums.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_id3_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modID3.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_language_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modLanguage.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_main_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modMain.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_midi_const_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modMidiConst.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_midi_utils_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modMidiUtils.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_mus_player_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modMusPlayer.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_net_radio_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modNetRadio.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_os_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modOS.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_playlist_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modPlaylist.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_settings_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modSettings.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_sid_player_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modSidPlayer.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_spectrum_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modSpectrum.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn audiostation_mod_volume_module_load() {
    let file_path = "./tests/data/audiostation/Audiostation/src/Modules/modVolume.bas";
    let module_file_bytes = std::fs::read(file_path).expect("Failed to read module file");

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let (module_file_opt, failures) = result.unpack();
    let Some(module_file) = module_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse '{file_path}' module file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/module/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(module_file);
}
