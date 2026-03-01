use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

// ========================================

#[test]
fn audiostation_about_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_About.frm");

    let source_file = SourceFile::decode_with_replacement("Form_About.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_About.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_About.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_busy_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Busy.frm");

    let source_file = SourceFile::decode_with_replacement("Form_Busy.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Busy.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Busy.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_init_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Init.frm");

    let source_file = SourceFile::decode_with_replacement("Form_Init.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Init.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Init.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_main_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Main.frm");

    let source_file = SourceFile::decode_with_replacement("Form_Main.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Main.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Main.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_midi_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Midi.frm");

    let source_file = SourceFile::decode_with_replacement("Form_Midi.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Midi.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Midi.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_normalize_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Normalize.frm");

    let source_file = SourceFile::decode_with_replacement("Form_Normalize.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Normalize.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Normalize.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_open_dialog_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_OpenDialog.frm");

    let source_file = SourceFile::decode_with_replacement("Form_OpenDialog.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_OpenDialog.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_OpenDialog.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_playlist_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Playlist.frm");

    let source_file = SourceFile::decode_with_replacement("Form_Playlist.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Playlist.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Playlist.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_plugins_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Plugins.frm");

    let source_file = SourceFile::decode_with_replacement("Form_Plugins.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Plugins.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Plugins.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_settings_recorder_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Settings_Recorder.frm");

    let source_file =
        SourceFile::decode_with_replacement("Form_Settings_Recorder.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Settings_Recorder.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Settings_Recorder.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_settings_record_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Settings_Record.frm");

    let source_file =
        SourceFile::decode_with_replacement("Form_Settings_Record.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Settings_Record.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Settings_Record.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_streams_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Streams.frm");

    let source_file = SourceFile::decode_with_replacement("Form_Streams.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Streams.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Streams.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_system_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_System.frm");

    let source_file = SourceFile::decode_with_replacement("Form_System.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_System.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_System.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn audiostation_track_properties_form_load() {
    let form_file_bytes =
        include_bytes!("../data/audiostation/Audiostation/src/Forms/Form_Track_Properties.frm");

    let source_file =
        SourceFile::decode_with_replacement("Form_Track_Properties.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Form_Track_Properties.frm' source file");
        }
    };

    let (form_file_opt, failures) = FormFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form_Track_Properties.frm' form file");
    }

    let form_file = form_file_opt.expect("Form should be present.");
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/audiostation");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}
