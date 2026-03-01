use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

#[test]
fn endless_runner_game_vb6_0_desert_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/desert.frm");

    let source_file = SourceFile::decode_with_replacement("desert.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'desert.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'desert.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_dialog_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Dialog.frm");

    let source_file = SourceFile::decode_with_replacement("Dialog.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Dialog.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Dialog.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form10_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form10.frm");

    let source_file = SourceFile::decode_with_replacement("Form10.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form10.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form10.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form11_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form11.frm");

    let source_file = SourceFile::decode_with_replacement("Form11.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form11.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form11.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form12_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form12.frm");

    let source_file = SourceFile::decode_with_replacement("Form12.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form12.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form12.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form2_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form2.frm");

    let source_file = SourceFile::decode_with_replacement("Form2.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form2.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form2.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form3_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form3.frm");

    let source_file = SourceFile::decode_with_replacement("Form3.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form3.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form3.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form5_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form5.frm");

    let source_file = SourceFile::decode_with_replacement("Form5.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form5.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form5.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form6_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form6.frm");

    let source_file = SourceFile::decode_with_replacement("Form6.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form6.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form6.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form7_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form7.frm");

    let source_file = SourceFile::decode_with_replacement("Form7.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form7.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form7.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form8_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form8.frm");

    let source_file = SourceFile::decode_with_replacement("Form8.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form8.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form8.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_form9_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Form9.frm");

    let source_file = SourceFile::decode_with_replacement("Form9.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Form9.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Form9.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_frmsplash_form_load() {
    let form_file_bytes = include_bytes!(
        "../data/Endless-runner-Game_VB6.0/Endless runner project files/frmSplash.frm"
    );

    let source_file = SourceFile::decode_with_replacement("frmSplash.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'frmSplash.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'frmSplash.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_howtoplay_form_load() {
    let form_file_bytes = include_bytes!(
        "../data/Endless-runner-Game_VB6.0/Endless runner project files/howtoplay.frm"
    );

    let source_file = SourceFile::decode_with_replacement("howtoplay.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'howtoplay.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'howtoplay.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_jp2_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/jp2.frm");

    let source_file = SourceFile::decode_with_replacement("jp2.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'jp2.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'jp2.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_jump4_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/jump4.frm");

    let source_file = SourceFile::decode_with_replacement("jump4.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'jump4.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'jump4.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_jump_king_form_load() {
    let form_file_bytes = include_bytes!(
        "../data/Endless-runner-Game_VB6.0/Endless runner project files/jump king.frm"
    );

    let source_file = SourceFile::decode_with_replacement("jump king.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'jump king.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'jump king.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_leaderboard_form_load() {
    let form_file_bytes = include_bytes!(
        "../data/Endless-runner-Game_VB6.0/Endless runner project files/Leaderboard.frm"
    );

    let source_file = SourceFile::decode_with_replacement("Leaderboard.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Leaderboard.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Leaderboard.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn endless_runner_game_vb6_0_report_form_load() {
    let form_file_bytes =
        include_bytes!("../data/Endless-runner-Game_VB6.0/Endless runner project files/Report.frm");

    let source_file = SourceFile::decode_with_replacement("Report.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Report.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Report.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/endless_runner_game_vb6_0");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}
