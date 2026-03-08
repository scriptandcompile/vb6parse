use vb6parse::*;

#[test]
fn cdiu_beat_up_editor_ccommondialog_class_load() {
    let class_bytes = include_bytes!("../data/CdiuBeatUpEditor/cCommonDialog.cls");

    let result = SourceFile::decode_with_replacement("cCommonDialog.cls", class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'cCommonDialog.cls': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn cdiu_beat_up_editor_clscryptoapiandcompression_class_load() {
    let class_bytes = include_bytes!("../data/CdiuBeatUpEditor/clsCryptoAPIandCompression.cls");

    let result = SourceFile::decode_with_replacement("clsCryptoAPIandCompression.cls", class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'clsCryptoAPIandCompression.cls': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn cdiu_beat_up_editor_gcommondialog_class_load() {
    let class_bytes = include_bytes!("../data/CdiuBeatUpEditor/GCommonDialog.cls");

    let result = SourceFile::decode_with_replacement("GCommonDialog.cls", class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'GCommonDialog.cls': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/cdiu_beat_up_editor");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}
