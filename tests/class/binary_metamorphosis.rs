use vb6parse::*;

#[test]
fn binary_metamorphosis_v1_common_dialog_class_load() {
    let file_path =
        "./tests/data/Binary-metamorphosis/Binary metamorphosis V1.0/src/cCommonDialog.cls";
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
    settings.set_snapshot_path("../../snapshots/tests/class/binary_metamorphosis");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn binary_metamorphosis_v2_common_dialog_class_load() {
    let file_path =
        "./tests/data/Binary-metamorphosis/Binary metamorphosis V2.0/src/cCommonDialog.cls";
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
    settings.set_snapshot_path("../../snapshots/tests/class/binary_metamorphosis");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn binary_metamorphosis_v3_common_dialog_class_load() {
    let file_path =
        "./tests/data/Binary-metamorphosis/Binary metamorphosis V3.0/src/cCommonDialog.cls";
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
    settings.set_snapshot_path("../../snapshots/tests/class/binary_metamorphosis");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}
