use vb6parse::*;

#[test]
fn artificial_life_organism_class_load() {
    let file_path = "./tests/data/vb6-code/Artificial-life/Organism.cls";
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
    settings.set_snapshot_path("../snapshots/tests/class");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn blacklight_effect_class_load() {
    let file_path1 = "./tests/data/vb6-code/Blacklight-effect/FastDrawing.cls";
    let class_bytes1 = std::fs::read(file_path1).expect("Failed to read class file");

    let result1 = SourceFile::decode_with_replacement(file_path1, &class_bytes1);

    let source_file1 = match result1 {
        Ok(source_file1) => source_file1,
        Err(e) => panic!("Failed to decode source file '{file_path1}': {e:?}"),
    };

    let (class_file_opt1, failures1) = ClassFile::parse(&source_file1).unpack();

    if !failures1.is_empty() {
        for failure in failures1 {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class1 = class_file_opt1.expect("Class should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/class");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class1);

    let file_path2 = "./tests/data/vb6-code/Blacklight-effect/pdOpenSaveDialog.cls";
    let class_bytes2 = std::fs::read(file_path2).expect("Failed to read class file");

    let result2 = SourceFile::decode_with_replacement(file_path2, &class_bytes2);

    let source_file2 = match result2 {
        Ok(source_file2) => source_file2,
        Err(e) => panic!("Failed to decode source file '{file_path2}': {e:?}"),
    };

    let (class_file_opt2, failures2) = ClassFile::parse(&source_file2).unpack();

    if !failures2.is_empty() {
        for failure in failures2 {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class2 = class_file_opt2.expect("Class should be present.");

    insta::assert_yaml_snapshot!(class2);
}

#[test]
fn gradient_2d_class_load() {
    let file_path = "./tests/data/vb6-code/Gradient-2D/cSystemColorDialog.cls";
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
    settings.set_snapshot_path("../snapshots/tests/class");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}

#[test]
fn hidden_markov_model_class_load() {
    let file_path = "./tests/data/vb6-code/Hidden-Markov-model/cCommonDialog.cls";
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
    settings.set_snapshot_path("../snapshots/tests/class");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}
