use vb6parse::*;

#[test]
fn artificial_life_organism_class_load() {
    let file_path = "./tests/data/vb6-code/Artificial-life/Organism.cls";
    let class_bytes = std::fs::read(file_path).unwrap();

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let result = ClassFile::parse(&source_file);

    if result.has_failures() {
        for failure in result.failures() {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class = result.unwrap();

    insta::assert_yaml_snapshot!(class);
}

#[test]
fn blacklight_effect_class_load() {
    let file_path1 = "./tests/data/vb6-code/Blacklight-effect/FastDrawing.cls";
    let class_bytes1 = std::fs::read(file_path1).unwrap();

    let result1 = SourceFile::decode_with_replacement(file_path1, &class_bytes1);

    let source_file1 = match result1 {
        Ok(source_file1) => source_file1,
        Err(e) => panic!("Failed to decode source file '{file_path1}': {e:?}"),
    };

    let result1 = ClassFile::parse(&source_file1);

    if result1.has_failures() {
        for failure in result1.failures() {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class1 = result1.unwrap();

    insta::assert_yaml_snapshot!(class1);

    let file_path2 = "./tests/data/vb6-code/Blacklight-effect/pdOpenSaveDialog.cls";
    let class_bytes2 = std::fs::read(file_path2).unwrap();

    let result2 = SourceFile::decode_with_replacement(file_path2, &class_bytes2);

    let source_file2 = match result2 {
        Ok(source_file2) => source_file2,
        Err(e) => panic!("Failed to decode source file '{file_path2}': {e:?}"),
    };

    let result2 = ClassFile::parse(&source_file2);

    if result2.has_failures() {
        for failure in result2.failures() {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class2 = result2.unwrap();

    insta::assert_yaml_snapshot!(class2);
}

#[test]
fn gradient_2d_class_load() {
    let file_path = "./tests/data/vb6-code/Gradient-2D/cSystemColorDialog.cls";
    let class_bytes = std::fs::read(file_path).unwrap();

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let result = ClassFile::parse(&source_file);

    if result.has_failures() {
        for failure in result.failures() {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class = result.unwrap();

    insta::assert_yaml_snapshot!(class);
}

#[test]
fn hidden_markov_model_class_load() {
    let file_path = "./tests/data/vb6-code/Hidden-Markov-model/cCommonDialog.cls";
    let class_bytes = std::fs::read(file_path).unwrap();

    let result = SourceFile::decode_with_replacement(file_path, &class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let result = ClassFile::parse(&source_file);

    if result.has_failures() {
        for failure in result.failures() {
            failure.print();
        }

        panic!("Class parse had failures");
    }

    let class = result.unwrap();

    insta::assert_yaml_snapshot!(class);
}
