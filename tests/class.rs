use vb6parse::parsers::class::VB6ClassFile;

#[test]
fn artificial_life_organism_class_load() {
    let organism_class_bytes = include_bytes!("./data/vb6-code/Artificial-life/Organism.cls");

    let class_file = match VB6ClassFile::parse(
        "Organism.cls".to_owned(),
        &mut organism_class_bytes.as_slice(),
    ) {
        Ok(class_file) => class_file,
        Err(e) => {
            panic!("Failed to parse class file 'Organism.cls': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(class_file);
}

#[test]
fn blacklight_effect_class_load() {
    let class1_file_bytes = include_bytes!("./data/vb6-code/Blacklight-effect/FastDrawing.cls");

    let class1_file = match VB6ClassFile::parse(
        "FastDrawing.cls".to_owned(),
        &mut class1_file_bytes.as_slice(),
    ) {
        Ok(class1_file) => class1_file,
        Err(e) => {
            panic!("Failed to parse class file 'FastDrawing.cls': {}", e);
        }
    };

    let class2_file_bytes =
        include_bytes!("./data/vb6-code/Blacklight-effect/pdOpenSaveDialog.cls");

    let class2_file = match VB6ClassFile::parse(
        "pdOpenSaveDialog.cls".to_owned(),
        &mut class2_file_bytes.as_slice(),
    ) {
        Ok(class2_file) => class2_file,
        Err(e) => {
            panic!("Failed to parse class file 'pdOpenSaveDialog.cls': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(class1_file);
    insta::assert_yaml_snapshot!(class2_file);
}

#[test]
fn gradient_2d_class_load() {
    let class_file_bytes = include_bytes!("./data/vb6-code/Gradient-2D/cSystemColorDialog.cls");

    let class_file = match VB6ClassFile::parse(
        "cSystemColorDialog.cls".to_owned(),
        &mut class_file_bytes.as_slice(),
    ) {
        Ok(class_file) => class_file,
        Err(e) => {
            panic!(
                "Failed to parse class file 'cSystemColorDialog.cls' form : {}",
                e
            );
        }
    };

    insta::assert_yaml_snapshot!(class_file);
}

#[test]
fn hidden_markov_model_class_load() {
    let class_file_bytes = include_bytes!("./data/vb6-code/Hidden-Markov-model/cCommonDialog.cls");

    let class_file = match VB6ClassFile::parse(
        "cCommonDialog.cls".to_owned(),
        &mut class_file_bytes.as_slice(),
    ) {
        Ok(class_file) => class_file,
        Err(e) => {
            panic!(
                "Failed to parse class file 'cCommonDialog.cls' form : {}",
                e
            );
        }
    };

    insta::assert_yaml_snapshot!(class_file);
}
