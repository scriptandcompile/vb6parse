use vb6parse::parsers::class::VB6ClassFile;

#[test]
fn artificial_life_organism_class_load() {
    let file_name = "Organism.cls".to_owned();

    let organism_class_bytes = include_bytes!("./data/vb6-code/Artificial-life/Organism.cls");

    let class_file = match VB6ClassFile::parse(file_name, &mut organism_class_bytes.as_slice()) {
        Ok(class_file) => class_file,
        Err(e) => {
            panic!("Failed to parse class file 'Organism': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(class_file);
}
