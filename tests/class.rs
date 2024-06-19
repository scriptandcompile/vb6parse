use vb6parse::class::VB6ClassFile;
use winnow::error::ErrMode;

#[test]
fn artificial_life_organism_class_load() {
    let organism_class_bytes = include_bytes!("./data/vb6-code/Artificial-life/Organism.cls");

    let organism_class_result = VB6ClassFile::parse(organism_class_bytes);

    match organism_class_result {
        Ok(organism_class) => {
            assert_eq!(organism_class.header.version.major, 1);
            assert_eq!(organism_class.header.version.minor, 0);

            assert_eq!(organism_class.header.multi_use, true);
            assert_eq!(organism_class.header.persistable, false);
            assert_eq!(organism_class.header.data_binding_behavior, false);
            assert_eq!(organism_class.header.data_source_behavior, false);
            assert_eq!(organism_class.header.mts_transaction_mode, false);

            assert_eq!(organism_class.header.attributes.name, b"Organism");
            assert_eq!(organism_class.header.attributes.global_name_space, false);
            assert_eq!(organism_class.header.attributes.creatable, true);
            assert_eq!(organism_class.header.attributes.pre_declared_id, false);
            assert_eq!(organism_class.header.attributes.exposed, false);
        }
        Err(e) => match e {
            ErrMode::Backtrack(e) | ErrMode::Cut(e) => {
                println!("{:?}", e);
                panic!("Failed to parse organism class.");
            }
            ErrMode::Incomplete(_) => {
                panic!("Incomplete parse of organism class.");
            }
        },
    }
}
