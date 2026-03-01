use vb6parse::*;

#[test]
fn chess_brain_vb_hashmap_class_load() {
    let class_bytes = include_bytes!("../data/ChessBrainVB/ChessBrainVB_V4_03a/Source/Modules/HashMap.cls");

    let result = SourceFile::decode_with_replacement("HashMap.cls", class_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'HashMap.cls': {e:?}"),
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
    settings.set_snapshot_path("../../snapshots/tests/class/chess_brain_vb");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(class);
}
