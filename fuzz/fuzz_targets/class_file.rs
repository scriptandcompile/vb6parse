#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::{ClassFile, SourceFile};

fuzz_target!(|data: &[u8]| {
    if let Ok(source_file) = SourceFile::decode_with_replacement("fuzz.cls", data) {
        let result = ClassFile::parse(&source_file);

        // Unpack and check failures
        let (class_opt, failures) = result.unpack();

        // Exercise all failure fields
        for failure in failures {
            let _ = &failure.kind;
            let _ = failure.error_offset;
            let _ = failure.line_start;
            let _ = failure.line_end;
        }

        // If we got a class file, validate its structure
        if let Some(class) = class_opt {
            // Exercise header properties
            let _ = &class.header.version;
            let _ = &class.header.properties;
            let _ = &class.header.attributes;

            // Exercise class-specific properties
            let _ = class.header.properties.multi_use;
            let _ = class.header.properties.persistable;
            let _ = class.header.properties.data_binding_behavior;
            let _ = class.header.properties.data_source_behavior;
            let _ = class.header.properties.mts_transaction_mode;

            // Exercise attributes
            let _ = &class.header.attributes.name;
            let _ = &class.header.attributes.global_name_space;
            let _ = &class.header.attributes.creatable;
            let _ = &class.header.attributes.predeclared_id;
            let _ = &class.header.attributes.exposed;

            // Exercise CST
            let serializable = class.cst.to_serializable();
            let _ = &serializable.root.kind;
            let _ = serializable.root.children.len();
        }
    }
});
