#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::{ModuleFile, SourceFile};

fuzz_target!(|data: &[u8]| {
    if let Ok(source_file) = SourceFile::decode_with_replacement("fuzz.bas", data) {
        let result = ModuleFile::parse(&source_file);

        // Unpack and check failures
        let (module_opt, failures) = result.unpack();

        // Exercise all failure fields
        for failure in failures {
            let _ = &failure.kind;
            let _ = failure.error_offset;
            let _ = failure.line_start;
            let _ = failure.line_end;
        }

        // If we got a module file, validate its structure
        if let Some(module) = module_opt {
            // Exercise module name
            let _ = module.name.as_str();
            let _ = module.name.len();

            // Exercise CST
            let serializable = module.cst.to_serializable();
            let _ = &serializable.root.kind;
            let _ = serializable.root.children.len();

            // Walk through CST children
            for child in &serializable.root.children {
                let _ = &child.kind;
                let _ = &child.text;
            }
        }
    }
});
