#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::{FormFile, SourceFile};

fuzz_target!(|data: &[u8]| {
    if let Ok(source_file) = SourceFile::decode_with_replacement("fuzz.frm", data) {
        let result = FormFile::parse(&source_file);

        // Unpack and check failures
        let (form_opt, failures) = result.unpack();

        // Exercise all failure fields
        for failure in failures {
            let _ = &failure.kind;
            let _ = failure.error_offset;
            let _ = failure.line_start;
            let _ = failure.line_end;
        }

        // If we got a form file, validate its structure
        if let Some(form) = form_opt {
            // Exercise version and attributes
            let _ = &form.version;
            let _ = &form.attributes.name;
            let _ = &form.attributes.global_name_space;
            let _ = &form.attributes.creatable;
            let _ = &form.attributes.predeclared_id;
            let _ = &form.attributes.exposed;
            
            // Exercise objects
            for obj in &form.objects {
                match obj {
                    vb6parse::files::common::ObjectReference::Compiled {
                        uuid,
                        version,
                        unknown1,
                        file_name,
                    } => {
                        let _ = uuid;
                        let _ = version;
                        let _ = unknown1;
                        let _ = file_name;
                    }
                    vb6parse::files::common::ObjectReference::Project { path } => {
                        let _ = path;
                    }
                }
            }
            
            // Exercise form control hierarchy
            let _ = form.form.kind();
            let _ = form.form.name();
            let _ = form.form.tag();
            let _ = form.form.index();
            
            // Walk through child controls recursively
            fn walk_controls(control: &vb6parse::language::Control) {
                let _ = control.kind();
                let _ = control.name();
                let _ = control.tag();
                let _ = control.index();
                
                if let Some(children) = control.children() {
                    for child in children {
                        walk_controls(child);
                    }
                }
            }
            walk_controls(&form.form);
            
            // Exercise CST
            let serializable = form.cst.to_serializable();
            let _ = &serializable.root.kind;
            let _ = serializable.root.children.len();
        }
    }
});
