#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::{ProjectFile, SourceFile};

fuzz_target!(|data: &[u8]| {
    if let Ok(source_file) = SourceFile::decode_with_replacement("fuzz.vbp", data) {
        let result = ProjectFile::parse(&source_file);

        // Unpack and check failures
        let (project_opt, failures) = result.unpack();

        // Exercise all failure fields
        for failure in failures {
            let _ = &failure.kind;
            let _ = failure.error_offset;
            let _ = failure.line_start;
            let _ = failure.line_end;
        }

        // If we got a project, validate its structure
        if let Some(project) = project_opt {
            // Test project properties (they are &str fields)
            let _ = project.project_type;
            let _ = project.properties.startup;
            let _ = project.properties.name;
            let _ = project.properties.exe_32_file_name;
            let _ = project.properties.title;

            // Iterate through all modules (fields, not methods)
            for module in project.modules() {
                let _ = module.name;
                let _ = module.path;
            }

            // Iterate through all forms (just &str)
            for form in project.forms() {
                let _ = form.len();
            }

            // Iterate through all classes (fields, not methods)
            for class in project.classes() {
                let _ = class.name;
                let _ = class.path;
            }

            // Iterate through all references
            for reference in project.references() {
                match reference {
                    vb6parse::files::project::ProjectReference::Compiled { uuid, path, .. } => {
                        let _ = uuid;
                        let _ = path;
                    }
                    vb6parse::files::project::ProjectReference::SubProject { path } => {
                        let _ = path;
                    }
                }
            }

            // Test user controls, documents, and designers (just &str)
            for control in project.user_controls() {
                let _ = control.len();
            }

            for doc in project.user_documents() {
                let _ = doc.len();
            }

            for designer in project.designers() {
                let _ = designer.len();
            }
        }
    }
});
