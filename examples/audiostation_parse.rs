//! Example demonstrating how to read the audiostation project file
//! and parse it into a Concrete Syntax Tree (CST).
//!

use std::env;
use vb6parse::files::ProjectFile;
use vb6parse::io::SourceFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let current_directory = env::current_dir()?;

    let project_directory = current_directory.join("tests/data/audiostation/Audiostation");
    let project_path = project_directory.join("Audiostation.vbp");

    println!("Reading project file from: {}", project_path.display());

    // Load the audiostation project file from the test/data directory
    let project_content =
        std::fs::read(&project_path).expect("Failed to read audiostation project file. This example requires the test data to be present. run `git submodule update --init --recursive` in the terminal from the repository root to fetch the test data.");

    // Create a SourceFile for the project file.
    let project_source = SourceFile::decode_with_replacement("audiostation.vbp", &project_content)
        .expect("Failed to decode project source file.");

    // Parse the project file
    let project_parse_result = ProjectFile::parse(&project_source);

    let (project_opt, failures) = project_parse_result.unpack();

    if !failures.is_empty() {
        eprintln!("Failure while parsing project file. Errors:");
        for failure in failures {
            failure.print();
        }
    }

    let project = project_opt.expect("Project should be present.");

    println!("Project Title: {}", project.properties.title);
    println!("Startup Object: {}", project.properties.startup);
    println!(
        "Number of Modules: {}",
        project.modules().collect::<Vec<_>>().len()
    );
    println!(
        "Number of Forms: {}",
        project.forms().collect::<Vec<_>>().len()
    );
    println!(
        "Number of Classes: {}",
        project.classes().collect::<Vec<_>>().len()
    );

    println!();

    // For demonstration, parse the last module into a CST
    if let Some(last_module) = project.modules().last() {
        let module_path = project_directory
            .join(last_module.path.replace('\\', "/"))
            .to_str()
            .unwrap()
            .to_string();
        println!("Reading module file from: {module_path:?}");

        let module_source =
            SourceFile::decode_with_replacement(&module_path, &std::fs::read(&module_path)?)
                .expect("Failed to decode module source file.");

        let module_parse_result = vb6parse::ModuleFile::parse(&module_source);

        let (module_file_opt, failures) = module_parse_result.unpack();

        if !failures.is_empty() {
            eprintln!("Failure while parsing module file. Errors:");
            for failure in failures {
                failure.print();
            }
        }

        let module_file = module_file_opt.expect("Module should be present.");

        let cst = &module_file.cst;

        println!("\nCST for module '{module_path}':");
        println!("CST Root Kind: {:?}", cst.root_kind());
        println!("Number of children: {}", cst.child_count());
        println!("\nFull text of the CST:");
        println!("{}", cst.text());
        println!("\nDebug tree structure:");
        println!("{}", cst.debug_tree());
    }

    Ok(())
}
