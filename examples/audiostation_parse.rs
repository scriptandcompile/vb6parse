//! Example demonstrating how to read the audiostation project file
//! and parse it into a Concrete Syntax Tree (CST).

use std::env;
use vb6parse::SourceFile;
use vb6parse::parse;
use vb6parse::parsers::project::VB6Project;

fn main() -> Result<(), Box<dyn std::error::Error>>{
    let current_directory = env::current_dir()?;

    // Load the audiostation project file from the test/data directory
    let project_directory = current_directory.join("tests/data/audiostation/Audiostation");
    let project_path = project_directory.join("Audiostation.vbp");

    print!("Reading project file from: {:?}\n", project_path);

    let project_content = std::fs::read(project_path)
        .expect("Failed to read audiostation project file.");

    // Create a SourceFile for the project file.
    let project_source = SourceFile::decode_with_replacement(
        "audiostation.vbp",
        &project_content,
    ).expect("Failed to decode project source file.");

    // Parse the project file
    let project_parse_result = VB6Project::parse(&project_source);
    
    if project_parse_result.has_failures() {
        for failure in &project_parse_result.failures {
            failure.print();
        }
    }
    let project = project_parse_result.unwrap();

    println!("Project Title: {}", project.properties.title);
    println!("Startup Object: {}", project.properties.startup);
    println!("Number of Modules: {}", project.modules.len());
    println!("Number of Forms: {}", project.forms.len());
    println!("Number of Classes: {}", project.classes.len());

    println!();

    // For demonstration, parse the first module into a CST
    if let Some(first_module) = project.modules.last() {

        let module_path = project_directory.join(first_module.path.replace("\\", "/")).to_str().unwrap().to_string();
        print !("Reading module file from: {:?}\n", module_path);

        let module_source = SourceFile::decode_with_replacement(
            &module_path,
            &std::fs::read(&module_path)?,
        ).expect("Failed to decode module source file.");

        println!("Module Source:\n");
        print!("{}", module_source.get_source_stream().contents);

        let module_parse_result = vb6parse::parsers::module::VB6ModuleFile::parse(&module_source);
        if module_parse_result.has_failures() {
            for failure in &module_parse_result.failures {
                failure.print();
            }
        }
        let module_file = module_parse_result.unwrap();
        let cst = parse(module_file.tokens);
        
        println!("\nCST for module '{}':", module_path);
        println!("CST Root Kind: {:?}", cst.root_kind());
        println!("Number of children: {}", cst.child_count());
        println!("\nFull text of the CST:");
        println!("{}", cst.text());
        println!("\nDebug tree structure:");
        println!("{}", cst.debug_tree());
    }

    Ok(())
}