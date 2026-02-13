use vb6parse::*;

fn main() {
    // VB6 code to parse
    let code = r#"Attribute VB_Name = "HelloWorld"

Public Sub SayHello()
    MsgBox "Hello from VB6Parse!"
End Sub
"#;

    // Step 1: Create a SourceFile (handles encoding)
    let source = SourceFile::from_string("HelloWorld.bas", code);

    // Step 2: Parse as a module
    let result = ModuleFile::parse(&source);

    // Step 3: Unpack the result (separates output from errors)
    let (module_opt, failures) = result.unpack();

    // Step 4: Handle the result
    if let Some(module) = module_opt {
        println!("✓ Successfully parsed module: {}", module.name);
        println!("  Version: {}", module.version);
        println!("  Has code: {}", module.cst.debug_tree());
    }

    // Step 5: Display any parsing errors
    if !failures.is_empty() {
        println!("\n⚠ Encountered {} parsing issues:", failures.len());
        for failure in failures {
            failure.print();
        }
    }
}
