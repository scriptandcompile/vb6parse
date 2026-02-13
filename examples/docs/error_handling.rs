use vb6parse::*;

fn main() {
    // Malformed VB6 code (missing End Sub, invalid syntax)
    let bad_code = r#"Attribute VB_Name = "BadModule"

Public Sub BrokenFunction()
    x = 5 + 
    ' Missing closing and the expression is incomplete

Public Sub AnotherFunction()
    MsgBox "This one is fine"
End Sub
"#;

    let source = SourceFile::from_string("BadModule.bas", bad_code);
    let result = ModuleFile::parse(&source);
    let (module, failures) = result.unpack();

    // We still might get a module!
    if let Some(module) = module {
        println!("✓ Parsed module: {}", module.name);
        println!("  (Despite {} errors)", failures.len());
    } else {
        println!("✗ Parsing completely failed");
    }

    // Always check and handle failures
    if !failures.is_empty() {
        println!("\nParsing Issues:");
        for failure in failures {
            println!(
                "  Line {}-{}: {:?}",
                failure.line_start, failure.line_end, failure.kind
            );

            // Print the error with source context
            failure.print();
        }
    }
}
