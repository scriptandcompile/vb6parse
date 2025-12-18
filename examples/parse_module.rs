use vb6parse::parsers::ModuleFile;
use vb6parse::SourceFile;

/// Example showing how to parse a VB6 module file from raw bytes.
/// This example uses a hardcoded byte array, but in a real application,
/// you would typically read the bytes from a `.bas` file on disk.
fn main() {
    // Hardcoded example of a VB6 module file content in bytes.
    let input = "Attribute VB_Name = \"Module1\"

' Some comment

Sub HelloWorld()
    MsgBox \"Hello, World!\"
End Sub";

    // Create a SourceFile from the input string.
    let source_file = SourceFile::from_string("module_parse.bas", input);

    // Parse the module file from the decoded source file.
    let module = ModuleFile::parse(&source_file).unwrap_or_fail();

    // Print out some information about the parsed module file.
    println!("Module Name: {}", module.name);

    // Print the concrete syntax tree (CST) of the parsed module file.
    println!("CST:");
    println!("{}", module.cst.debug_tree());
}
