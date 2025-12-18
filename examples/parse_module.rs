use vb6parse::parsers::ModuleFile;
use vb6parse::SourceFile;

/// Example showing how to parse a VB6 module file from raw bytes.
/// This example uses a hardcoded byte array, but in a real application,
/// you would typically read the bytes from a `.bas` file on disk.
fn main() {
    // Hardcoded example of a VB6 module file content in bytes.
    let input = b"Attribute VB_Name = \"Module1\"

' Some comment

Sub HelloWorld()
    MsgBox \"Hello, World!\"
End Sub";

    // Decode the source file from the byte array.
    // The filename is provided for reference in error messages.
    // In a real application, use the actual filename.
    // Decode with replacement to handle any invalid characters gracefully.
    let result = SourceFile::decode_with_replacement("module_parse.bas", input);

    // Handle potential decoding errors.
    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'module_parse.bas': {e:?}"),
    };

    // Parse the module file from the decoded source file.
    let module = ModuleFile::parse(&source_file).unwrap_or_fail();

    // Print out some information about the parsed module file.
    println!("Module Name: {}", module.name);

    // Print the concrete syntax tree (CST) of the parsed module file.
    println!("CST:");
    println!("{}", module.cst.debug_tree());
}
