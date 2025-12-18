use vb6parse::parsers::ClassFile;
use vb6parse::SourceFile;

/// Example showing how to parse a VB6 class file from raw bytes.
/// This example uses a hardcoded byte array, but in a real application,
/// you would typically read the bytes from a `.cls` file on disk.
fn main() {
    // Hardcoded example of a VB6 class file content in bytes.
    let input = b"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = \"Something\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Some comment";

    // Decode the source file from the byte array.
    // The filename is provided for reference in error messages.
    // In a real application, use the actual filename.
    // Decode with replacement to handle any invalid characters gracefully.
    let result = SourceFile::decode_with_replacement("class_parse.cls", input);

    // Handle potential decoding errors.
    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'class_parse.cls': {e:?}"),
    };

    // Parse the class file from the decoded source file.
    let class = ClassFile::parse(&source_file).unwrap_or_fail();

    // Print out some information about the parsed class file.
    println!("Class Version Major: {}", class.header.version.major);
    println!("Class Version Minor: {}", class.header.version.minor);
    println!("Class Properties:");
    println!("\tMulti Use: {:?}", class.header.properties.multi_use);
    println!("\tPersistable: {:?}", class.header.properties.persistable);
    println!(
        "\tData Binding Behavior: {:?}",
        class.header.properties.data_binding_behavior
    );
    println!(
        "\tData Source Behavior: {:?}",
        class.header.properties.data_source_behavior
    );
    println!(
        "\tMTS Transaction Mode: {:?}",
        class.header.properties.mts_transaction_mode
    );
    println!("Class Attributes:");
    println!("\tName: {:?}", class.header.attributes.name);
    println!(
        "\tGlobal Name Space: {:?}",
        class.header.attributes.global_name_space
    );
    println!("\tCreatable: {:?}", class.header.attributes.creatable);
    println!(
        "\tPredeclared ID: {:?}",
        class.header.attributes.predeclared_id
    );
    println!("\tExposed: {:?}", class.header.attributes.exposed);
    println!("Ext Attributes:");
    for ext in &class.header.attributes.ext_key {
        println!("\t{} = {}", ext.0, ext.1);
    }

    // Print the concrete syntax tree (CST) of the parsed class file.
    println!("CST:");
    println!("{}", class.cst.debug_tree());
}
