//! Example showing how to parse a VB6 form file from raw bytes.
//! This example uses a hardcoded byte array, but in a real application,
//! you would typically read the bytes from a `.frm` file on disk.
//!
//! This is a minimal example with just a couple of controls.
//!

use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

fn main() {
    let input = b"VERSION 5.00
Begin VB.Form Form1
    BorderStyle     =   1  'Fixed Single
    Caption         =   \"Form1\"
    ClientHeight    =   3195
    ClientLeft      =   60
    ClientTop       =   345
    ClientWidth     =   4680
    LinkTopic       =   \"\"
    MaxButton       =   1   'True
    MinButton       =   1   'True
    ScaleHeight     =   3195
    ScaleWidth      =   4680
    StartUpPosition =   3  'Windows Default
    BeginProperty Font
        Name            =   \"MS Sans Serif\"
        Size            =   8.25
        Charset         =   0
        Weight          =   400
        Underline       =   0   'False
        Italic          =   0   'False
        Strikethrough   =   0   'False
    EndProperty
    Begin VB.TextBox Text1
        Height          =   315
        Left            =   120
        TabIndex        =   0
        Top             =   720
        Width           =   1215
    End
    Begin VB.CommandButton Command1
        Caption         =   \"Command1\"
        Height          =   375
        Left            =   1560
        TabIndex        =   1
        Top             =   720
        Width           =   1215
    End
End
Attribute VB_Name = \"Form1\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Some comment";

    // Decode the source file from the byte array.
    // The filename is provided for reference in error messages.
    // In a real application, use the actual filename.
    // Decode with replacement to handle any invalid characters gracefully.
    let result = SourceFile::decode_with_replacement("form_parse.frm", input);

    // Handle potential decoding errors.
    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'form_parse.frm': {e:?}"),
    };

    // Parse the form file from the decoded source file.
    let form_file = FormFile::parse(&source_file).unwrap_or_fail();

    // Print out some information about the parsed form file.
    println!("Form Version Major: {}", form_file.version.major);
    println!("Form Version Minor: {}", form_file.version.minor);
    println!("Form Properties:");

    // Extract the properties, controls, and menus from the form root.
    let vb6parse::language::FormRoot::Form(form) = &form_file.form else {
        panic!("Unexpected form kind - expected Form, got MDIForm");
    };

    // Print out the form properties.
    println!("\tCaption: {:?}", form.properties.caption);
    println!("\tBorder Style: {:?}", form.properties.border_style);
    println!("\tMax Button: {:?}", form.properties.max_button);
    println!("\tMin Button: {:?}", form.properties.min_button);
    println!(
        "\tStartup Position: {:?}",
        form.properties.start_up_position
    );
    println!("\tClient Height: {:?}", form.properties.client_height);
    println!("\tClient Width: {:?}", form.properties.client_width);
    println!("\tLink Topic: {:?}", form.properties.link_topic);
    println!("\tNumber of Controls: {}", form.controls.len());
    println!("\tNumber of Menus: {}", form.menus.len());
    println!("Form Attributes:");
    println!("\tName: {:?}", form_file.attributes.name);
    println!(
        "\tGlobal Name Space: {:?}",
        form_file.attributes.global_name_space
    );
    println!("\tCreatable: {:?}", form_file.attributes.creatable);
    println!(
        "\tPredeclared ID: {:?}",
        form_file.attributes.predeclared_id
    );
    println!("\tExposed: {:?}", form_file.attributes.exposed);
    println!("Ext Attributes:");
    for ext in &form_file.attributes.ext_key {
        println!("\t{} = {}", ext.0, ext.1);
    }

    // Print the concrete syntax tree (CST) of the parsed form file.
    println!("CST:");
    println!("{}", form_file.cst.debug_tree());
}
