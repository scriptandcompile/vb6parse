//! Example demonstrating control-only parsing of VB6 Form files.
//!
//! This example shows how to use the fast-path `parse_control_only()` API
//! to parse only the VERSION statement and control hierarchy from a Form file,
//! without parsing the code section or creating a full CST.
//!
//! This is useful for scenarios that only need UI information (control hierarchy,
//! properties) and don't need the code implementation.

use vb6parse::{tokenize, FormFile, SourceFile};

fn main() {
    // Example Form file content
    let form_content = br#"VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Example Form"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1
      Caption         =   "Click Me"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1
      Caption         =   "Enter text:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is where the code would start, but we won't parse it
Private Sub Command1_Click()
    MsgBox "Hello, World!"
End Sub
"#;

    println!("=== Control-Only Form Parsing Example ===\n");

    // Decode the file with Windows-1252 encoding
    let source_file = match SourceFile::decode_with_replacement("example.frm", form_content) {
        Ok(sf) => sf,
        Err(e) => {
            eprintln!("Failed to decode source file: {e:?}");
            return;
        }
    };

    // Tokenize the source
    let mut source_stream = source_file.source_stream();
    let parse_result = tokenize(&mut source_stream);

    // Unpack the tokenization result
    let (token_stream_opt, tok_failures) = parse_result.unpack();

    if !tok_failures.is_empty() {
        println!("Tokenization failures encountered:");
        for failure in &tok_failures {
            println!("  - {failure:?}");
        }
        println!();
    }

    let Some(token_stream) = token_stream_opt else {
        eprintln!("Tokenization failed");
        return;
    };

    // Parse only the control (fast path - skips code section)
    let result = FormFile::parse_control_only(token_stream);

    // Unpack the result to get the parsed data and any failures
    let (parse_result, failures) = result.unpack();

    // Report any parsing failures
    if !failures.is_empty() {
        println!("Parsing failures encountered:");
        for failure in &failures {
            println!("  - {failure:?}");
        }
        println!();
    }

    // Process the parsed control
    match parse_result {
        Some((version, control_opt, remaining_stream)) => {
            // Display version
            if let Some(v) = version {
                println!("Version: {}.{:02}", v.major, v.minor);
            } else {
                println!("Version: Not found");
            }

            println!();

            // Display control information
            if let Some(control) = control_opt {
                println!("Form name: {}", control.name());

                // Display form properties
                if let vb6parse::language::ControlKind::Form {
                    properties,
                    controls,
                    ..
                } = control.kind()
                {
                    println!("Form caption: {:?}", properties.caption);
                    println!(
                        "Form dimensions: {}x{}",
                        properties.width, properties.height
                    );
                    println!("Number of child controls: {}", controls.len());

                    println!("\nChild controls:");
                    for child in controls {
                        print_control_info(&child, 1);
                    }
                } else {
                    println!("Warning: Parsed control is not a Form");
                }
            } else {
                println!("Warning: No control was parsed");
            }

            println!();

            // Show information about remaining tokens
            let remaining_tokens: Vec<_> = remaining_stream.clone().collect();
            println!(
                "Remaining tokens after control: {} tokens",
                remaining_tokens.len()
            );
            if !remaining_tokens.is_empty() {
                println!("First remaining token: {:?}", remaining_tokens[0]);
            }
        }
        None => {
            println!("Failed to parse control");
        }
    }

    println!("\n=== Parsing Complete ===");
}

/// Helper function to print control information with indentation
fn print_control_info(control: &vb6parse::language::Control, indent_level: usize) {
    let indent = "  ".repeat(indent_level);

    println!(
        "{}- {} ({})",
        indent,
        control.name(),
        control_type_name(control.kind())
    );

    // If this control has children, print them recursively
    if let Some(children) = get_child_controls(control.kind()) {
        for child in children {
            print_control_info(child, indent_level + 1);
        }
    }
}

/// Get the type name of a control
fn control_type_name(kind: &vb6parse::language::ControlKind) -> &'static str {
    match kind {
        vb6parse::language::ControlKind::Form { .. } => "Form",
        vb6parse::language::ControlKind::CommandButton { .. } => "CommandButton",
        vb6parse::language::ControlKind::TextBox { .. } => "TextBox",
        vb6parse::language::ControlKind::Label { .. } => "Label",
        vb6parse::language::ControlKind::CheckBox { .. } => "CheckBox",
        vb6parse::language::ControlKind::ComboBox { .. } => "ComboBox",
        vb6parse::language::ControlKind::Frame { .. } => "Frame",
        vb6parse::language::ControlKind::ListBox { .. } => "ListBox",
        vb6parse::language::ControlKind::OptionButton { .. } => "OptionButton",
        vb6parse::language::ControlKind::PictureBox { .. } => "PictureBox",
        vb6parse::language::ControlKind::VScrollBar { .. } => "VScrollBar",
        vb6parse::language::ControlKind::HScrollBar { .. } => "HScrollBar",
        vb6parse::language::ControlKind::Timer { .. } => "Timer",
        vb6parse::language::ControlKind::Image { .. } => "Image",
        vb6parse::language::ControlKind::Shape { .. } => "Shape",
        vb6parse::language::ControlKind::Line { .. } => "Line",
        vb6parse::language::ControlKind::Menu { .. } => "Menu",
        vb6parse::language::ControlKind::MDIForm { .. } => "MDIForm",
        vb6parse::language::ControlKind::Data { .. } => "Data",
        vb6parse::language::ControlKind::FileListBox { .. } => "FileListBox",
        vb6parse::language::ControlKind::DriveListBox { .. } => "DriveListBox",
        vb6parse::language::ControlKind::DirListBox { .. } => "DirListBox",
        vb6parse::language::ControlKind::Ole { .. } => "OLE",
        vb6parse::language::ControlKind::Custom { .. } => "Custom",
    }
}

/// Get child controls from a control kind, if any
fn get_child_controls(
    kind: &vb6parse::language::ControlKind,
) -> Option<&Vec<vb6parse::language::Control>> {
    match kind {
        vb6parse::language::ControlKind::Form { controls, .. }
        | vb6parse::language::ControlKind::Frame { controls, .. }
        | vb6parse::language::ControlKind::PictureBox { controls, .. } => Some(controls),
        _ => None,
    }
}
