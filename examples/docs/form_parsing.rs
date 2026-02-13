use vb6parse::language::ControlKind;
use vb6parse::*;

fn main() {
    let form_content = r#"VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "My Application"
   ClientHeight    =   3090
   ClientWidth     =   4560
   Begin VB.CommandButton btnSubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "MainForm"

Private Sub btnSubmit_Click()
    MsgBox "Button clicked!"
End Sub
"#;

    let source = SourceFile::from_string("MainForm.frm", form_content);
    let result = FormFile::parse(&source);
    let (form_opt, _failures) = result.unpack();

    let Some(form_file) = form_opt else {
        println!("Failed to parse the form file.");
        return;
    };

    // Access the root form control
    match form_file.form.kind() {
        ControlKind::Form {
            controls: _,
            menus: _,
            properties,
        } => {
            println!("Form: {}", form_file.attributes.name);
            println!("  Caption: {}", properties.caption);
            println!(
                "  Size: {}x{}",
                properties.client_width, properties.client_height
            );
        }
        ControlKind::MDIForm {
            properties,
            controls: _,
            menus: _,
        } => {
            println!("MDI Form: {}", form_file.attributes.name);
            println!("  Caption: {}", properties.caption);
            println!("  Size: {}x{}", properties.width, properties.height);
        }
        _ => {
            println!("Only Form and MDIForm are valid top level controls in a form file.");
            return;
        }
    }

    // Iterate child controls
    println!("\n  Controls:");
    if let Some(children) = form_file.form.children() {
        for child in children {
            println!("      - {} ({})", child.name(), child.kind());
        }
    };
}
