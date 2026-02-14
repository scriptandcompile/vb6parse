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
    match &form_file.form {
        vb6parse::language::FormRoot::Form(form) => {
            println!("Form: {}", form_file.attributes.name);
            println!("  Caption: {}", form.properties.caption);
            println!(
                "  Size: {}x{}",
                form.properties.client_width, form.properties.client_height
            );

            // Iterate child controls
            println!("\n  Controls:");
            for child in &form.controls {
                println!("      - {} ({})", child.name(), child.kind());
            }
        }
        vb6parse::language::FormRoot::MDIForm(mdi_form) => {
            println!("MDI Form: {}", form_file.attributes.name);
            println!("  Caption: {}", mdi_form.properties.caption);
            println!(
                "  Size: {}x{}",
                mdi_form.properties.width, mdi_form.properties.height
            );

            // Iterate child controls
            println!("\n  Controls:");
            for child in &mdi_form.controls {
                println!("      - {} ({})", child.name(), child.kind());
            }
        }
    }
}
