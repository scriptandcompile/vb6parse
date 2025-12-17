//! # `SavePicture` Statement
//!
//! Saves a graphical image from a control or form to a file.
//!
//! ## Syntax
//!
//! ```vb
//! SavePicture picture, stringexpression
//! ```
//!
//! ## Parts
//!
//! - **picture**: Required. A property or graphic object from which to save the image. The image
//!   can be from the `Picture` property of a Form, `PictureBox`, or Image control, or from the
//!   `Image` property of a `PictureBox` or Form.
//! - **stringexpression**: Required. String expression specifying the name of the file to which
//!   the graphic is saved. Can include a drive and path specification.
//!
//! ## Remarks
//!
//! - **File Format**: `SavePicture` saves graphics in bitmap (.bmp) format. The file created is
//!   compatible with bitmap files created by other applications.
//! - **Picture Property**: When used with the `Picture` property, `SavePicture` saves the persistent
//!   bitmap from the property. This is the image loaded at design time or assigned at run time via
//!   `LoadPicture` or other means.
//! - **Image Property**: When used with the `Image` property, `SavePicture` saves the current
//!   appearance of the form or picture box, including any graphics drawn with graphics methods.
//!   This creates a snapshot of the visible content.
//! - **File Path**: If no path is specified, the file is saved in the current directory.
//! - **Overwriting**: If a file with the specified name already exists, it is overwritten without
//!   warning.
//! - **Relative Paths**: You can use relative path specifications (e.g., "..\Images\MyPic.bmp").
//! - **Graphics Methods**: To save graphics created with graphics methods (Line, Circle, `PSet`,
//!   etc.), you must use the `Image` property, not the `Picture` property.
//! - **Clipboard Graphics**: `SavePicture` can also be used with graphics from the Clipboard object.
//!
//! ## Examples
//!
//! ### Save Form's Picture Property
//!
//! ```vb
//! ' Save the persistent bitmap from a form
//! SavePicture Form1.Picture, "C:\Images\Form1.bmp"
//! ```
//!
//! ### Save Form's Current Appearance
//!
//! ```vb
//! ' Save the current appearance of a form (including drawn graphics)
//! SavePicture Form1.Image, "C:\Images\FormSnapshot.bmp"
//! ```
//!
//! ### Save `PictureBox` Image
//!
//! ```vb
//! ' Save the picture from a PictureBox control
//! SavePicture Picture1.Picture, "C:\Temp\MyPicture.bmp"
//! ```
//!
//! ### Save with Variable Path
//!
//! ```vb
//! Dim FileName As String
//! FileName = "C:\Output\Image_" & Format$(Now, "yyyymmdd_hhnnss") & ".bmp"
//! SavePicture Picture1.Image, FileName
//! ```
//!
//! ### Save Clipboard Image
//!
//! ```vb
//! ' Save an image from the clipboard
//! SavePicture Clipboard.GetData(), "C:\Temp\ClipImage.bmp"
//! ```
//!
//! ### Error Handling
//!
//! ```vb
//! On Error Resume Next
//! SavePicture Picture1.Picture, "C:\Images\Output.bmp"
//! If Err.Number <> 0 Then
//!     MsgBox "Error saving picture: " & Err.Description
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Common Errors
//!
//! - **Error 53**: File not found - the specified path does not exist
//! - **Error 75**: Path/File access error - insufficient permissions or read-only file
//! - **Error 76**: Path not found - invalid directory path
//!
//! ## See Also
//!
//! - `LoadPicture` function (load images from files)
//! - `Picture` property (persistent bitmap property)
//! - `Image` property (current appearance snapshot)
//! - Graphics methods (`Line`, `Circle`, `PSet`, etc.)
//!
//! ## References
//!
//! - [SavePicture Statement - Microsoft Docs](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa268097(v=vs.60))

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a `SavePicture` statement.
    pub(crate) fn parse_savepicture_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::SavePictureStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn savepicture_simple() {
        let source = r#"
Sub Test()
    SavePicture Form1.Picture, "output.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("SavePictureKeyword"));
        assert!(debug.contains("Form1"));
        assert!(debug.contains("Picture"));
    }

    #[test]
    fn savepicture_at_module_level() {
        let source = "SavePicture Picture1.Picture, \"image.bmp\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_with_image_property() {
        let source = r#"
Sub Test()
    SavePicture Form1.Image, "snapshot.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("Form1"));
        assert!(debug.contains("Image"));
    }

    #[test]
    fn savepicture_with_path() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Picture, "C:\Images\output.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("Picture1"));
    }

    #[test]
    fn savepicture_with_variable() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Picture, fileName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("fileName"));
    }

    #[test]
    fn savepicture_with_concatenation() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Image, basePath & "\image.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("basePath"));
    }

    #[test]
    fn savepicture_with_function_call() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Picture, GetFileName()
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("GetFileName"));
    }

    #[test]
    fn savepicture_clipboard() {
        let source = r#"
Sub Test()
    SavePicture Clipboard.GetData(), "clipboard.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("Clipboard"));
    }

    #[test]
    fn savepicture_nested_property() {
        let source = r#"
Sub Test()
    SavePicture frmMain.picDisplay.Picture, "display.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("frmMain"));
        assert!(debug.contains("picDisplay"));
    }

    #[test]
    fn savepicture_with_format_function() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Picture, "Image_" & Format$(Now, "yyyymmdd") & ".bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("Picture1"));
    }

    #[test]
    fn savepicture_inside_if_statement() {
        let source = r#"
If saveFlag Then
    SavePicture Picture1.Image, "output.bmp"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_inside_loop() {
        let source = r#"
For i = 1 To 10
    SavePicture Pictures(i).Picture, "Pic" & i & ".bmp"
Next i
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_with_comment() {
        let source = r#"
Sub Test()
    SavePicture Form1.Image, "snapshot.bmp" ' Save form image
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("' Save form image"));
    }

    #[test]
    fn savepicture_preserves_whitespace() {
        let source = "SavePicture   Picture1.Picture  ,   \"file.bmp\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_with_array_element() {
        let source = r#"
Sub Test()
    SavePicture Pictures(index).Picture, fileName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("Pictures"));
        assert!(debug.contains("index"));
    }

    #[test]
    fn savepicture_in_select_case() {
        let source = r#"
Select Case format
    Case 1
        SavePicture Picture1.Picture, "output1.bmp"
    Case 2
        SavePicture Picture1.Image, "output2.bmp"
End Select
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_with_error_handling() {
        let source = r#"
On Error Resume Next
SavePicture Picture1.Picture, fileName
If Err.Number <> 0 Then
    MsgBox "Error"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_in_with_block() {
        let source = r#"
With Picture1
    SavePicture .Picture, "output.bmp"
End With
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("Picture"));
    }

    #[test]
    fn savepicture_multiple_on_same_line() {
        let source = "SavePicture Pic1.Picture, \"a.bmp\": SavePicture Pic2.Picture, \"b.bmp\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("Pic1"));
        assert!(debug.contains("Pic2"));
    }

    #[test]
    fn savepicture_relative_path() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Picture, "..\Images\output.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_in_sub() {
        let source = r#"
Sub SaveCurrentImage()
    SavePicture Form1.Image, "current.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_in_function() {
        let source = r#"
Function ExportImage() As Boolean
    SavePicture Picture1.Picture, outputPath
    ExportImage = True
End Function
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("outputPath"));
    }

    #[test]
    fn savepicture_with_app_path() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Picture, App.Path & "\output.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("App"));
        assert!(debug.contains("Path"));
    }

    #[test]
    fn savepicture_control_array() {
        let source = r#"
Sub Test()
    SavePicture imgArray(5).Picture, "array_item.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("imgArray"));
    }

    #[test]
    fn savepicture_in_class_module() {
        let source = r#"
Private picData As PictureBox

Public Sub ExportPicture(fileName As String)
    SavePicture picData.Picture, fileName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("picData"));
    }

    #[test]
    fn savepicture_with_long_path() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Image, "C:\Program Files\MyApp\Data\Images\snapshot.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_with_line_continuation() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Picture, _
        "output.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }

    #[test]
    fn savepicture_dynamic_filename() {
        let source = r#"
Sub Test()
    SavePicture Picture1.Picture, "Image_" & CStr(counter) & ".bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
        assert!(debug.contains("counter"));
    }

    #[test]
    fn savepicture_unc_path() {
        let source = r#"
Sub Test()
    SavePicture Form1.Image, "\\Server\Share\Images\output.bmp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SavePictureStatement"));
    }
}
