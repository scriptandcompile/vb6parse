//! Defines the `FormFile` struct and related parsing functions for VB6 Form files.
//! Handles extraction of version, objects, attributes, and controls from the CST.
//!
use std::fmt::{Debug, Display};
use std::vec::Vec;

use serde::Serialize;

use crate::{
    files::common::{FileAttributes, FileFormatVersion, ObjectReference},
    io::SourceFile,
    language::{Form, FormRoot},
    lexer::{tokenize, TokenStream},
    parsers::{
        cst::{parse, ConcreteSyntaxTree},
        ParseResult,
    },
};

pub mod control_only;

/// Helper function to serialize `ConcreteSyntaxTree` as `SerializableTree`
fn serialize_cst<S>(cst: &ConcreteSyntaxTree, serializer: S) -> Result<S::Ok, S::Error>
where
    S: serde::Serializer,
{
    cst.to_serializable().serialize(serializer)
}

/// Represents a VB6 Form file.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct FormFile {
    /// The top-level form root (`Form` or `MDIForm`).
    pub form: FormRoot,
    /// The list of object references in the form file.
    pub objects: Vec<ObjectReference>,
    /// The VB6 file format version.
    pub version: FileFormatVersion,
    /// The file attributes extracted from the form file.
    pub attributes: FileAttributes,
    /// The concrete syntax tree of the form file.
    /// Note: This CST excludes nodes that are already extracted into other fields
    #[serde(serialize_with = "serialize_cst")]
    pub cst: ConcreteSyntaxTree,
}

impl Display for FormFile {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "FormFile {{ form name: {:?}, objects count: {:?} }}",
            self.form.name(),
            self.objects.len()
        )
    }
}

impl FormFile {
    /// Parses a VB6 Form file from a `SourceFile`.
    ///
    /// # Arguments
    ///
    /// * `source_file` - The source file containing the VB6 Form code.
    ///
    /// # Returns
    /// * `ParseResult<FormFile>` - The result of parsing, containing either the `FormFile` or parsing errors.
    ///
    #[must_use]
    pub fn parse(source_file: &SourceFile) -> ParseResult<'_, Self> {
        let mut source_stream = source_file.source_stream();

        // TODO: Handle errors from tokenization.
        let token_stream = tokenize(&mut source_stream).unwrap();

        // Phase 7: Use direct extraction instead of building full CST first
        let tokens = token_stream.into_tokens();
        let mut parser = crate::parsers::cst::Parser::new_direct_extraction(tokens, 0);

        // Collect all parsing failures
        let mut all_failures = Vec::new();

        // Parse VERSION directly (no CST overhead)
        let (version_opt, version_failures) = parser.parse_version_direct().unpack();
        all_failures.extend(version_failures);
        let version = version_opt.unwrap_or(FileFormatVersion { major: 5, minor: 0 });

        // Parse Objects directly (no CST overhead)
        let objects = parser.parse_objects_direct();

        // Parse form root directly (no CST overhead)
        let (form_root_opt, form_failures) = parser.parse_properties_block_to_form_root().unpack();
        all_failures.extend(form_failures);

        let mut form = form_root_opt.unwrap_or_else(|| {
            use crate::language::FormProperties;

            FormRoot::Form(Form {
                name: String::new(),
                tag: String::new(),
                index: 0,
                properties: FormProperties::default(),
                controls: Vec::new(),
                menus: Vec::new(),
            })
        });

        // Parse Attributes directly (no CST overhead)
        let attributes = parser.parse_attributes_direct();

        // The form's name comes from the VB_Name attribute if present,
        // otherwise from the PropertiesName in the Begin statement
        if !attributes.name.is_empty() {
            form.name_mut().clone_from(&attributes.name);
        }
        // If attributes.name is empty, form.name already has the name from the Begin statement

        // Get remaining tokens and build CST only for the code section
        let remaining_tokens = parser.into_tokens();
        let remaining_stream = TokenStream::from_tokens(remaining_tokens);
        let cst = parse(remaining_stream);

        ParseResult::new(
            Some(FormFile {
                form,
                objects,
                version,
                attributes,
                cst,
            }),
            all_failures,
        )
    }

    /// Parse only the VERSION statement and the first control (Form) from a token stream.
    ///
    /// This is a fast-path parsing method that stops after parsing the control definition,
    /// without parsing the code section or creating a full CST. It's useful for scenarios
    /// that only need UI information (control hierarchy, properties) and don't need the
    /// code implementation.
    ///
    /// # Arguments
    ///
    /// * `token_stream` - The token stream to parse.
    ///
    /// # Returns
    ///
    /// * `ParseResult<ControlOnlyResult>` - A tuple containing:
    ///   - `Option<FileFormatVersion>` - The parsed version
    ///   - `Option<FormRoot>` - The parsed form root (`Form` or `MDIForm` with nested controls)
    ///   - `TokenStream` - The remaining token stream positioned after the control
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::{SourceFile, tokenize, FormFile};
    ///
    /// let source = b"VERSION 5.00
    /// Begin VB.Form Form1
    ///    Caption = \"Test Form\"
    ///    Begin VB.CommandButton Command1
    ///       Caption = \"Click Me\"
    ///    End
    /// End
    /// ";
    ///
    /// let source_file = SourceFile::decode_with_replacement("test.frm", source).expect("Failed to decode source file");
    /// let mut source_stream = source_file.source_stream();
    /// let result = tokenize(&mut source_stream);
    /// let (token_stream, _tok_failures) = result.unpack();
    ///
    /// if let Some(ts) = token_stream {
    ///     let result = FormFile::parse_control_only(ts);
    ///     let (parse_result, failures) = result.unpack();
    ///
    ///     if let Some((version, form, _remaining)) = parse_result {
    ///         if let Some(f) = form {
    ///             println!("Form: {}", f.name());
    ///         }
    ///     }
    /// }
    /// ```
    #[must_use]
    pub fn parse_control_only(
        token_stream: TokenStream<'_>,
    ) -> ParseResult<'_, control_only::ControlOnlyResult<'_>> {
        control_only::parse_control_from_tokens(token_stream)
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    use crate::files::common::extract_version;
    use crate::io::SourceFile;
    use crate::SyntaxKind;
    use either::Either;

    #[test]
    fn extract_version_from_cst() {
        // Test VERSION 5.00 (typical form version)
        let source = "VERSION 5.00\nBegin VB.Form Form1\nEnd\n";
        let mut source_stream = crate::SourceStream::new("test.frm", source);
        let token_stream = tokenize(&mut source_stream).unwrap();
        let cst = parse(token_stream);

        let version = extract_version(&cst);
        assert!(version.is_some());
        let version = version.unwrap();
        assert_eq!(version.major, 5);
        assert_eq!(version.minor, 0);
    }

    #[test]
    fn extract_version_from_cst_class() {
        // Test VERSION 1.0 CLASS (typical class version)
        let source = "VERSION 1.0 CLASS\nBegin\nEnd\n";
        let mut source_stream = crate::SourceStream::new("test.cls", source);
        let token_stream = tokenize(&mut source_stream).unwrap();
        let cst = parse(token_stream);

        let version = extract_version(&cst);
        assert!(version.is_some());
        let version = version.unwrap();
        assert_eq!(version.major, 1);
        assert_eq!(version.minor, 0);
    }

    #[test]
    fn extract_version_no_version_statement() {
        // Test without VERSION statement
        let source = "Begin VB.Form Form1\nEnd\n";
        let mut source_stream = crate::SourceStream::new("test.frm", source);
        let token_stream = tokenize(&mut source_stream).unwrap();
        let cst = parse(token_stream);

        let version = extract_version(&cst);
        assert!(version.is_none());
    }

    #[test]
    fn object_statement_parsing() {
        // Test that Object statements are now parsed as ObjectStatement nodes
        let source = r#"VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1
End
"#;
        let mut source_stream = crate::SourceStream::new("test.frm", source);
        let token_stream = tokenize(&mut source_stream).unwrap();
        let cst = parse(token_stream);

        // Verify we have VersionStatement
        assert!(cst.contains_kind(SyntaxKind::VersionStatement));

        // Verify we have ObjectStatements
        assert!(cst.contains_kind(SyntaxKind::ObjectStatement));

        // Verify we have exactly 2 ObjectStatements
        let obj_statements: Vec<_> = cst.children_by_kind(SyntaxKind::ObjectStatement).collect();
        assert_eq!(obj_statements.len(), 2);

        // Verify the content of the first Object statement
        assert!(obj_statements[0]
            .text()
            .contains("831FDD16-0C5C-11D2-A9FC-0000F8754DA1"));
        assert!(obj_statements[0].text().contains("mscomctl.ocx"));

        // Verify the content of the second Object statement
        assert!(obj_statements[1]
            .text()
            .contains("F9043C88-F6F2-101A-A3C9-08002B2F49FB"));
        assert!(obj_statements[1].text().contains("COMDLG32.OCX"));
    }

    #[allow(clippy::too_many_lines)]
    #[test]
    fn nested_property_group() {
        use crate::files::form::FormFile;

        let input = b"VERSION 5.00\r
    Object = \"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0\"; \"mscomctl.ocx\"\r
    Begin VB.Form Form_Main \r
       BackColor       =   &H00000000&\r
       BorderStyle     =   1  'Fixed Single\r
       Caption         =   \"Audiostation\"\r
       ClientHeight    =   10005\r
       ClientLeft      =   4695\r
       ClientTop       =   1275\r
       ClientWidth     =   12960\r
       BeginProperty Font \r
          Name            =   \"Verdana\"\r
          Size            =   8.25\r
          Charset         =   0\r
          Weight          =   400\r
          Underline       =   0   'False\r
          Italic          =   0   'False\r
          Strikethrough   =   0   'False\r
       EndProperty\r
       LinkTopic       =   \"Form1\"\r
       MaxButton       =   0   'False\r
       OLEDropMode     =   1  'Manual\r
       ScaleHeight     =   10005\r
       ScaleWidth      =   12960\r
       StartUpPosition =   2  'CenterScreen\r
       Begin MSComctlLib.ImageList Imagelist_CDDisplay \r
          Left            =   12000\r
          Top             =   120\r
          _ExtentX        =   1005\r
          _ExtentY        =   1005\r
          BackColor       =   -2147483643\r
          ImageWidth      =   53\r
          ImageHeight     =   42\r
          MaskColor       =   12632256\r
          _Version        =   393216\r
          BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} \r
             NumListImages   =   5\r
             BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
                _Version        =   9\r
                Key             =   \"\"\r
             EndProperty\r
             BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
                _Version        =   1\r
                Key             =   \"\"\r
             EndProperty\r
             BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
                _Version        =   1\r
                Key             =   \"\"\r
             EndProperty\r
             BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
                _Version        =   5\r
                Key             =   \"\"\r
             EndProperty\r
             BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
                _Version        =   1\r
                Key             =   \"\"\r
             EndProperty\r
          EndProperty\r
       End\r
    End\r
    Attribute VB_Name = \"Form_Main\"\r
    ";

        let source_file = SourceFile::decode_with_replacement("form_parse.frm", input).unwrap();
        let (parse_result_opt, _failures) = FormFile::parse(&source_file).unpack();

        assert!(parse_result_opt.is_some());

        let result = parse_result_opt.expect("Expected parse result");

        assert_eq!(result.objects.len(), 1);
        assert_eq!(result.version.major, 5);
        assert_eq!(result.version.minor, 0);
        assert_eq!(result.form.name(), "Form_Main");
        assert!(result.form.is_form());

        if let crate::language::FormRoot::Form(form) = &result.form {
            assert_eq!(form.controls.len(), 1);
            assert_eq!(form.menus.len(), 0);
            assert_eq!(form.properties.caption, "Audiostation");
            assert_eq!(form.controls[0].name(), "Imagelist_CDDisplay");
            assert!(matches!(
                form.controls[0].kind(),
                crate::language::ControlKind::Custom { .. }
            ));

            if let crate::language::ControlKind::Custom {
                properties,
                property_groups,
            } = form.controls[0].kind()
            {
                assert_eq!(properties.len(), 9);
                assert_eq!(property_groups.len(), 1);

                if let Some(group) = property_groups.first() {
                    assert_eq!(group.name, "Images");
                    assert_eq!(group.properties.len(), 6);

                    if let Some(Either::Right(image1)) = group.properties.get("ListImage1") {
                        assert_eq!(image1.name, "ListImage1");
                        assert_eq!(image1.properties.len(), 2);
                    } else {
                        panic!("Expected nested ListImage1");
                    }

                    if let Some(Either::Right(image2)) = group.properties.get("ListImage2") {
                        assert_eq!(image2.name, "ListImage2");
                        assert_eq!(image2.properties.len(), 2);
                    } else {
                        panic!("Expected nested ListImage2");
                    }

                    if let Some(Either::Right(image3)) = group.properties.get("ListImage3") {
                        assert_eq!(image3.name, "ListImage3");
                        assert_eq!(image3.properties.len(), 2);
                    } else {
                        panic!("Expected nested ListImage3");
                    }
                } else {
                    panic!("Expected property group");
                }
            } else {
                panic!("Expected custom control");
            }
        } else {
            panic!("Expected form kind");
        }
    }

    #[test]
    fn parse_indented_menu_valid() {
        use crate::language::VB_WINDOW_BACKGROUND;
        use crate::language::{MenuControl, MenuProperties};

        let input = b"VERSION 5.00\r
        Begin VB.Form frmExampleForm\r
            BackColor       =   &H80000005&\r
            Caption         =   \"example form\"\r
            ClientHeight    =   6210\r
            ClientLeft      =   60\r
            ClientTop       =   645\r
            ClientWidth     =   9900\r
            BeginProperty Font\r
                Name            =   \"Arial\"\r
                Size            =   8.25\r
                Charset         =   0\r
                Weight          =   400\r
                Underline       =   0   'False\r
                Italic          =   0   'False\r
                Strikethrough   =   0   'False\r
            EndProperty\r
            LinkTopic       =   \"Form1\"\r
            ScaleHeight     =   414\r
            ScaleMode       =   3  'Pixel\r
            ScaleWidth      =   660\r
            StartUpPosition =   2  'CenterScreen\r
            Begin VB.Menu mnuFile\r
                Caption         =   \"&File\"\r
                Begin VB.Menu mnuOpenImage\r
                    Caption         =   \"&Open image\"\r
               End\r
            End\r
        End\r
        Attribute VB_Name = \"frmExampleForm\"\r
        ";

        let source_file = SourceFile::decode_with_replacement("form_parse.frm", input.as_ref());
        let source_file = source_file.expect("Expected source file");

        let (parse_result_opt, _failures) = FormFile::parse(&source_file).unpack();

        assert!(parse_result_opt.is_some());

        let result = parse_result_opt.expect("Expected parse result");

        assert_eq!(result.version.major, 5);
        assert_eq!(result.version.minor, 0);

        assert_eq!(result.form.name(), "frmExampleForm");

        if let crate::language::FormRoot::Form(form) = &result.form {
            assert_eq!(form.tag, "");
            assert_eq!(form.index, 0);
            assert_eq!(form.controls.len(), 0);
            assert_eq!(form.menus.len(), 1);
            assert_eq!(form.properties.caption, "example form");
            assert_eq!(form.properties.back_color, VB_WINDOW_BACKGROUND);
            assert_eq!(
                form.menus,
                vec![MenuControl::new(
                    "mnuFile".into(),
                    String::new(),
                    0,
                    MenuProperties {
                        caption: "&File".into(),
                        ..Default::default()
                    },
                    vec![MenuControl::new(
                        "mnuOpenImage".into(),
                        String::new(),
                        0,
                        MenuProperties {
                            caption: "&Open image".into(),
                            ..Default::default()
                        },
                        vec![],
                    )]
                )]
            );
        } else {
            panic!("Expected form kind");
        }
    }

    #[test]
    fn extract_form_with_controls() {
        let input = r#"VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Form"
   ClientHeight    =   3195
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
"#;

        let source_file =
            SourceFile::decode_with_replacement("test.frm", input.as_bytes()).unwrap();
        let (result_opt, _failures) = FormFile::parse(&source_file).unpack();

        assert!(result_opt.is_some());
        let form_file = result_opt.expect("Expected parse result");

        assert_eq!(form_file.form.name(), "Form1");

        if let crate::language::FormRoot::Form(form) = &form_file.form {
            assert_eq!(form.controls.len(), 2);
            assert_eq!(form.controls[0].name(), "Command1");
            assert_eq!(form.controls[1].name(), "Text1");

            // Check control types
            assert!(matches!(
                form.controls[0].kind(),
                crate::language::ControlKind::CommandButton { .. }
            ));
            assert!(matches!(
                form.controls[1].kind(),
                crate::language::ControlKind::TextBox { .. }
            ));
        } else {
            panic!("Expected Form kind");
        }
    }
}
