//! Defines the `FormFile` struct and related parsing functions for VB6 Form files.
//! Handles extraction of version, objects, attributes, and controls from the CST.
//!
use std::collections::HashMap;
use std::fmt::{Debug, Display};
use std::vec::Vec;

use either::Either;
use serde::Serialize;
use uuid::Uuid;

use crate::{
    cst::{parse, ConcreteSyntaxTree},
    errors::FormErrorKind,
    language::{Control, ControlKind, MenuControl, PropertyGroup},
    parsers::{
        header::{extract_attributes, extract_version, FileAttributes, FileFormatVersion},
        CstNode, ObjectReference, ParseResult, SyntaxKind,
    },
    tokenize, Properties, SourceFile,
};

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
    /// The form control and its hierarchy.
    pub form: Control,
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
            self.form.name,
            self.objects.len()
        )
    }
}

/// Extract the VB6 file format version from a CST.
///
/// Searches for a `VersionStatement` node in the CST and traverses its children
/// to find the `SingleLiteral` token containing the version number (e.g., "5.00" or "1.0").
/// The expected structure is:
///     - `VersionKeyword` ("VERSION")
///     - Whitespace
///     - `SingleLiteral` (the version number as a float)
///     - Optional: Whitespace, `ClassKeyword`, `FormKeyword`
///
/// # Arguments
///
/// * `cst` - The concrete syntax tree to extract the version from
///
/// # Returns
///
/// * `Some(VB6FileFormatVersion)` - If a valid version statement is found with a parseable version
/// * `None` - If no version statement is found or the version cannot be parsed
///
/// VB6 object statements have the format:
///
/// `Object = "{UUID}#version#flags"; "filename"`
///
/// or
///
/// `Object = *\G{UUID}#version#flags; "filename"`
fn extract_objects(cst: &ConcreteSyntaxTree) -> Vec<ObjectReference> {
    let mut objects = Vec::new();

    // Find all ObjectStatement nodes
    let obj_statements: Vec<_> = cst
        .children()
        .into_iter()
        .filter(|c| c.kind == SyntaxKind::ObjectStatement)
        .collect();

    for obj_stmt in obj_statements {
        // Navigate the children of the ObjectStatement node
        // Expected structure: ObjectKeyword, Whitespace, EqualityOperator, Whitespace,
        //                     [MultiplicationOperator, Identifier,] StringLiteral, Semicolon, Whitespace, StringLiteral

        let mut uuid_part = String::new();
        let mut file_name = String::new();
        let mut found_equals = false;
        let mut found_semicolon = false;

        for child in &obj_stmt.children {
            if !child.is_token {
                continue;
            }

            match child.kind {
                SyntaxKind::EqualityOperator => {
                    found_equals = true;
                }
                SyntaxKind::StringLiteral if found_equals && !found_semicolon => {
                    // First string literal: contains UUID#version#flags
                    uuid_part = child.text.trim_matches('"').to_string();
                }
                SyntaxKind::Semicolon => {
                    found_semicolon = true;
                }
                SyntaxKind::StringLiteral if found_semicolon => {
                    // Second string literal: contains filename
                    file_name = child.text.trim_matches('"').to_string();
                    break;
                }
                _ => {}
            }
        }

        // Parse the UUID part: {UUID}#version#flags
        if !uuid_part.is_empty() && !file_name.is_empty() {
            let parts: Vec<&str> = uuid_part.split('#').collect();
            if parts.len() >= 3 {
                // Extract UUID (remove braces)
                let uuid_str = parts[0].trim_matches(|c| c == '{' || c == '}');

                if let Ok(uuid) = Uuid::parse_str(uuid_str) {
                    let version = parts[1].to_string();
                    let unknown1 = parts[2].to_string();

                    objects.push(ObjectReference::Compiled {
                        uuid,
                        version,
                        unknown1,
                        file_name,
                    });
                }
            }
        }
    }

    objects
}

/// Extracts the form and its controls from the CST.
///
/// Recursively processes BEGIN...END blocks (`PropertiesBlock` nodes) to build
/// a hierarchy of `VB6Control` structures.
fn extract_control(cst: &ConcreteSyntaxTree) -> Option<Control> {
    // Find the first PropertiesBlock (should be the form)
    let properties_blocks: Vec<_> = cst
        .children()
        .into_iter()
        .filter(|c| c.kind == SyntaxKind::PropertiesBlock)
        .collect();

    if properties_blocks.is_empty() {
        return None;
    }

    // Process the first block which should be the form
    Some(extract_properties_block(&properties_blocks[0]))
}

/// Extracts a `VB6Control` from a `PropertiesBlock` CST node.
///
/// Recursively processes nested `PropertiesBlock` nodes for child controls.
fn extract_properties_block(block: &CstNode) -> Control {
    // Extract the type and name from the PropertiesBlock
    let mut control_type = String::new();
    let mut control_name = String::new();
    let mut properties = Properties::new();
    let mut child_blocks: Vec<&CstNode> = Vec::new();
    let mut property_groups: Vec<PropertyGroup> = Vec::new();

    for child in &block.children {
        match child.kind {
            SyntaxKind::PropertiesType => {
                // Extract the full type name (e.g., "VB.Form", "VB.CommandButton")
                control_type = child.text.trim().to_string();
            }
            SyntaxKind::PropertiesName => {
                // Extract the control name
                control_name = child.text.trim().to_string();
            }
            SyntaxKind::Property => {
                // Extract key-value properties
                if let Some((key, value)) = extract_property(child) {
                    properties.insert(key, value);
                }
            }
            SyntaxKind::PropertyGroup => {
                // Extract property group
                if let Some(group) = extract_property_group(child) {
                    property_groups.push(group);
                }
            }
            SyntaxKind::PropertiesBlock => {
                // Nested control
                child_blocks.push(child);
            }
            _ => {}
        }
    }

    // Parse child controls and menus recursively
    // First pass: determine type of each child block
    let mut child_controls: Vec<Control> = Vec::new();
    let mut menu_blocks: Vec<&CstNode> = Vec::new();

    for child_block in child_blocks {
        // Check if this is a menu by looking at its PropertiesType
        let mut is_menu = false;
        for child in &child_block.children {
            if child.kind == SyntaxKind::PropertiesType {
                let block_type = child.text.trim();
                if block_type == "VB.Menu" {
                    is_menu = true;
                    break;
                }
            }
        }

        if is_menu {
            menu_blocks.push(child_block);
        } else {
            child_controls.push(extract_properties_block(child_block));
        }
    }

    // Extract menus
    let mut menus: Vec<MenuControl> = Vec::new();
    for menu_block in menu_blocks {
        let menu = extract_menu_control(menu_block);
        menus.push(menu);
    }

    let tag = properties.get("Tag").cloned().unwrap_or_default();
    let index = properties
        .get("Index")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);

    // Determine the control kind based on the type
    let kind = match control_type.as_str() {
        "VB.Form" => ControlKind::Form {
            properties: properties.into(),
            controls: child_controls,
            menus,
        },
        "VB.CommandButton" => ControlKind::CommandButton {
            properties: properties.into(),
        },
        "VB.TextBox" => ControlKind::TextBox {
            properties: properties.into(),
        },
        "VB.Label" => ControlKind::Label {
            properties: properties.into(),
        },
        "VB.PictureBox" => ControlKind::PictureBox {
            properties: properties.into(),
        },
        "VB.Frame" => ControlKind::Frame {
            properties: properties.into(),
            controls: child_controls,
        },
        "VB.CheckBox" => ControlKind::CheckBox {
            properties: properties.into(),
        },
        "VB.OptionButton" => ControlKind::OptionButton {
            properties: properties.into(),
        },
        "VB.ListBox" => ControlKind::ListBox {
            properties: properties.into(),
        },
        "VB.ComboBox" => ControlKind::ComboBox {
            properties: properties.into(),
        },
        "VB.Timer" => ControlKind::Timer {
            properties: properties.into(),
        },
        "VB.HScrollBar" => ControlKind::HScrollBar {
            properties: properties.into(),
        },
        "VB.VScrollBar" => ControlKind::VScrollBar {
            properties: properties.into(),
        },
        "VB.Image" => ControlKind::Image {
            properties: properties.into(),
        },
        "VB.Line" => ControlKind::Line {
            properties: properties.into(),
        },
        "VB.Shape" => ControlKind::Shape {
            properties: properties.into(),
        },
        "VB.FileListBox" => ControlKind::FileListBox {
            properties: properties.into(),
        },
        "VB.DirListBox" => ControlKind::DirListBox {
            properties: properties.into(),
        },
        "VB.DriveListBox" => ControlKind::DriveListBox {
            properties: properties.into(),
        },
        "VB.Data" => ControlKind::Data {
            properties: properties.into(),
        },
        "VB.OLE" => ControlKind::Ole {
            properties: properties.into(),
        },
        _ => {
            // Unknown or custom control
            ControlKind::Custom {
                properties: properties.into(),
                property_groups,
            }
        }
    };

    Control {
        name: control_name,
        tag,
        index,
        kind,
    }
}

/// Extracts a `MenuControl` from a `PropertiesBlock` CST node.
///
/// Recursively processes nested `PropertiesBlock` nodes for sub-menus.
fn extract_menu_control(block: &CstNode) -> MenuControl {
    // Extract the name and properties from the menu block
    let mut menu_name = String::new();
    let mut properties = Properties::new();
    let mut child_menu_blocks: Vec<&CstNode> = Vec::new();

    for child in &block.children {
        match child.kind {
            SyntaxKind::PropertiesName => {
                // Extract the menu name
                menu_name = child.text.trim().to_string();
            }
            SyntaxKind::Property => {
                // Extract key-value properties
                if let Some((key, value)) = extract_property(child) {
                    properties.insert(key, value);
                }
            }
            SyntaxKind::PropertiesBlock => {
                // Check if this is a nested menu
                for sub_child in &child.children {
                    if sub_child.kind == SyntaxKind::PropertiesType {
                        let block_type = sub_child.text.trim();
                        if block_type == "VB.Menu" {
                            child_menu_blocks.push(child);
                            break;
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // Recursively extract sub-menus
    let mut sub_menus: Vec<MenuControl> = Vec::new();
    for child_menu_block in child_menu_blocks {
        let sub_menu = extract_menu_control(child_menu_block);
        sub_menus.push(sub_menu);
    }

    let tag = properties.get("Tag").cloned().unwrap_or_default();
    let index = properties
        .get("Index")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);

    MenuControl {
        name: menu_name,
        tag,
        index,
        properties: properties.into(),
        sub_menus,
    }
}

/// Extracts a `PropertyGroup` from a `PropertyGroup` CST node.
fn extract_property_group(group_node: &CstNode) -> Option<PropertyGroup> {
    let mut name = String::new();
    let mut guid: Option<Uuid> = None;
    let mut properties: HashMap<String, Either<String, PropertyGroup>> = HashMap::new();

    // Extract the property group name and GUID
    for child in &group_node.children {
        if child.kind == SyntaxKind::PropertyGroupName {
            name = child.text.trim().to_string();
            // Check if there's a GUID in the text after the name
            let full_text = child.text.trim();
            if let Some(start) = full_text.find('{') {
                if let Some(end) = full_text.find('}') {
                    let uuid_str = &full_text[start + 1..end];
                    if let Ok(uuid) = Uuid::parse_str(uuid_str) {
                        guid = Some(uuid);
                    }
                }
            }
        }
    }

    // If GUID wasn't in the name node, check the parent text
    if guid.is_none() {
        let full_text = group_node.text.trim();
        if let Some(start) = full_text.find('{') {
            if let Some(end) = full_text.find('}') {
                let uuid_str = &full_text[start + 1..end];
                if let Ok(uuid) = Uuid::parse_str(uuid_str) {
                    guid = Some(uuid);
                }
            }
        }
    }

    // Extract properties and nested property groups
    for child in &group_node.children {
        match child.kind {
            SyntaxKind::Property => {
                if let Some((key, value)) = extract_property(child) {
                    properties.insert(key, Either::Left(value));
                }
            }
            SyntaxKind::PropertyGroup => {
                if let Some(nested_group) = extract_property_group(child) {
                    let group_name = nested_group.name.clone();
                    properties.insert(group_name, Either::Right(nested_group));
                }
            }
            _ => {}
        }
    }

    if name.is_empty() {
        None
    } else {
        Some(PropertyGroup {
            name,
            guid,
            properties,
        })
    }
}

/// Extracts a key-value pair from a Property CST node.
fn extract_property(property_node: &CstNode) -> Option<(String, String)> {
    let mut key = String::new();
    let mut value = String::new();

    for child in &property_node.children {
        match child.kind {
            SyntaxKind::PropertyKey => {
                key = child.text.trim().to_string();
            }
            SyntaxKind::PropertyValue => {
                let trimmed = child.text.trim();
                // Remove surrounding quotes if this is a string literal
                value = if trimmed.starts_with('"') && trimmed.ends_with('"') && trimmed.len() >= 2
                {
                    trimmed[1..trimmed.len() - 1].to_string()
                } else {
                    trimmed.to_string()
                };
            }
            _ => {}
        }
    }

    if key.is_empty() {
        None
    } else {
        Some((key, value))
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
    /// * `ParseResult<FormFile, FormErrorKind>` - The result of parsing, containing either the `FormFile` or parsing errors.
    ///
    #[must_use]
    pub fn parse(source_file: &SourceFile) -> ParseResult<'_, Self, FormErrorKind> {
        let mut source_stream = source_file.get_source_stream();

        // TODO: Handle errors from tokenization.
        let token_stream = tokenize(&mut source_stream).unwrap();

        let cst = parse(token_stream);

        // Extract version from CST
        let format_version = extract_version(&cst);

        // Extract objects from CST
        let objects = extract_objects(&cst);

        // Extract form and controls from CST
        let mut form = extract_control(&cst).unwrap_or_else(|| {
            use crate::language::FormProperties;

            Control {
                name: String::new(),
                tag: String::new(),
                index: 0,
                kind: ControlKind::Form {
                    properties: FormProperties::default(),
                    controls: Vec::new(),
                    menus: Vec::new(),
                },
            }
        });

        // Extract attributes from CST
        let attributes = extract_attributes(&cst);

        // The form's name comes from the VB_Name attribute if present,
        // otherwise from the PropertiesName in the Begin statement
        if !attributes.name.is_empty() {
            form.name = attributes.name.clone();
        }
        // If attributes.name is empty, form.name already has the name from the Begin statement

        // Filter out nodes that are already extracted to avoid duplication
        let filtered_cst = cst.without_kinds(&[
            SyntaxKind::VersionStatement,
            SyntaxKind::ObjectStatement,
            SyntaxKind::AttributeStatement,
            SyntaxKind::PropertiesBlock, // Form and all controls (already in form field)
        ]);

        ParseResult {
            result: Some(FormFile {
                form,
                objects,
                version: format_version.unwrap_or(FileFormatVersion { major: 5, minor: 0 }),
                attributes,
                cst: filtered_cst,
            }),
            failures: Vec::new(),
        }
    }
}

#[cfg(test)]
mod tests {

    use super::*;
    use crate::SourceFile;

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
        let obj_statements = cst.find_children_by_kind(SyntaxKind::ObjectStatement);
        assert_eq!(obj_statements.len(), 2);

        // Verify the content of the first Object statement
        assert!(obj_statements[0]
            .text
            .contains("831FDD16-0C5C-11D2-A9FC-0000F8754DA1"));
        assert!(obj_statements[0].text.contains("mscomctl.ocx"));

        // Verify the content of the second Object statement
        assert!(obj_statements[1]
            .text
            .contains("F9043C88-F6F2-101A-A3C9-08002B2F49FB"));
        assert!(obj_statements[1].text.contains("COMDLG32.OCX"));
    }

    #[test]
    fn nested_property_group() {
        use crate::parsers::form::FormFile;

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
        let parse_result = FormFile::parse(&source_file);

        assert!(parse_result.result.is_some());

        let result = parse_result.result.unwrap();

        assert_eq!(result.objects.len(), 1);
        assert_eq!(result.version.major, 5);
        assert_eq!(result.version.minor, 0);
        assert_eq!(result.form.name, "Form_Main");
        assert!(matches!(result.form.kind, ControlKind::Form { .. }));

        if let ControlKind::Form {
            controls,
            properties,
            menus,
        } = &result.form.kind
        {
            assert_eq!(controls.len(), 1);
            assert_eq!(menus.len(), 0);
            assert_eq!(properties.caption, "Audiostation");
            assert_eq!(controls[0].name, "Imagelist_CDDisplay");
            assert!(matches!(controls[0].kind, ControlKind::Custom { .. }));

            if let ControlKind::Custom {
                properties,
                property_groups,
            } = &controls[0].kind
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

        let source_file =
            SourceFile::decode_with_replacement("form_parse.frm", input.as_ref()).unwrap();
        let parse_result = FormFile::parse(&source_file);

        assert!(parse_result.result.is_some());

        let result = parse_result.result.unwrap();

        assert_eq!(result.version.major, 5);
        assert_eq!(result.version.minor, 0);

        assert_eq!(result.form.name, "frmExampleForm");
        assert_eq!(result.form.tag, "");
        assert_eq!(result.form.index, 0);

        if let ControlKind::Form {
            controls,
            properties,
            menus,
        } = &result.form.kind
        {
            assert_eq!(controls.len(), 0);
            assert_eq!(menus.len(), 1);
            assert_eq!(properties.caption, "example form");
            assert_eq!(properties.back_color, VB_WINDOW_BACKGROUND);
            assert_eq!(
                menus,
                &vec![MenuControl {
                    name: "mnuFile".into(),
                    tag: String::new(),
                    index: 0,
                    properties: MenuProperties {
                        caption: "&File".into(),
                        ..Default::default()
                    },
                    sub_menus: vec![MenuControl {
                        name: "mnuOpenImage".into(),
                        tag: String::new(),
                        index: 0,
                        properties: MenuProperties {
                            caption: "&Open image".into(),
                            ..Default::default()
                        },
                        sub_menus: vec![],
                    }]
                }]
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
        let result = FormFile::parse(&source_file);

        assert!(result.result.is_some());
        let form_file = result.result.unwrap();

        assert_eq!(form_file.form.name, "Form1");

        if let ControlKind::Form { controls, .. } = &form_file.form.kind {
            assert_eq!(controls.len(), 2);
            assert_eq!(controls[0].name, "Command1");
            assert_eq!(controls[1].name, "Text1");

            // Check control types
            assert!(matches!(
                controls[0].kind,
                ControlKind::CommandButton { .. }
            ));
            assert!(matches!(controls[1].kind, ControlKind::TextBox { .. }));
        } else {
            panic!("Expected Form kind");
        }
    }
}
