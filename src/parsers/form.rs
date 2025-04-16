use std::collections::HashMap;
use std::fmt::Debug;
use std::vec::Vec;

use bstr::{BStr, BString, ByteSlice};
use either::Either;
use serde::Serialize;
use uuid::Uuid;
use winnow::{
    ascii::{line_ending, space0, space1},
    combinator::{alt, opt},
    error::{ErrMode, ParserError},
    token::{literal, take_till, take_until},
    Parser,
};

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    language::{VB6Control, VB6ControlKind, VB6MenuControl, VB6Token},
    parsers::{
        header::{
            attributes_parse, key_resource_offset_line_parse, object_parse, version_parse,
            HeaderKind, VB6FileAttributes, VB6FileFormatVersion,
        },
        Properties, VB6ObjectReference, VB6Stream,
    },
    vb6::{keyword_parse, line_comment_parse, string_parse, vb6_parse, VB6Result},
};

/// Represents a VB6 Form file.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct VB6FormFile<'a> {
    pub form: VB6Control,
    pub objects: Vec<VB6ObjectReference<'a>>,
    pub format_version: VB6FileFormatVersion,
    pub attributes: VB6FileAttributes<'a>,
    pub tokens: Vec<VB6Token<'a>>,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
struct VB6FullyQualifiedName {
    pub namespace: BString,
    pub kind: BString,
    pub name: BString,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6PropertyGroup {
    pub name: BString,
    pub guid: Option<Uuid>,
    pub properties: HashMap<BString, Either<BString, VB6PropertyGroup>>,
}

impl Serialize for VB6PropertyGroup {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("VB6PropertyGroup", 3)?;

        state.serialize_field("name", &self.name)?;

        if let Some(guid) = &self.guid {
            state.serialize_field("guid", &guid.to_string())?;
        } else {
            state.serialize_field("guid", &"None")?;
        }

        state.serialize_field("properties", &self.properties)?;

        state.end()
    }
}

impl<'a> VB6FormFile<'a> {
    /// Parses a VB6 form file from a byte slice.
    ///
    /// # Arguments
    ///
    /// * `input` The byte slice to parse.
    ///
    /// # Returns
    ///
    /// A result containing the parsed VB6 form file or an error.
    ///
    /// # Errors
    ///
    /// An error will be returned if the input is not a valid VB6 form file.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::parsers::VB6FormFile;
    ///
    /// let input = b"VERSION 5.00\r
    /// Begin VB.Form frmExampleForm\r
    ///    BackColor       =   &H80000005&\r
    ///    Caption         =   \"example form\"\r
    ///    ClientHeight    =   6210\r
    ///    ClientLeft      =   60\r
    ///    ClientTop       =   645\r
    ///    ClientWidth     =   9900\r
    ///    BeginProperty Font\r
    ///       Name            =   \"Arial\"\r
    ///       Size            =   8.25\r
    ///       Charset         =   0\r
    ///       Weight          =   400\r
    ///       Underline       =   0   'False\r
    ///       Italic          =   0   'False\r
    ///       Strikethrough   =   0   'False\r
    ///    EndProperty\r
    ///    LinkTopic       =   \"Form1\"\r
    ///    ScaleHeight     =   414\r
    ///    ScaleMode       =   3  'Pixel\r
    ///    ScaleWidth      =   660\r
    ///    StartUpPosition =   2  'CenterScreen\r
    ///    Begin VB.Menu mnuFile\r
    ///       Caption         =   \"&File\"\r
    ///       Begin VB.Menu mnuOpenImage\r
    ///          Caption         =   \"&Open image\"\r
    ///       End\r
    ///    End\r
    /// End\r
    /// Attribute VB_Name = \"frmExampleForm\"\r
    /// ";
    ///
    /// let result = VB6FormFile::parse("form_parse.frm".to_owned(), &mut input.as_ref(), resource_file_resolver);
    ///
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(
        file_name: String,
        input: &'a [u8],
        resource_resolver: impl Fn(String, u32) -> Result<Vec<u8>, std::io::Error>,
    ) -> Result<Self, VB6Error> {
        let mut input = VB6Stream::new(file_name, input);

        let format_version = match version_parse(HeaderKind::Form).parse_next(&mut input) {
            Ok(version) => version,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        let objects = match form_object_parse.parse_next(&mut input) {
            Ok(objects) => objects,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        match (space0, keyword_parse("BEGIN"), space1).parse_next(&mut input) {
            Ok(_) => (),
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        let form = match block_parse(&resource_resolver).parse_next(&mut input) {
            Ok(form) => form,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        let attributes = match attributes_parse.parse_next(&mut input) {
            Ok(attributes) => attributes,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        let tokens = match vb6_parse.parse_next(&mut input) {
            Ok(tokens) => tokens,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        Ok(VB6FormFile {
            form,
            objects,
            format_version,
            attributes,
            tokens,
        })
    }
}

pub fn resource_file_resolver<'a>(
    filename: String,
    _offset: u32,
) -> Result<Vec<u8>, std::io::Error> {
    // This is a stub for the resource file resolver.
    // In a real implementation, this function would look up the resource file
    // based on the filename and offset provided.
    Err(std::io::Error::new(
        std::io::ErrorKind::NotFound,
        format!("Resource file not found: {}", filename),
    ))
}

fn form_object_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<Vec<VB6ObjectReference<'a>>> {
    let mut objects = vec![];

    loop {
        space0.parse_next(input)?;

        if literal::<_, _, VB6ErrorKind>("Object")
            .parse_next(input)
            .is_err()
        {
            break;
        }

        let object = object_parse.parse_next(input)?;

        objects.push(object);
    }

    Ok(objects)
}

fn block_parse<'a>(
    resource_resolver: impl Fn(String, u32) -> Result<Vec<u8>, std::io::Error>,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<VB6Control> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<VB6Control> {
        let fully_qualified_name = property_parse.parse_next(input)?;

        let mut controls = vec![];
        let mut menus = vec![];
        let mut property_groups = vec![];
        let mut properties = Properties::new();

        while !input.is_empty() {
            if (space0, keyword_parse("END"), space0, line_ending)
                .parse_next(input)
                .is_ok()
            {
                match build_control(
                    fully_qualified_name,
                    controls,
                    menus,
                    properties,
                    property_groups,
                ) {
                    Ok(control) => return Ok(control),
                    Err(err) => return Err(ErrMode::Cut(err)),
                };
            }

            if (space0, keyword_parse("BEGIN"), space1)
                .parse_next(input)
                .is_ok()
            {
                let control = block_parse(&resource_resolver).parse_next(input)?;
                if control.kind.is_menu() {
                    menus.push(control);
                } else {
                    controls.push(control);
                }

                continue;
            }

            if let Ok(property_group) = property_group_parse.parse_next(input) {
                property_groups.push(property_group);
                continue;
            }

            if let Ok((name, resource_file, offset)) =
                key_resource_offset_line_parse.parse_next(input)
            {
                let resource = match resource_resolver(resource_file.to_string(), offset) {
                    Ok(res) => res,
                    Err(err) => {
                        return Err(ErrMode::Cut(VB6ErrorKind::ResourceFile(err)));
                    }
                };

                properties.insert_resource(name, resource);

                continue;
            }

            space0.parse_next(input)?;

            let name = take_till(1.., (b' ', b'\t', b'=')).parse_next(input)?;

            space0.parse_next(input)?;

            "=".parse_next(input)?;

            space0.parse_next(input)?;

            let value = alt((string_parse, take_till(1.., (' ', '\t', '\'', '\r', '\n'))))
                .parse_next(input)?;

            properties.insert(name, value.as_bytes());

            space0.parse_next(input)?;

            opt(line_comment_parse).parse_next(input)?;

            line_ending.parse_next(input)?;
        }

        Err(ParserError::assert(input, "Unknown control kind"))
    }
}

fn property_group_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6PropertyGroup> {
    (space0, keyword_parse("BeginProperty"), space1).parse_next(input)?;

    let Ok(property_group_name): VB6Result<&BStr> = take_till(1.., |c| {
        c == b'{' || c == b'\r' || c == b'\t' || c == b' ' || c == b'\n'
    })
    .parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoPropertyName));
    };

    space0.parse_next(input)?;

    // Check if we have a GUID here.
    let guid = if literal::<_, _, VB6ErrorKind>('{').parse_next(input).is_ok() {
        let uuid_segment = take_until(1.., '}').parse_next(input)?;
        '}'.parse_next(input)?;

        space0.parse_next(input)?;

        let Ok(uuid) = Uuid::parse_str(uuid_segment.to_str().unwrap()) else {
            return Err(ErrMode::Cut(VB6ErrorKind::UnableToParseUuid));
        };

        Some(uuid)
    } else {
        None
    };

    alt((line_comment_parse, line_ending)).parse_next(input)?;

    let mut property_group = VB6PropertyGroup {
        guid,
        name: property_group_name.into(),
        properties: HashMap::new(),
    };

    while !input.is_empty() {
        if (space0, keyword_parse("EndProperty"), space0)
            .parse_next(input)
            .is_ok()
        {
            if line_ending::<_, VB6Error>.parse_next(input).is_err() {
                return Err(ErrMode::Cut(VB6ErrorKind::NoLineEndingAfterEndProperty));
            }

            break;
        }

        // looks like we have a nested property group.
        if let Ok(nested_property_group) = property_group_parse.parse_next(input) {
            property_group.properties.insert(
                nested_property_group.name.clone(),
                Either::Right(nested_property_group),
            );

            continue;
        }

        if let Ok((name, _resource_file, _offset)) =
            key_resource_offset_line_parse.parse_next(input)
        {
            // TODO: At the moment we just eat the resource file look up.
            property_group
                .properties
                .insert(name.into(), Either::Left("".into()));

            continue;
        }

        space0.parse_next(input)?;

        let name = take_till(1.., (b'\t', b' ', b'=')).parse_next(input)?;

        (space0, "=", space0).parse_next(input)?;

        let value =
            alt((string_parse, take_till(1.., (' ', '\t', '\'', '\r', '\n')))).parse_next(input)?;

        property_group
            .properties
            .insert(name.into(), Either::Left(value.into()));

        (space0, opt(line_comment_parse), line_ending).parse_next(input)?;
    }

    Ok(property_group)
}

fn property_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6FullyQualifiedName> {
    let Ok(namespace) = take_until::<_, _, VB6Error>(0.., ".").parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoNamespaceAfterBegin));
    };

    if literal::<&str, _, VB6Error>(".").parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoDotAfterNamespace));
    };

    let Ok(kind) = take_until::<_, _, VB6Error>(0.., (" ", "\t")).parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoUserControlNameAfterDot));
    };

    if space1::<_, VB6Error>.parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoSpaceAfterControlKind));
    }

    let Ok(name) =
        take_till::<_, _, VB6Error>(0.., (b" ", b"\t", b"\r", b"\r\n", b"\n")).parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoControlNameAfterControlKind));
    };

    // If there are spaces after the control name, eat those up since we don't care about them.
    space0.parse_next(input)?;
    // eat the line ending and move on.
    line_ending.parse_next(input)?;

    Ok(VB6FullyQualifiedName {
        namespace: namespace.into(),
        kind: kind.into(),
        name: name.into(),
    })
}

fn build_control<'a>(
    fully_qualified_name: VB6FullyQualifiedName,
    controls: Vec<VB6Control>,
    menus: Vec<VB6Control>,
    properties: Properties<'a>,
    property_groups: Vec<VB6PropertyGroup>,
) -> Result<VB6Control, VB6ErrorKind> {
    let tag = match properties.get(b"Tag".into()) {
        Some(text) => text.into(),
        None => b"".into(),
    };

    if fully_qualified_name.namespace != "VB" {
        let custom_control = VB6Control {
            name: fully_qualified_name.name,
            tag: tag,
            index: 0,
            kind: VB6ControlKind::Custom {
                properties: properties.into(),
                property_groups,
            },
        };

        return Ok(custom_control);
    }

    let kind = match fully_qualified_name.kind.as_bytes() {
        b"Form" => {
            let mut converted_menus = vec![];

            for menu in menus {
                if let VB6ControlKind::Menu {
                    properties: menu_properties,
                    sub_menus,
                } = menu.kind
                {
                    converted_menus.push(VB6MenuControl {
                        name: menu.name,
                        tag: menu.tag,
                        index: menu.index,
                        properties: menu_properties,
                        sub_menus,
                    });
                }
            }

            VB6ControlKind::Form {
                controls,
                properties: properties.into(),
                menus: converted_menus,
            }
        }
        b"MDIForm" => {
            let mut converted_menus = vec![];

            for menu in menus {
                if let VB6ControlKind::Menu {
                    properties: menu_properties,
                    sub_menus,
                } = menu.kind
                {
                    converted_menus.push(VB6MenuControl {
                        name: menu.name,
                        tag: menu.tag,
                        index: menu.index,
                        properties: menu_properties,
                        sub_menus,
                    });
                }
            }

            VB6ControlKind::MDIForm {
                controls,
                properties: properties.into(),
                menus: converted_menus,
            }
        }
        b"Menu" => {
            let menu_properties = properties.into();
            let mut converted_menus = vec![];

            for menu in menus {
                if let VB6ControlKind::Menu {
                    properties: menu_properties,
                    sub_menus,
                } = menu.kind
                {
                    converted_menus.push(VB6MenuControl {
                        name: menu.name,
                        tag: menu.tag,
                        index: menu.index,
                        properties: menu_properties,
                        sub_menus,
                    });
                }
            }

            VB6ControlKind::Menu {
                properties: menu_properties,
                sub_menus: converted_menus,
            }
        }
        b"Frame" => VB6ControlKind::Frame {
            controls,
            properties: properties.into(),
        },
        b"CheckBox" => VB6ControlKind::CheckBox {
            properties: properties.into(),
        },
        b"ComboBox" => VB6ControlKind::ComboBox {
            properties: properties.into(),
        },
        b"CommandButton" => VB6ControlKind::CommandButton {
            properties: properties.into(),
        },
        b"Data" => VB6ControlKind::Data {
            properties: properties.into(),
        },
        b"DirListBox" => VB6ControlKind::DirListBox {
            properties: properties.into(),
        },
        b"DriveListBox" => VB6ControlKind::DriveListBox {
            properties: properties.into(),
        },
        b"FileListBox" => VB6ControlKind::FileListBox {
            properties: properties.into(),
        },
        b"Image" => VB6ControlKind::Image {
            properties: properties.into(),
        },
        b"Label" => VB6ControlKind::Label {
            properties: properties.into(),
        },
        b"Line" => VB6ControlKind::Line {
            properties: properties.into(),
        },
        b"ListBox" => VB6ControlKind::ListBox {
            properties: properties.into(),
        },
        b"OLE" => VB6ControlKind::Ole {
            properties: properties.into(),
        },
        b"OptionButton" => VB6ControlKind::OptionButton {
            properties: properties.into(),
        },
        b"PictureBox" => VB6ControlKind::PictureBox {
            properties: properties.into(),
        },
        b"HScrollBar" => VB6ControlKind::HScrollBar {
            properties: properties.into(),
        },
        b"VScrollBar" => VB6ControlKind::VScrollBar {
            properties: properties.into(),
        },
        b"Shape" => VB6ControlKind::Shape {
            properties: properties.into(),
        },
        b"TextBox" => VB6ControlKind::TextBox {
            properties: properties.into(),
        },
        b"Timer" => VB6ControlKind::Timer {
            properties: properties.into(),
        },
        _ => {
            return Err(VB6ErrorKind::UnknownControlKind);
        }
    };

    let parent_control = VB6Control {
        name: fully_qualified_name.name,
        tag: tag.clone().into(),
        index: 0,
        kind,
    };

    Ok(parent_control)
}

#[cfg(test)]
mod tests {

    use super::*;

    #[test]
    fn begin_property_valid() {
        let source = b"BeginProperty Font\r
                        Name = \"Arial\"\r
                        Size = 8.25\r
                        Charset = 0\r
                        Weight = 400\r
                        Underline = 0 'False\r
                        Italic = 0 'False\r
                        Strikethrough = 0 'False\r
                    EndProperty\r\n";

        let mut input = VB6Stream::new("", source);
        let result = property_group_parse.parse_next(&mut input);

        assert!(result.is_ok());

        let result = result.unwrap();
        assert_eq!(result.name, "Font");
        assert_eq!(result.properties.len(), 7);
    }

    #[test]
    fn nested_property_group() {
        use crate::language::VB6ControlKind;
        use crate::parsers::form::VB6FormFile;

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
   Icon            =   \"Form_Main.frx\":0000\r
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
            Picture         =   \"Form_Main.frx\":17789\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":1921B\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":1ACAD\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":1C73F\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":1E1D1\r
            Key             =   \"\"\r
         EndProperty\r
      EndProperty\r
   End\r
End\r
Attribute VB_Name = \"Form_Main\"\r
";

        let result = VB6FormFile::parse(
            "form_parse.frm".to_owned(),
            input.as_bytes(),
            resource_file_resolver,
        )
        .unwrap();

        assert_eq!(result.objects.len(), 1);
        assert_eq!(result.format_version.major, 5);
        assert_eq!(result.format_version.minor, 0);
        assert_eq!(result.form.name, "Form_Main");
        assert_eq!(
            matches!(result.form.kind, VB6ControlKind::Form { .. }),
            true
        );

        if let VB6ControlKind::Form {
            controls,
            properties,
            menus,
        } = &result.form.kind
        {
            assert_eq!(controls.len(), 1);
            assert_eq!(menus.len(), 0);
            assert_eq!(properties.caption, "Audiostation");
            assert_eq!(controls[0].name, "Imagelist_CDDisplay");
            assert!(matches!(controls[0].kind, VB6ControlKind::Custom { .. }));

            if let VB6ControlKind::Custom {
                properties,
                property_groups,
            } = &controls[0].kind
            {
                assert_eq!(properties.len(), 9);
                assert_eq!(property_groups.len(), 1);

                if let Some(group) = property_groups.get(0) {
                    assert_eq!(group.name, "Images");
                    assert_eq!(group.properties.len(), 6);

                    if let Some(Either::Right(image1)) =
                        group.properties.get(BStr::new("ListImage1"))
                    {
                        assert_eq!(image1.name, BStr::new("ListImage1"));
                        assert_eq!(image1.properties.len(), 2);
                    } else {
                        panic!("Expected nested ListImage1");
                    }

                    if let Some(Either::Right(image2)) =
                        group.properties.get(BStr::new("ListImage2"))
                    {
                        assert_eq!(image2.name, BStr::new("ListImage2"));
                        assert_eq!(image2.properties.len(), 2);
                    } else {
                        panic!("Expected nested ListImage2");
                    }

                    if let Some(Either::Right(image3)) =
                        group.properties.get(BStr::new("ListImage3"))
                    {
                        assert_eq!(image3.name, BStr::new("ListImage3"));
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
    fn parse_english_code_non_english_text() {
        use crate::errors::VB6ErrorKind;
        use crate::parsers::form::VB6FormFile;

        let input = "VERSION 5.00\r
Object = \"{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0\"; \"msscript.ocx\"\r
Object = \"{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0\"; \"COMDLG32.OCX\"\r
Object = \"{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0\"; \"Imagex.ocx\"\r
Begin VB.Form FormMainMode \r
   BorderStyle     =   1  '��u�T�w\r
   Caption         =   \"UnlightVBE-QS Origin\"\r
   ClientHeight    =   11100\r
   ClientLeft      =   45\r
   ClientTop       =   375\r
   ClientWidth     =   20400\r
   BeginProperty Font \r
      Name            =   \"�L�n������\"\r
      Size            =   12\r
      Charset         =   136\r
      Weight          =   400\r
      Underline       =   0   'False\r
      Italic          =   0   'False\r
      Strikethrough   =   0   'False\r
   EndProperty\r
   Icon            =   \"FormMainMode.frx\":0000\r
   LinkTopic       =   \"Form3\"\r
   MaxButton       =   0   'False\r
   ScaleHeight     =   11100\r
   ScaleWidth      =   20400\r
   StartUpPosition =   2  '�ù�����\r
   Tag             =   \"UnlightVBE-QS Origin\"\r
   Begin VB.CommandButton �v�l�]�w \r
         Caption         =   \"�v�l�]�w\"\r
         BeginProperty Font \r
            Name            =   \"�L�n������\"\r
            Size            =   8.25\r
            Charset         =   136\r
            Weight          =   400\r
            Underline       =   0   'False\r
            Italic          =   0   'False\r
            Strikethrough   =   0   'False\r
         EndProperty\r
         Height          =   405\r
         Left            =   8760\r
         TabIndex        =   124\r
         Top             =   9360\r
         Width           =   975\r
   End\r
End\r
Attribute VB_Name = \"FormMainMode\"\r
";

        let result = VB6FormFile::parse(
            "form_parse.frm".to_owned(),
            input.as_bytes(),
            resource_file_resolver,
        );

        assert!(matches!(
            result.err().unwrap().kind,
            VB6ErrorKind::LikelyNonEnglishCharacterSet
        ));
    }

    #[test]
    fn parse_indented_menu_valid() {
        use crate::language::MenuProperties;
        use crate::language::VB_WINDOW_BACKGROUND;

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

        let result = VB6FormFile::parse(
            "form_parse.frm".to_owned(),
            &mut input.as_ref(),
            resource_file_resolver,
        );

        assert!(result.is_ok());

        let result = result.unwrap();

        assert_eq!(result.format_version.major, 5);
        assert_eq!(result.format_version.minor, 0);

        assert_eq!(result.form.name, "frmExampleForm");
        assert_eq!(result.form.tag, "");
        assert_eq!(result.form.index, 0);

        if let VB6ControlKind::Form {
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
                &vec![VB6MenuControl {
                    name: "mnuFile".into(),
                    tag: "".into(),
                    index: 0,
                    properties: MenuProperties {
                        caption: "&File".into(),
                        ..Default::default()
                    },
                    sub_menus: vec![VB6MenuControl {
                        name: "mnuOpenImage".into(),
                        tag: "".into(),
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
}
