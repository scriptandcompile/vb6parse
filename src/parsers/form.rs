use std::collections::HashMap;
use std::vec::Vec;

use bstr::BStr;

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    language::{
        CheckBoxProperties, ComboBoxProperties, CommandButtonProperties, DataProperties,
        DirListBoxProperties, FormProperties, FrameProperties, ImageProperties, LabelProperties,
        LineProperties, ListBoxProperties, MenuProperties, OLEProperties, OptionButtonProperties,
        PictureBoxProperties, ScrollBarProperties, ShapeProperties, TextBoxProperties,
        TimerProperties, VB6Color, VB6Control, VB6ControlKind, VB6MenuControl, VB6Token,
    },
    parsers::{
        header::{key_resource_offset_line_parse, version_parse, HeaderKind, VB6FileFormatVersion},
        VB6ObjectReference, VB6Stream,
    },
    vb6::{keyword_parse, line_comment_parse, vb6_parse, VB6Result},
};

use bstr::ByteSlice;
use serde::Serialize;

use winnow::error::ParserError;
use winnow::{
    ascii::{line_ending, space0, space1},
    combinator::{alt, opt},
    error::ErrMode,
    token::{literal, take_till, take_until},
    Parser,
};

use super::{header::object_parse, vb6::string_parse};

/// Represents a VB6 Form file.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct VB6FormFile<'a> {
    pub form: VB6Control<'a>,
    pub objects: Vec<VB6ObjectReference<'a>>,
    pub format_version: VB6FileFormatVersion,
    pub tokens: Vec<VB6Token<'a>>,
}

#[derive(Debug, PartialEq, Eq, Clone, Copy, Serialize)]
struct VB6FullyQualifiedName<'a> {
    pub namespace: &'a str,
    pub kind: &'a str,
    pub name: &'a str,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6PropertyGroup<'a> {
    pub name: &'a str,
    pub properties: HashMap<&'a str, &'a str>,
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
    /// ";
    ///
    /// let result = VB6FormFile::parse("form_parse.frm".to_owned(), &mut input.as_ref());
    ///
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(file_name: String, input: &'a [u8]) -> Result<Self, VB6Error> {
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

        let form = match block_parse.parse_next(&mut input) {
            Ok(form) => form,
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
            tokens,
        })
    }
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

        (space0, "=", space0).parse_next(input)?;

        let object = object_parse.parse_next(input)?;

        line_ending.parse_next(input)?;

        objects.push(object);
    }

    Ok(objects)
}

fn block_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6Control<'a>> {
    let fully_qualified_name = begin_parse.parse_next(input)?;

    let mut controls = vec![];
    let mut menus = vec![];
    let mut property_groups = vec![];
    let mut properties = HashMap::new();

    while !input.is_empty() {
        space0.parse_next(input)?;

        if let Ok(property_group) = begin_property_parse.parse_next(input) {
            property_groups.push(property_group);
            continue;
        }

        if (space0, keyword_parse("BEGIN"), space1)
            .parse_next(input)
            .is_ok()
        {
            let control = block_parse.parse_next(input)?;
            if control.kind.is_menu() {
                menus.push(control);
            } else {
                controls.push(control);
            }

            continue;
        }

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

        if let Ok((name, _resource_file, _offset)) =
            key_resource_offset_line_parse("=").parse_next(input)
        {
            // TODO: At the moment we just eat the resource file look up.

            properties.insert(name, BStr::new(""));
            continue;
        }

        space0.parse_next(input)?;

        let name = take_until(1.., (" ", "\t", "=")).parse_next(input)?;

        space0.parse_next(input)?;

        "=".parse_next(input)?;

        space0.parse_next(input)?;

        let value =
            alt((string_parse, take_till(1.., (' ', '\t', '\'', '\r', '\n')))).parse_next(input)?;

        properties.insert(name, value);

        space0.parse_next(input)?;

        opt(line_comment_parse).parse_next(input)?;

        line_ending.parse_next(input)?;

        continue;
    }

    Err(ParserError::assert(input, "Unknown control kind"))
}

fn build_control<'a>(
    fully_qualified_name: VB6FullyQualifiedName<'a>,
    controls: Vec<VB6Control<'a>>,
    menus: Vec<VB6Control<'a>>,
    properties: HashMap<&'a BStr, &'a BStr>,
    _property_groups: Vec<VB6PropertyGroup<'a>>,
) -> Result<VB6Control<'a>, VB6ErrorKind> {
    // This is wrong.
    // TODO: When we start work on custom controls we will need
    // to handle fully verified name parsing. This will work for now though.
    let kind = match fully_qualified_name.kind.as_bytes() {
        b"Form" => {
            let mut form_properties = FormProperties::default();

            // TODO: We are not correctly handling property assignment for each control.
            let caption_key = BStr::new("Caption");
            if properties.contains_key(caption_key) {
                form_properties.caption = properties[caption_key];
            }

            let backcolor_key = BStr::new("BackColor");
            if properties.contains_key(backcolor_key) {
                let color_ascii = properties[backcolor_key];

                let Ok(back_color) = VB6Color::from_hex(color_ascii.to_str().unwrap()) else {
                    return Err(VB6ErrorKind::HexColorParseError);
                };
                form_properties.back_color = back_color;
            }

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
                properties: form_properties,
                menus: converted_menus,
            }
        }
        b"Menu" => {
            let mut menu_properties = MenuProperties::default();

            // TODO: We are not correctly handling property assignment for each control.
            let caption_key = BStr::new("Caption");
            if properties.contains_key(caption_key) {
                menu_properties.caption = properties[caption_key];
            }

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
        b"Frame" => {
            let frame_properties =
                FrameProperties::construct_control(properties, _property_groups)?;

            VB6ControlKind::Frame {
                controls,
                properties: frame_properties,
            }
        }
        b"TextBox" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::TextBox {
                properties: TextBoxProperties::default(),
            }
        }
        b"Timer" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::Timer {
                properties: TimerProperties::default(),
            }
        }
        b"CheckBox" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::CheckBox {
                properties: CheckBoxProperties::default(),
            }
        }
        b"Ole" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::Ole {
                properties: OLEProperties::default(),
            }
        }
        b"OptionButton" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::OptionButton {
                properties: OptionButtonProperties::default(),
            }
        }
        b"Line" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::Line {
                properties: LineProperties::default(),
            }
        }
        b"Shape" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::Shape {
                properties: ShapeProperties::default(),
            }
        }
        b"ListBox" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::ListBox {
                properties: ListBoxProperties::default(),
            }
        }
        b"Label" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::Label {
                properties: LabelProperties::default(),
            }
        }
        b"ComboBox" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::ComboBox {
                properties: ComboBoxProperties::default(),
            }
        }
        b"CommandButton" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::CommandButton {
                properties: CommandButtonProperties::default(),
            }
        }
        b"PictureBox" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::PictureBox {
                properties: PictureBoxProperties::default(),
            }
        }
        b"HScrollBar" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::HScrollBar {
                properties: ScrollBarProperties::default(),
            }
        }
        b"VScrollBar" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::VScrollBar {
                properties: ScrollBarProperties::default(),
            }
        }
        b"DirListBox" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::DirListBox {
                properties: DirListBoxProperties::default(),
            }
        }
        b"Image" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::Image {
                properties: ImageProperties::default(),
            }
        }
        b"Data" => {
            // TODO: We are not correctly handling property assignment for each control.
            VB6ControlKind::Data {
                properties: DataProperties::default(),
            }
        }
        _ => {
            return Err(VB6ErrorKind::UnknownControlKind);
        }
    };

    let parent_control = VB6Control {
        name: fully_qualified_name.name,
        tag: "",
        index: 0,
        kind,
    };

    Ok(parent_control)
}

fn begin_property_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6PropertyGroup<'a>> {
    (space0, keyword_parse("BeginProperty"), space1).parse_next(input)?;

    let Ok(name) = take_till::<(u8, u8, u8, u8), _, VB6Error>(1.., (b'\r', b'\t', b' ', b'\n'))
        .parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoPropertyName));
    };

    let Ok(property_name) = name.to_str() else {
        return Err(ErrMode::Cut(VB6ErrorKind::PropertyNameAsciiConversionError));
    };

    space0.parse_next(input)?;

    alt((line_comment_parse, line_ending)).parse_next(input)?;

    let mut property_group = VB6PropertyGroup {
        name: property_name,
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

        space0.parse_next(input)?;

        let name = take_until(1.., ("\t", " ", "=")).parse_next(input)?;

        let Ok(name_ascii) = name.to_str() else {
            return Err(ErrMode::Cut(VB6ErrorKind::PropertyNameAsciiConversionError));
        };

        space0.parse_next(input)?;

        "=".parse_next(input)?;

        space0.parse_next(input)?;

        let value =
            alt((string_parse, take_till(1.., (' ', '\t', '\'', '\r', '\n')))).parse_next(input)?;

        let Ok(value_ascii) = value.to_str() else {
            return Err(ErrMode::Cut(
                VB6ErrorKind::PropertyValueAsciiConversionError,
            ));
        };

        property_group.properties.insert(name_ascii, value_ascii);

        space0.parse_next(input)?;

        opt(line_comment_parse).parse_next(input)?;

        line_ending.parse_next(input)?;
    }

    Ok(property_group)
}

fn begin_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6FullyQualifiedName<'a>> {
    let Ok(namespace) = take_until::<_, _, VB6Error>(0.., ".").parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoNamespaceAfterBegin));
    };

    let Ok(namespace_ascii) = namespace.to_str() else {
        return Err(ErrMode::Cut(VB6ErrorKind::NamespaceAsciiConversionError));
    };

    if literal::<&str, _, VB6Error>(".").parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoDotAfterNamespace));
    };

    let Ok(kind) = take_until::<_, _, VB6Error>(0.., (" ", "\t")).parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoUserControlNameAfterDot));
    };

    let Ok(kind_ascii) = kind.to_str() else {
        return Err(ErrMode::Cut(VB6ErrorKind::ControlKindAsciiConversionError));
    };

    if space1::<_, VB6Error>.parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoSpaceAfterControlKind));
    }

    let Ok(name) =
        take_till::<_, _, VB6Error>(0.., (b" ", b"\t", b"\r", b"\r\n", b"\n")).parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoControlNameAfterControlKind));
    };
    let Ok(name_ascii) = name.to_str() else {
        return Err(ErrMode::Cut(
            VB6ErrorKind::QualifiedControlNameAsciiConversionError,
        ));
    };

    // If there are spaces after the control name, eat those up since we don't care about them.
    space0.parse_next(input)?;
    // eat the line ending and move on.
    line_ending.parse_next(input)?;

    Ok(VB6FullyQualifiedName {
        namespace: namespace_ascii,
        kind: kind_ascii,
        name: name_ascii,
    })
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
        let result = begin_property_parse.parse_next(&mut input);

        assert!(result.is_ok());

        let result = result.unwrap();
        assert_eq!(result.name, "Font");
        assert_eq!(result.properties.len(), 7);
    }

    #[test]
    fn parse_indented_menu_valid() {
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
    ";

        let result = VB6FormFile::parse("form_parse.frm".to_owned(), &mut input.as_ref());

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
                    name: "mnuFile",
                    tag: "",
                    index: 0,
                    properties: MenuProperties {
                        caption: BStr::new("&File"),
                        ..Default::default()
                    },
                    sub_menus: vec![VB6MenuControl {
                        name: "mnuOpenImage",
                        tag: "",
                        index: 0,
                        properties: MenuProperties {
                            caption: BStr::new("&Open image"),
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
