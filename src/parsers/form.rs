#![warn(clippy::pedantic)]

use std::collections::HashMap;
use std::vec::Vec;

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    language::{
        CheckBoxProperties, ComboBoxProperties, CommandButtonProperties, FormProperties,
        FrameProperties, LabelProperties, LineProperties, MenuProperties, PictureBoxProperties,
        ScrollBarProperties, TextBoxProperties, VB6Color, VB6Control, VB6ControlKind,
        VB6MenuControl, VB6Token,
    },
    parsers::{
        header::{key_value_line_parse, version_parse, HeaderKind, VB6FileFormatVersion},
        VB6Stream,
    },
    vb6::{keyword_parse, line_comment_parse, vb6_parse, VB6Result},
};

use bstr::ByteSlice;

use winnow::error::ParserError;
use winnow::{
    ascii::{line_ending, space0, space1},
    combinator::opt,
    error::ErrMode,
    token::{literal, take_till, take_until},
    Parser,
};

/// Represents a VB6 Form file.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6FormFile<'a> {
    pub form: VB6Control<'a>,
    pub format_version: VB6FileFormatVersion,
    pub tokens: Vec<VB6Token<'a>>,
}

#[derive(Debug, PartialEq, Eq, Clone)]
struct VB6FullyQualifiedName<'a> {
    pub namespace: &'a str,
    pub kind: &'a str,
    pub name: &'a str,
}

#[derive(Debug, PartialEq, Eq, Clone)]
struct VB6PropertyGroup<'a> {
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
    /// //assert!(result.is_ok());
    /// ```
    pub fn parse(file_name: String, input: &'a [u8]) -> Result<Self, VB6Error> {
        let mut input = VB6Stream::new(file_name, input);

        let format_version = match version_parse(HeaderKind::Form).parse_next(&mut input) {
            Ok(version) => version,
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
            format_version,
            tokens,
        })
    }
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
        } else if (space0, keyword_parse("BEGIN"), space1)
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
        } else if (space0, keyword_parse("END"), space0, line_ending)
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
        } else {
            let (name, value) = key_value_line_parse("=").parse_next(input)?;

            let Ok(name_ascii) = name.to_str() else {
                return Err(ErrMode::Cut(VB6ErrorKind::PropertyNameAsciiConversionError));
            };

            let Ok(value_ascii) = value.to_str() else {
                return Err(ErrMode::Cut(
                    VB6ErrorKind::PropertyValueAsciiConversionError,
                ));
            };

            properties.insert(name_ascii, value_ascii);
        }
    }

    Err(ParserError::assert(input, "Unknown control kind"))
}

fn build_control<'a>(
    fully_qualified_name: VB6FullyQualifiedName<'a>,
    controls: Vec<VB6Control<'a>>,
    menus: Vec<VB6Control<'a>>,
    properties: HashMap<&'a str, &'a str>,
    _property_groups: Vec<VB6PropertyGroup<'a>>,
) -> Result<VB6Control<'a>, VB6ErrorKind> {
    // This is wrong.
    // TODO: When we start work on custom controls we will need
    // to handle fully verified name parsing. This will work for now though.
    let kind = match fully_qualified_name.kind.as_bytes() {
        b"Form" => {
            let mut form_properties = FormProperties::default();

            if properties.contains_key("Caption") {
                form_properties.caption = properties["Caption"];
            }

            if properties.contains_key("BackColor") {
                let color_ascii = properties["BackColor"];

                let Ok(back_color) = VB6Color::from_hex(color_ascii) else {
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

            let form = VB6ControlKind::Form {
                controls,
                properties: form_properties,
                menus: converted_menus,
            };

            form
        }
        b"Menu" => {
            let mut menu_properties = MenuProperties::default();

            if properties.contains_key("Caption") {
                menu_properties.caption = properties["Caption"];
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

            let menu = VB6ControlKind::Menu {
                properties: menu_properties,
                sub_menus: converted_menus,
            };
            menu
        }
        b"Frame" => {
            let frame = VB6ControlKind::Frame {
                controls,
                properties: FrameProperties::default(),
            };
            frame
        }
        b"TextBox" => {
            let textbox = VB6ControlKind::TextBox {
                properties: TextBoxProperties::default(),
            };
            textbox
        }
        b"CheckBox" => {
            let checkbox = VB6ControlKind::CheckBox {
                properties: CheckBoxProperties::default(),
            };
            checkbox
        }
        b"Line" => {
            let line = VB6ControlKind::Line {
                properties: LineProperties::default(),
            };
            line
        }
        b"Label" => {
            let label = VB6ControlKind::Label {
                properties: LabelProperties::default(),
            };
            label
        }
        b"ComboBox" => {
            let combobox = VB6ControlKind::ComboBox {
                properties: ComboBoxProperties::default(),
            };
            combobox
        }
        b"CommandButton" => {
            let commandbutton = VB6ControlKind::CommandButton {
                properties: CommandButtonProperties::default(),
            };
            commandbutton
        }
        b"PictureBox" => {
            let picturebox = VB6ControlKind::PictureBox {
                properties: PictureBoxProperties::default(),
            };
            picturebox
        }
        b"HScrollBar" => {
            let hscrollbar = VB6ControlKind::HScrollBar {
                properties: ScrollBarProperties::default(),
            };
            hscrollbar
        }
        b"VScrollBar" => {
            let vscrollbar = VB6ControlKind::VScrollBar {
                properties: ScrollBarProperties::default(),
            };
            vscrollbar
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

    let name = match take_till::<(u8, u8, u8, u8), _, VB6Error>(1.., (b'\r', b'\t', b' ', b'\n'))
        .parse_next(input)
    {
        Ok(name) => name,
        Err(_) => {
            return Err(ErrMode::Cut(VB6ErrorKind::NoPropertyName));
        }
    };

    let Ok(property_name) = name.to_str() else {
        return Err(ErrMode::Cut(VB6ErrorKind::PropertyNameAsciiConversionError));
    };

    space0.parse_next(input)?;

    opt(line_comment_parse).parse_next(input)?;

    if line_ending::<_, VB6Error>.parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    let mut property_group = VB6PropertyGroup {
        name: property_name,
        properties: HashMap::new(),
    };

    while !input.is_empty() {
        if (space0, keyword_parse("EndProperty"), space0)
            .parse_next(input)
            .is_ok()
        {
            break;
        }

        let (name, value) = key_value_line_parse("=").parse_next(input)?;

        let Ok(name_ascii) = name.to_str() else {
            return Err(ErrMode::Cut(VB6ErrorKind::PropertyNameAsciiConversionError));
        };

        let Ok(value_ascii) = value.to_str() else {
            return Err(ErrMode::Cut(
                VB6ErrorKind::PropertyValueAsciiConversionError,
            ));
        };
        property_group.properties.insert(name_ascii, value_ascii);
    }

    if line_ending::<_, VB6Error>.parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEndingAfterEndProperty));
    }

    Ok(property_group)
}

fn begin_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6FullyQualifiedName<'a>> {
    let namespace = match take_until::<_, _, VB6Error>(0.., ".").parse_next(input) {
        Ok(namespace) => namespace,
        Err(_) => {
            return Err(ErrMode::Cut(VB6ErrorKind::NoNamespaceAfterBegin));
        }
    };

    if literal::<&str, _, VB6Error>(".").parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoDotAfterNamespace));
    };

    let kind = match take_until::<_, _, VB6Error>(0.., (" ", "\t")).parse_next(input) {
        Ok(kind) => kind,
        Err(_) => {
            return Err(ErrMode::Cut(VB6ErrorKind::NoUserControlNameAfterDot));
        }
    };

    if space1::<_, VB6Error>.parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoSpaceAfterControlKind));
    }

    let name = match take_till::<_, _, VB6Error>(0.., (b" ", b"\t", b"\r", b"\r\n", b"\n"))
        .parse_next(input)
    {
        Ok(name) => name,
        Err(_) => {
            return Err(ErrMode::Cut(VB6ErrorKind::NoControlNameAfterControlKind));
        }
    };

    let Ok(namespace_ascii) = namespace.to_str() else {
        return Err(ErrMode::Cut(VB6ErrorKind::NamespaceAsciiConversionError));
    };

    let Ok(kind_ascii) = kind.to_str() else {
        return Err(ErrMode::Cut(VB6ErrorKind::ControlKindAsciiConversionError));
    };

    let Ok(name_ascii) = name.to_str() else {
        return Err(ErrMode::Cut(
            VB6ErrorKind::QualifiedControlNameAsciiConversionError,
        ));
    };

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
        let _result = begin_property_parse.parse_next(&mut input);

        //assert!(result.is_ok());

        //let result = result.unwrap();
        //assert_eq!(result.name, "Font");
        //assert_eq!(result.properties.len(), 7);
    }

    #[test]
    fn larger_parse_valid() {
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
                        caption: &"&File",
                        ..Default::default()
                    },
                    sub_menus: vec![VB6MenuControl {
                        name: "mnuOpenImage",
                        tag: "",
                        index: 0,
                        properties: MenuProperties {
                            caption: &"&Open image",
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
    fn parse_valid() {
        use crate::language::VB_WINDOW_BACKGROUND;

        let source = b"VERSION 5.00\r
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

        let _result = VB6FormFile::parse("form_parse.frm".to_owned(), source);
    }
}
