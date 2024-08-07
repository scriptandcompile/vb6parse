#![warn(clippy::pedantic)]

use std::vec;

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    header::{key_value_parse, version_parse, HeaderKind},
    vb6::{keyword_parse, line_comment_parse, vb6_parse, whitespace_parse, VB6Token},
    vb6stream::VB6Stream,
    VB6FileFormatVersion,
};

use bstr::{BStr, ByteSlice};

use winnow::error::ParserError;
use winnow::{
    ascii::{line_ending, space0, space1, Caseless},
    combinator::opt,
    error::{ContextError, ErrMode},
    token::{literal, take_till, take_until},
    PResult, Parser,
};

/// Represents a VB6 Form file.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6FormFile<'a> {
    pub form: VB6Control<'a>,
    pub format_version: VB6FileFormatVersion,
    pub tokens: Vec<VB6Token<'a>>,
}

/// Represents a VB6 control.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Control<'a> {
    pub common: VB6ControlCommonInformation<'a>,
    pub kind: VB6ControlKind<'a>,
}

/// Represents a VB6 control common information.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ControlCommonInformation<'a> {
    pub name: &'a BStr,
    pub caption: &'a BStr,
}

/// Represents a VB6 control kind.
/// A VB6 control kind is an enumeration of the different kinds of
/// standard VB6 controls.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum VB6ControlKind<'a> {
    CommandButton {},
    TextBox {},
    CheckBox {},
    Line {},
    Label {},
    Frame {},
    PictureBox {},
    ComboBox {},
    HScrollBar {},
    VScrollBar {},
    Menu {
        caption: &'a BStr,
        controls: Vec<VB6Control<'a>>,
    },
    Form {
        controls: Vec<VB6Control<'a>>,
    },
}

#[derive(Debug, PartialEq, Eq, Clone)]
struct VB6FullyQualifiedName<'a> {
    pub namespace: &'a BStr,
    pub kind: &'a BStr,
    pub name: &'a BStr,
}

#[derive(Debug, PartialEq, Eq, Clone)]
struct VB6PropertyGroup<'a> {
    pub name: &'a BStr,
    pub properties: Vec<VB6Property<'a>>,
}

#[derive(Debug, PartialEq, Eq, Clone)]
struct VB6Property<'a> {
    pub name: &'a BStr,
    pub value: &'a BStr,
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
    /// use vb6parse::form::VB6FormFile;
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
    /// println!("{:?}", result);
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(file_name: String, input: &'a [u8]) -> Result<Self, VB6Error> {
        let mut input = VB6Stream::new(file_name, input);

        let format_version = match version_parse(HeaderKind::Form).parse_next(&mut input) {
            Ok(version) => version,
            Err(err) => return Err(err.into_inner().unwrap()),
        };

        match (space0, keyword_parse("BEGIN"), space1).parse_next(&mut input) {
            Ok(_) => (),
            Err(err) => return Err(err.into_inner().unwrap()),
        };

        let form = match block_parse.parse_next(&mut input) {
            Ok(form) => form,
            Err(err) => return Err(err.into_inner().unwrap()),
        };

        let tokens = match vb6_parse.parse_next(&mut input) {
            Ok(tokens) => tokens,
            Err(err) => return Err(err.into_inner().unwrap()),
        };

        Ok(VB6FormFile {
            form,
            format_version,
            tokens,
        })
    }
}

fn block_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6Control<'a>, VB6Error> {
    let fully_qualified_name = begin_parse.parse_next(input)?;

    let mut controls = vec![];
    let mut property_groups = vec![];
    let mut properties: Vec<VB6Property> = vec![];

    while !input.is_empty() {
        space0.parse_next(input)?;

        if let Ok(property_group) = begin_property_parse.parse_next(input) {
            property_groups.push(property_group);
            continue;
        } else if (
            space0::<VB6Stream<'a>, ContextError>,
            Caseless("Begin"),
            space1,
        )
            .parse_next(input)
            .is_ok()
        {
            let control = block_parse.parse_next(input)?;
            controls.push(control);

            continue;
        } else if (
            space0::<VB6Stream<'a>, ContextError>,
            Caseless("End"),
            space0,
            line_ending,
        )
            .parse_next(input)
            .is_ok()
        {
            let kind = match fully_qualified_name.kind.as_bytes() {
                b"Form" => VB6ControlKind::Form { controls },
                b"Menu" => VB6ControlKind::Menu {
                    controls,
                    caption: properties[0].value,
                },
                b"TextBox" => VB6ControlKind::TextBox {},
                b"CheckBox" => VB6ControlKind::CheckBox {},
                b"Line" => VB6ControlKind::Line {},
                b"Label" => VB6ControlKind::Label {},
                b"Frame" => VB6ControlKind::Frame {},
                b"ComboBox" => VB6ControlKind::ComboBox {},
                b"CommandButton" => VB6ControlKind::CommandButton {},
                b"PictureBox" => VB6ControlKind::PictureBox {},
                b"HScrollBar" => VB6ControlKind::HScrollBar {},
                b"VScrollBar" => VB6ControlKind::HScrollBar {},
                _ => {
                    return Err(ParserError::assert(input, "Unknown control kind"));
                }
            };

            let parent_control = VB6Control {
                common: VB6ControlCommonInformation {
                    name: fully_qualified_name.name,
                    caption: fully_qualified_name.name,
                },
                kind,
            };

            return Ok(parent_control);
        }

        let (name, value) = key_value_parse("=").parse_next(input)?;

        properties.push(VB6Property { name, value });
    }

    Err(ParserError::assert(input, "Unknown control kind"))
}

fn begin_property_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6PropertyGroup<'a>, VB6Error> {
    (space0, keyword_parse("BeginProperty"), space1).parse_next(input)?;

    let property_name = match take_till::<(u8, u8, u8, u8), VB6Stream<'a>, VB6Error>(
        1..,
        (b'\r', b'\t', b' ', b'\n'),
    )
    .parse_next(input)
    {
        Ok(name) => name,
        Err(_) => {
            return Err(ErrMode::Cut(input.error(VB6ErrorKind::NoPropertyName)));
        }
    };

    space0.parse_next(input)?;

    opt(line_comment_parse).parse_next(input)?;

    if line_ending::<VB6Stream<'a>, VB6Error>
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(input.error(VB6ErrorKind::NoLineEnding)));
    }

    let mut property_group = VB6PropertyGroup {
        name: property_name.as_bstr(),
        properties: vec![],
    };

    while !input.is_empty() {
        if (space0, keyword_parse("EndProperty"), space0)
            .parse_next(input)
            .is_ok()
        {
            break;
        }

        let (name, value) = key_value_parse("=").parse_next(input)?;

        property_group.properties.push(VB6Property { name, value });

        opt(line_comment_parse).parse_next(input)?;

        if line_ending::<VB6Stream<'a>, VB6Error>
            .parse_next(input)
            .is_err()
        {
            return Err(ErrMode::Cut(input.error(VB6ErrorKind::NoLineEnding)));
        }
    }

    if line_ending::<VB6Stream<'a>, VB6Error>
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(
            input.error(VB6ErrorKind::NoLineEndingAfterEndProperty),
        ));
    }

    Ok(property_group)
}

fn begin_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6FullyQualifiedName<'a>, VB6Error> {
    let namespace = match take_until::<_, VB6Stream, VB6Error>(0.., ".").parse_next(input) {
        Ok(namespace) => namespace,
        Err(_) => {
            return Err(ErrMode::Cut(
                input.error(VB6ErrorKind::NoNamespaceAfterBegin),
            ));
        }
    };

    if literal::<&str, VB6Stream, VB6Error>(".")
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(input.error(VB6ErrorKind::NoDotAfterNamespace)));
    };

    let kind = match take_until::<_, VB6Stream, VB6Error>(0.., (" ", "\t")).parse_next(input) {
        Ok(kind) => kind,
        Err(_) => {
            return Err(ErrMode::Cut(
                input.error(VB6ErrorKind::NoUserControlNameAfterDot),
            ));
        }
    };

    if space1::<_, VB6Error>.parse_next(input).is_err() {
        return Err(ErrMode::Cut(
            input.error(VB6ErrorKind::NoSpaceAfterControlKind),
        ));
    }

    let name = match take_until::<_, VB6Stream, VB6Error>(0.., (" ", "\t", "\r\n", "\n"))
        .parse_next(input)
    {
        Ok(name) => name,
        Err(_) => {
            return Err(ErrMode::Cut(
                input.error(VB6ErrorKind::NoControlNameAfterControlKind),
            ));
        }
    };

    opt(whitespace_parse).parse_next(input)?;

    if line_ending::<_, VB6Error>.parse_next(input).is_err() {
        return Err(ErrMode::Cut(
            input.error(VB6ErrorKind::NoLineEndingAfterControlName),
        ));
    };

    Ok(VB6FullyQualifiedName {
        namespace: namespace.as_bstr(),
        kind: kind.as_bstr(),
        name: name.as_bstr(),
    })
}

#[cfg(test)]
mod tests {

    use miette::IntoDiagnostic;

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
    fn parse_valid() {
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

        let result = VB6FormFile::parse("form_parse.frm".to_owned(), source);
        let diag = result.into_diagnostic();

        println!("{:?}", diag);

        //assert!(result.is_ok());

        //let result = result.unwrap();

        //assert_eq!(result.format_version.major, 5);
        //assert_eq!(result.format_version.minor, 0);
    }
}
