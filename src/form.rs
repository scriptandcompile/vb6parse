#![warn(clippy::pedantic)]

use std::vec;

use crate::vb6::{line_comment_parse, vb6_parse, VB6Token};
use crate::vb6stream::VB6Stream;
use crate::VB6FileFormatVersion;

use bstr::{BStr, ByteSlice};

use winnow::error::ParserError;
use winnow::{
    ascii::{digit1, line_ending, space0, space1, Caseless},
    combinator::{alt, delimited, opt},
    error::{ContextError, StrContext},
    token::{take_till, take_until},
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
    /// let input = b"VERSION 5.00
    /// Begin VB.Form frmExampleForm
    ///    BackColor       =   &H80000005&
    ///    Caption         =   \"example form\"
    ///    ClientHeight    =   6210
    ///    ClientLeft      =   60
    ///    ClientTop       =   645
    ///    ClientWidth     =   9900
    ///    BeginProperty Font
    ///       Name            =   \"Arial\"
    ///       Size            =   8.25
    ///       Charset         =   0
    ///       Weight          =   400
    ///       Underline       =   0   'False
    ///       Italic          =   0   'False
    ///       Strikethrough   =   0   'False
    ///    EndProperty
    ///    LinkTopic       =   \"Form1\"
    ///    ScaleHeight     =   414
    ///    ScaleMode       =   3  'Pixel
    ///    ScaleWidth      =   660
    ///    StartUpPosition =   2  'CenterScreen
    ///    Begin VB.Menu mnuFile
    ///       Caption         =   \"&File\"
    ///       Begin VB.Menu mnuOpenImage
    ///          Caption         =   \"&Open image\"
    ///       End
    ///    End
    /// End
    /// ";
    ///
    /// let result = VB6FormFile::parse(&mut input.as_ref());
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(input: &'a [u8]) -> PResult<Self> {
        let mut input = VB6Stream::new(input);

        let format_version = version_information_parse.parse_next(&mut input)?;

        (space0, Caseless("BEGIN"), space1)
            .context(StrContext::Label("Expected 'Begin' keyword"))
            .parse_next(&mut input)?;

        let form = block_parse.parse_next(&mut input)?;

        let tokens = vb6_parse.parse_next(&mut input)?;

        Ok(VB6FormFile {
            form,
            format_version,
            tokens,
        })
    }
}

fn version_information_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6FileFormatVersion> {
    (space0, Caseless("VERSION"), space1)
        .context(StrContext::Label(
            "Version information not found at the start of the VB6 form file",
        ))
        .parse_next(input)?;

    let major_digits = digit1
        .context(StrContext::Label("Expected major version number"))
        .parse_next(input)?;

    let major_version_number =
        u8::from_str_radix(bstr::BStr::new(major_digits).to_string().as_str(), 10).unwrap();

    ".".context(StrContext::Label("Expected '.' after major version number"))
        .parse_next(input)?;

    let minor_digits = digit1
        .context(StrContext::Label("Expected minor version number"))
        .parse_next(input)?;

    let minor_version_number =
        u8::from_str_radix(bstr::BStr::new(minor_digits).to_string().as_str(), 10).unwrap();

    opt(line_comment_parse).parse_next(input)?;

    line_ending
        .context(StrContext::Label(
            "Expected line ending after version information",
        ))
        .parse_next(input)?;

    Ok(VB6FileFormatVersion {
        major: major_version_number,
        minor: minor_version_number,
    })
}

fn block_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6Control<'a>> {
    let fully_qualified_name = begin_parse.parse_next(input)?;

    let mut controls = vec![];
    let mut property_groups = vec![];
    let mut properties: Vec<VB6Property> = vec![];

    while !input.is_empty() {
        space0.parse_next(input)?;

        if (
            space0::<VB6Stream<'a>, ContextError>,
            Caseless("BeginProperty"),
            space1,
        )
            .parse_next(input)
            .is_ok()
        {
            let property_group = begin_property_parse.parse_next(input)?;

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

        let (name, value) = key_value_pair_parse.parse_next(input)?;

        properties.push(VB6Property { name, value });
    }

    Err(ParserError::assert(input, "Unknown control kind"))
}

fn begin_property_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6PropertyGroup<'a>> {
    let property_name = take_till(1.., (b"\r", b"\t", b" ", b"\n"))
        .context(StrContext::Label(
            "Expected property name after 'BeginProperty' keyword",
        ))
        .parse_next(input)?;

    space0.parse_next(input)?;

    opt(line_comment_parse).parse_next(input)?;

    line_ending
        .context(StrContext::Label(
            "Expected line ending after property name",
        ))
        .parse_next(input)?;

    let mut property_group = VB6PropertyGroup {
        name: property_name.as_bstr(),
        properties: vec![],
    };

    while !input.is_empty() {
        if (
            space0::<VB6Stream<'a>, ContextError>,
            Caseless("EndProperty"),
            space0,
        )
            .context(StrContext::Label(
                "Expected 'EndProperty' keyword to end property group",
            ))
            .parse_next(input)
            .is_ok()
        {
            break;
        }

        let (name, value) = key_value_pair_parse(input)?;
        property_group.properties.push(VB6Property { name, value });
    }

    line_ending
        .context(StrContext::Label(
            "Expected line ending after EndProperty keyword.",
        ))
        .parse_next(input)?;

    Ok(property_group)
}

// TODO: it looks like I can break some of this out into a module
// specifically for parsing VB6 header information since these
// headers are basically a language of their own and shared between
// the different VB6 file types.
// this should apply to:
// quoted_value_parse, and unqouted_value_parse, the Begin/End blocks,
// and the BeginProperty/EndProperty blocks.

/// Parses a qouted-value from a byte slice.
/// The qouted value excludes the double qoutes.
fn quoted_value_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<&'a BStr> {
    let value = delimited('\"', take_till(0.., '\"'), '\"')
        .context(StrContext::Label("Expected quoted value"))
        .parse_next(input)?;

    Ok(value.as_bstr())
}

/// Parses an unquoted-value from a byte slice.
fn unqouted_value_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<&'a BStr> {
    let value = take_till(0.., (b"\r\n", b"\n", b"'", b" "))
        .context(StrContext::Label("Expected unquoted value"))
        .parse_next(input)?;

    Ok(value.as_bstr())
}

fn key_value_pair_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<(&'a BStr, &'a BStr)> {
    space0.parse_next(input)?;

    let key = take_till(0.., (b" ", b"="))
        .context(StrContext::Label("Expected key before '='"))
        .parse_next(input)?;

    space0.parse_next(input)?;

    "=".context(StrContext::Label("Expected '=' after key"))
        .parse_next(input)?;

    space0.parse_next(input)?;

    let value = alt((quoted_value_parse, unqouted_value_parse)).parse_next(input)?;

    space0.parse_next(input)?;

    opt(line_comment_parse).parse_next(input)?;

    line_ending
        .context(StrContext::Label("Expected line ending after value"))
        .parse_next(input)?;

    Ok((key.as_bstr(), value))
}

fn begin_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6FullyQualifiedName<'a>> {
    let namespace = take_until(0.., ".")
        .context(StrContext::Label(
            "Expected namespace after 'Begin' keyword",
        ))
        .parse_next(input)?;

    ".".context(StrContext::Label("Expected '.' after namespace"))
        .parse_next(input)?;

    let kind = take_until(0.., (" ", "\t"))
        .context(StrContext::Label("Expected control kind after '.'"))
        .parse_next(input)?;

    space1.parse_next(input)?;

    let name = take_until(0.., "\n")
        .context(StrContext::Label(
            "Expected control name after control kind",
        ))
        .parse_next(input)?;

    line_ending
        .context(StrContext::Label("Expected line ending after control name"))
        .parse_next(input)?;

    Ok(VB6FullyQualifiedName {
        namespace: namespace.as_bstr(),
        kind: kind.as_bstr(),
        name: name.as_bstr(),
    })
}
