#![warn(clippy::pedantic)]

use std::vec;

use crate::vb6::{eol_comment_parse, vb6_parse, VB6Token};
use crate::VB6FileFormatVersion;

use bstr::{BStr, ByteSlice};

use winnow::{
    ascii::{digit1, line_ending, space0, space1, Caseless},
    combinator::opt,
    error::{ContextError, ErrMode, ParserError, StrContext},
    token::{literal, take_until},
    PResult, Parser,
};

/// Represents a VB6 Form file.
/// A VB6 Form file contains a header and a list of controls.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6FormFile<'a> {
    pub name: &'a BStr,
    pub format_version: VB6FileFormatVersion,
    pub controls: Vec<VB6Control<'a>>,
    pub tokens: Vec<VB6Token<'a>>,
}

/// Represents a VB6 control.
/// A VB6 control contains a name and a kind.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Control<'a> {
    pub name: &'a BStr,
    pub kind: VB6ControlKind<'a>,
}

/// Represents a VB6 control kind.
/// A VB6 control kind is an enumeration of the different kinds of
/// standard VB6 controls.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum VB6ControlKind<'a> {
    CommandButton { caption: &'a BStr },
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
    pub fn parse(input: &mut &'a [u8]) -> PResult<Self> {
        let format_version = version_information_parse(input)?;

        Ok(Self {
            name: b"".as_bstr(),
            format_version,
            controls: vec![],
            tokens: vec![],
        })
    }
}

fn version_information_parse<'a>(input: &mut &'a [u8]) -> PResult<VB6FileFormatVersion> {
    (space0, "VERSION", space1)
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

    opt(eol_comment_parse).parse_next(input)?;

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
