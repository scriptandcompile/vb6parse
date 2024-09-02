#![warn(clippy::pedantic)]

use bstr::BStr;

use crate::{
    errors::VB6ErrorKind,
    parsers::VB6Stream,
    vb6::{keyword_parse, line_comment_parse, VB6Result},
};

use winnow::{
    ascii::{digit1, line_ending, space0, space1},
    combinator::{alt, delimited, eof, opt, separated_pair},
    error::ErrMode,
    token::{literal, take_while},
    Parser,
};

/// Represents a VB6 file format version.
/// A VB6 file format version contains a major version number and a minor version number.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub struct VB6FileFormatVersion {
    pub major: u8,
    pub minor: u8,
}

pub enum HeaderKind {
    Class,
    Form,
    //UserControl,
}

pub fn version_parse<'a>(
    header_kind: HeaderKind,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<VB6FileFormatVersion> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<VB6FileFormatVersion> {
        (space0, keyword_parse("VERSION"), space1).parse_next(input)?;

        let Ok(major_digits): VB6Result<&'a BStr> = digit1.parse_next(input) else {
            return Err(ErrMode::Cut(VB6ErrorKind::MajorVersionUnparseable));
        };

        let Ok(major_version) = major_digits.to_string().as_str().parse::<u8>() else {
            return Err(ErrMode::Cut(VB6ErrorKind::MajorVersionUnparseable));
        };

        ".".parse_next(input)?;

        let Ok(minor_digits): VB6Result<&'a BStr> = digit1.parse_next(input) else {
            return Err(ErrMode::Cut(VB6ErrorKind::MinorVersionUnparseable));
        };

        let Ok(minor_version) = minor_digits.to_string().as_str().parse::<u8>() else {
            return Err(ErrMode::Cut(VB6ErrorKind::MinorVersionUnparseable));
        };

        match header_kind {
            HeaderKind::Class => {
                space1.parse_next(input)?;
                keyword_parse("CLASS").parse_next(input)?;
            }
            HeaderKind::Form => {
                // Form headers only have the version keyword and the
                // major/minor version numbers.
                // There is no 'FORM' keyword in the version line.
            } //HeaderKind::UserControl => {
              // User Control headers only have the version keyword and the
              // major/minor version numbers.
              // There is no 'UserControl' keyword in the version line.
              //}
        }

        (space0, line_ending).parse_next(input)?;

        Ok(VB6FileFormatVersion {
            major: major_version,
            minor: minor_version,
        })
    }
}

pub fn key_value_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<(&'a BStr, &'a BStr)> {
        let (key, value) = separated_pair(
            delimited(
                space0,
                take_while(1.., ('_', '-', '+', '&', 'a'..='z', 'A'..='Z', '0'..='9')),
                space0,
            ),
            literal(divider),
            alt((
                delimited(
                    (space0, "\""),
                    take_while(1.., (' '..='!', '#'..='~', '\t')),
                    ("\"", space0),
                ),
                delimited(
                    space0,
                    take_while(
                        1..,
                        ('_', '.', '-', '+', '&', 'a'..='z', 'A'..='Z', '0'..='9'),
                    ),
                    space0,
                ),
            )),
        )
        .parse_next(input)?;

        Ok((key, value))
    }
}

pub fn key_value_line_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<(&'a BStr, &'a BStr)> {
        let (key, value) = key_value_parse(divider).parse_next(input)?;

        // we have to check for eof here because it's perfectly possible to have a
        // header file that is empty of actual code. This means the last line of the file
        // should be an empty line, but it might be that the filed ends at the end of the
        // header attribute section.
        (space0, opt(line_comment_parse), alt((line_ending, eof))).parse_next(input)?;

        Ok((key, value))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parsers::VB6Stream;

    use super::HeaderKind;

    #[test]
    fn test_class_version_parse() {
        let mut stream = VB6Stream::new("", b"VERSION 1.0 CLASS\r\n");
        let version = version_parse(HeaderKind::Class)(&mut stream).unwrap();

        assert_eq!(version.major, 1);
        assert_eq!(version.minor, 0);
    }

    #[test]
    fn test_form_version_parse() {
        let mut stream = VB6Stream::new("", b"VERSION 5.00\r\n");
        let version = version_parse(HeaderKind::Form)(&mut stream).unwrap();

        assert_eq!(version.major, 5);
        assert_eq!(version.minor, 0);
    }

    #[test]
    fn test_key_value_parse() {
        let mut stream = VB6Stream::new("", b"Attribute1 = Value1\r\n");
        let (key, value) = key_value_parse("=")(&mut stream).unwrap();

        assert_eq!(key, "Attribute1".as_bytes());
        assert_eq!(value, "Value1".as_bytes());
    }

    #[test]
    fn test_key_value_line_parse() {
        let mut stream = VB6Stream::new("", b"Attribute1 = Value1\r\n");
        let (key, value) = key_value_line_parse("=")(&mut stream).unwrap();

        assert_eq!(key, "Attribute1".as_bytes());
        assert_eq!(value, "Value1".as_bytes());
    }
}
