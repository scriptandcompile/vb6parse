#![warn(clippy::pedantic)]

use bstr::BStr;
use miette::Result;

use crate::{
    vb6::{keyword_parse, line_comment_parse},
    vb6stream::VB6Stream,
    VB6FileFormatVersion,
};

use winnow::{
    ascii::{digit1, line_ending, space0, space1},
    combinator::{alt, delimited, eof, opt, separated_pair},
    error::{ContextError, ErrMode},
    error::{ParserError, StrContext, StrContextValue},
    token::{literal, take_while},
    PResult, Parser,
};

pub enum HeaderKind {
    Class,
    Form,
    UserControl,
}

pub fn version_parse<'a>(
    header_kind: HeaderKind,
) -> impl FnMut(&mut VB6Stream<'a>) -> PResult<VB6FileFormatVersion> {
    move |input: &mut VB6Stream<'a>| -> Result<VB6FileFormatVersion, ErrMode<ContextError>> {
        space0.parse_next(input)?;

        keyword_parse("VERSION")
            .context(StrContext::Expected(StrContextValue::Description(
                "'VERSION' header element not found.",
            )))
            .parse_next(input)?;

        space1
            .context(StrContext::Expected(StrContextValue::Description(
                "At least one space is required between the 'VERSION' header element and the header major version number.",
            ))).parse_next(input)?;

        let major_digits = digit1
            .context(StrContext::Expected(StrContextValue::Description(
                "Major version number not found.",
            )))
            .parse_next(input)?;

        let Ok(major_version) = bstr::BStr::new(major_digits)
            .to_string()
            .as_str()
            .parse::<u8>()
        else {
            let error = ParserError::assert(input, "Unable to parse major version number.");

            return Err(ErrMode::Cut(error));
        };

        ".".context(StrContext::Expected(StrContextValue::Description(
            "Version decimal character not found.",
        )))
        .parse_next(input)?;

        let minor_digits = digit1
            .context(StrContext::Expected(StrContextValue::Description(
                "Minor version number not found.",
            )))
            .parse_next(input)?;

        let Ok(minor_version) = bstr::BStr::new(minor_digits)
            .to_string()
            .as_str()
            .parse::<u8>()
        else {
            let error = ParserError::assert(input, "Unable to parse minor version number.");

            return Err(ErrMode::Cut(error));
        };

        match header_kind {
            HeaderKind::Class => {
                space1.context(
                    StrContext::Expected(StrContextValue::Description(
                        "At least one space is required between the header minor version number and the 'CLASS' header element"
                    ))).parse_next(input)?;
                keyword_parse("CLASS")
                    .context(StrContext::Expected(StrContextValue::Description(
                        "'CLASS' header element not found.",
                    )))
                    .parse_next(input)?;
            }
            HeaderKind::Form => {
                // Form headers only have the version keyword and the
                // major/minor version numbers.
                // There is no 'FORM' keyword in the version line.
            }
            HeaderKind::UserControl => {
                // User Control headers only have the version keyword and the
                // major/minor version numbers.
                // There is no 'UserControl' keyword in the version line.
            }
        }

        space0.parse_next(input)?;

        line_ending
            .context(StrContext::Expected(StrContextValue::Description(
                "Newline expected after version header element.",
            )))
            .parse_next(input)?;

        Ok(VB6FileFormatVersion {
            major: major_version,
            minor: minor_version,
        })
    }
}

pub fn key_value_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> PResult<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> Result<(&'a BStr, &'a BStr), ErrMode<ContextError>> {
        let (key, value) = separated_pair(
            delimited(
                space0,
                take_while(1.., ('_', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9')),
                space0,
            ),
            literal(divider),
            alt((
                delimited(
                    (space0, opt("\"")),
                    take_while(1.., ('_', '.', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9')),
                    (opt("\""), space0),
                ),
                delimited(
                    space0,
                    take_while(1.., ('_', '.', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9')),
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
) -> impl FnMut(&mut VB6Stream<'a>) -> PResult<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> PResult<(&'a BStr, &'a BStr)> {
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
    use crate::vb6stream::VB6Stream;

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
