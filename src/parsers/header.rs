use bstr::{BStr, ByteSlice};

use uuid::Uuid;

use crate::{
    errors::VB6ErrorKind,
    parsers::{VB6ObjectReference, VB6Stream},
    vb6::{keyword_parse, line_comment_parse, take_until_line_ending, VB6Result},
};

use winnow::{
    ascii::{digit1, line_ending, space0, space1},
    combinator::{alt, delimited, eof, opt, separated_pair},
    error::ErrMode,
    stream::Stream,
    token::{literal, take_till, take_until, take_while},
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

fn compiled_object_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ObjectReference<'a>> {
    // the GUID may or may not be wrapped in double-qoutes.
    opt("\"").parse_next(input)?;

    "{".parse_next(input)?;

    let uuid_segment = take_until(1.., "}").parse_next(input)?;

    let Ok(uuid) = Uuid::parse_str(uuid_segment.to_str().unwrap()) else {
        return Err(ErrMode::Cut(VB6ErrorKind::UnableToParseUuid));
    };

    "}#".parse_next(input)?;

    // still not sure what this element or the next represents.
    let version = take_until(1.., "#").parse_next(input)?;

    "#".parse_next(input)?;

    // we have to take until the next semi-colon or the next semi-colon wrapped in double-qoutes since it could be qouted or not.
    let unknown1 = alt((take_until(1.., ";"), take_until(1.., "\";"))).parse_next(input)?;

    opt("\"").parse_next(input)?;
    // the file name is preceded by a semi-colon then a space. not sure why the
    // space is there, but it is. this strips it and the semi-colon out.
    "; ".parse_next(input)?;

    // the filename may or may not be wrapped in double-qoutes.
    opt("\"").parse_next(input)?;

    // the filename is the rest of the input.
    // the filename may or may not be wrapped in double-qoutes.
    let file_name = alt((take_until_line_ending, take_until(1.., "\""))).parse_next(input)?;

    opt("\"").parse_next(input)?;

    let object = VB6ObjectReference::Compiled {
        uuid,
        version,
        unknown1,
        file_name,
    };

    Ok(object)
}

fn project_object_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ObjectReference<'a>> {
    // we have a qouted project path (likely a just a filename). aka Object = "*\\ADropStack.vbp"
    if (space0::<VB6Stream<'a>, VB6ErrorKind>, "\"*\\A")
        .parse_next(input)
        .is_ok()
    {
        let path = take_until(1.., "\"").parse_next(input)?;
        let object = VB6ObjectReference::Project { path };
        take_until_line_ending.parse_next(input)?;

        return Ok(object);
    }

    "*\\A".parse_next(input)?;

    // the path is the rest of the input.
    let path = take_until_line_ending.parse_next(input)?;

    let object = VB6ObjectReference::Project { path };

    Ok(object)
}

pub fn object_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ObjectReference<'a>> {
    alt((compiled_object_parse, project_object_parse)).parse_next(input)
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
                delimited(space0, vb6_string_parse, space0),
                delimited(
                    space0,
                    take_while(
                        1..,
                        (
                            '_',
                            '.',
                            '^',
                            '-',
                            '+',
                            '&',
                            '/',
                            ':',
                            'a'..='z',
                            'A'..='Z',
                            '0'..='9',
                        ),
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
        let checkpoint = input.checkpoint();

        let (key, value) = match key_value_parse(divider).parse_next(input) {
            Ok((key, value)) => (key, value),
            Err(e) => {
                input.reset(&checkpoint);
                return Err(e);
            }
        };

        // we have to check for eof here because it's perfectly possible to have a
        // header file that is empty of actual code. This means the last line of the file
        // should be an empty line, but it might be that the filed ends at the end of the
        // header attribute section.
        (space0, opt(line_comment_parse), alt((line_ending, eof))).parse_next(input)?;

        Ok((key, value))
    }
}

pub fn key_resource_offset_line_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<(&'a BStr, &'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<(&'a BStr, &'a BStr, &'a BStr)> {
        let checkpoint = input.checkpoint();

        let (key, resource_file_name) = match separated_pair(
            delimited(
                space0,
                take_while(1.., ('_', '-', '+', '&', 'a'..='z', 'A'..='Z', '0'..='9')),
                space0,
            ),
            literal(divider),
            delimited(
                (space0, opt("$"), "\""),
                take_while(1.., (' '..='!', '#'..='~', '\t')),
                "\"",
            ),
        )
        .parse_next(input)
        {
            Ok((key, resource_file_name)) => (key, resource_file_name),
            Err(e) => {
                input.reset(&checkpoint);
                return Err(e);
            }
        };

        let offset = match (":", take_while(1.., ('0'..='9', 'A'..='F'))).parse_next(input) {
            Ok((_, offset)) => offset,
            Err(e) => {
                input.reset(&checkpoint);
                return Err(e);
            }
        };

        // we have to check for eof here because it's perfectly possible to have a
        // header file that is empty of actual code. This means the last line of the file
        // should be an empty line, but it might be that the filed ends at the end of the
        // header attribute section.
        (space0, opt(line_comment_parse), alt((line_ending, eof))).parse_next(input)?;

        Ok((key, resource_file_name, offset))
    }
}

pub fn vb6_string_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    "\"".parse_next(input)?;

    let start_index = input.index;

    loop {
        take_till(1.., ['"']).parse_next(input)?;

        if literal::<_, _, VB6ErrorKind>("\"\"")
            .parse_next(input)
            .is_err()
        {
            let end_index = input.index;
            "\"".parse_next(input)?;

            return Ok(input.stream[start_index..end_index].as_bstr());
        }
    }
}

#[cfg(test)]
mod tests {
    use winnow::stream::StreamIsPartial;

    use super::*;
    use crate::parsers::VB6Stream;

    use super::HeaderKind;

    #[test]
    fn vb6_no_double_quote_string_parse() {
        let input_line = b"\"This is a string\"\r\n";
        let mut stream = VB6Stream::new("", input_line);
        let string = vb6_string_parse(&mut stream).unwrap();

        assert_eq!(string, "This is a string");
    }

    #[test]
    fn vb6_with_double_quote_string_parse() {
        let input_line = b"\"This is also \"\"a\"\" string\"\r\n";
        let mut stream = VB6Stream::new("", input_line);
        let string = vb6_string_parse(&mut stream).unwrap();

        assert_eq!(string, "This is also \"\"a\"\" string");
    }

    #[test]
    fn compiled_object_line_valid() {
        let mut input = VB6Stream::new(
            "",
            b"Object={C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0; crviewer.dll\r\n",
        );

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Object=".parse_next(&mut input);

        let result = object_parse.parse_next(&mut input).unwrap();

        let expected_uuid = Uuid::parse_str("C4847593-972C-11D0-9567-00A0C9273C2A").unwrap();

        // we don't consume the line ending, so we should have 2 bytes left.
        assert_eq!(input.complete(), 2);

        match result {
            VB6ObjectReference::Compiled {
                uuid,
                version,
                unknown1,
                file_name,
            } => {
                assert_eq!(uuid, expected_uuid);
                assert_eq!(version, "8.0");
                assert_eq!(unknown1, "0");
                assert_eq!(file_name, "crviewer.dll");
            }
            _ => panic!("Expected a compiled object reference."),
        }
    }

    #[test]
    fn project_object_line_valid() {
        let mut input = VB6Stream::new("", b"Object=*\\A..\\vbGraph.vbp\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Object=".parse_next(&mut input);

        let result = object_parse.parse_next(&mut input).unwrap();

        // we don't consume the line ending, so we should have 2 bytes left.
        assert_eq!(input.complete(), 2);

        match result {
            VB6ObjectReference::Project { path } => {
                assert_eq!(path, "..\\vbGraph.vbp");
            }
            _ => panic!("Expected a project object reference."),
        }
    }

    #[test]
    fn test_key_resource_offset_line_parse() {
        let input_line = b"      Picture         =   \"Brightness.frx\":0000\r\n";
        let mut stream = VB6Stream::new("", input_line);
        let (key, resource_file_name, offset) =
            key_resource_offset_line_parse("=")(&mut stream).unwrap();

        assert_eq!(key, "Picture".as_bytes());
        assert_eq!(resource_file_name, "Brightness.frx".as_bytes());
        assert_eq!(offset, "0000".as_bytes());
    }

    #[test]
    fn test_key_resource_offset_line_with_comment_parse() {
        let input_line = b"      Picture         =   \"Brightness.frx\":0000\r\n";
        let mut stream = VB6Stream::new("", input_line);
        let (key, resource_file_name, offset) =
            key_resource_offset_line_parse("=")(&mut stream).unwrap();

        assert_eq!(key, "Picture".as_bytes());
        assert_eq!(resource_file_name, "Brightness.frx".as_bytes());
        assert_eq!(offset, "0000".as_bytes());
    }

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
