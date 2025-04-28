use std::collections::HashMap;

use bstr::{BStr, ByteSlice};
use serde::Serialize;
use uuid::Uuid;
use winnow::{
    ascii::{digit1, line_ending, space0, space1},
    combinator::{alt, eof, opt},
    error::ErrMode,
    stream::Stream,
    token::{literal, take_till, take_until, take_while},
    Parser,
};

use crate::{
    errors::{PropertyError, VB6ErrorKind},
    parsers::{objectreference::VB6ObjectReference, VB6Stream},
    vb6::{keyword_parse, line_comment_parse, string_parse, take_until_line_ending, VB6Result},
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

/// Represents if a class is in the global or local name space.
///
/// The global name space is the default name space for a class.
/// In the file, `VB_GlobalNameSpace` of 'False' means the class is in the local name space.
/// `VB_GlobalNameSpace` of 'True' means the class is in the global name space.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum NameSpace {
    Global,
    Local,
}

/// The creatable attribute is used to determine if the class can be created.
///
/// If True, the class can be created from anywhere. The class is essentially public.
/// If False, the class can only be created from within the class itself.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum Creatable {
    False,
    True,
}

/// Used to determine if the class has a pre-declared ID.
///
/// If True, the class has a pre-declared ID and can be accessed by
/// the class name without creating an instance of the class.
///
/// If False, the class does not have a pre-declared ID and must be
/// accessed by creating an instance of the class.
///
/// If True and the `VB_GlobalNameSpace` is True, the class shares namespace
/// access semantics with the VB6 intrinsic classes.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum PreDeclaredID {
    False,
    True,
}

/// Used to determine if the class is exposed.
///
/// The `VB_Exposed` attribute is not normally visible in the code editor region.
///
/// ----------------------------------------------------------------------------
///
/// True is public and False is internal.
/// Used in combination with the Creatable attribute to create a matrix of
/// scoping behavior.
///
/// ----------------------------------------------------------------------------
///
/// Private (Default).
///
/// `VB_Exposed` = False and `VB_Creatable` = False.
/// The class is accessible only within the enclosing project.
///
/// Instances of the class can only be created by modules contained within the
/// project that defines the class.
///
/// ----------------------------------------------------------------------------
///
/// Public Not Creatable.
///
/// `VB_Exposed` = True and `VB_Creatable` = False.
/// The class is accessible within the enclosing project and within projects
/// that reference the enclosing project.
///
/// Instances of the class can only be created by modules within the enclosing
/// project. Modules in other projects can reference the class name as a
/// declared type but can’t instantiate the class using new or the
/// `CreateObject` function.
///
/// ----------------------------------------------------------------------------
///
/// Public Creatable.
///
/// `VB_Exposed` = True and `VB_Creatable` = True.
/// The class is accessible within the enclosing project and within the
/// enclosing project and within projects that reference the enclosing project.
///
/// Any module that can access the class can create instances of it.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum Exposed {
    False,
    True,
}

/// Represents the attributes of a VB6 file file.
/// The attributes contain the name, global name space, creatable, pre-declared id, and exposed.
///
/// None of these values are normally visible in the code editor region.
/// They are only visible in the file property explorer.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6FileAttributes<'a> {
    pub name: &'a BStr,                       // Attribute VB_Name = "Organism"
    pub global_name_space: NameSpace,         // (True/False) Attribute VB_GlobalNameSpace = False
    pub creatable: Creatable,                 // (True/False) Attribute VB_Creatable = True
    pub pre_declared_id: PreDeclaredID,       // (True/False) Attribute VB_PredeclaredId = False
    pub exposed: Exposed,                     // (True/False) Attribute VB_Exposed = False
    pub description: Option<&'a BStr>,        // Attribute VB_Description = "Description"
    pub ext_key: HashMap<&'a BStr, &'a BStr>, // Additional attributes
}

impl Default for VB6FileAttributes<'_> {
    fn default() -> Self {
        VB6FileAttributes {
            name: BStr::new(""),
            global_name_space: NameSpace::Local,
            creatable: Creatable::True,
            pre_declared_id: PreDeclaredID::False,
            exposed: Exposed::False,
            description: None,
            ext_key: HashMap::new(),
        }
    }
}
enum Attributes {
    Name,
    GlobalNameSpace,
    Creatable,
    PredeclaredId,
    Exposed,
    Description,
    ExtKey,
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
    if (space0::<_, VB6ErrorKind>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    let object = match alt((compiled_object_parse, project_object_parse)).parse_next(input) {
        Ok(object) => object,
        Err(e) => return Err(ErrMode::Cut(e.into_inner().unwrap())),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    Ok(object)
}

pub fn attributes_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6FileAttributes<'a>> {
    let _ = space0::<_, VB6ErrorKind>.parse_next(input);

    let mut name = None;
    let mut global_name_space = NameSpace::Local;
    let mut creatable = Creatable::True;
    let mut pre_declared_id = PreDeclaredID::False;
    let mut exposed = Exposed::False;
    let mut description = None;
    let mut ext_key = HashMap::new();

    while (space0, keyword_parse("Attribute"), space0)
        .parse_next(input)
        .is_ok()
    {
        space0.parse_next(input)?;

        let Ok(key) = alt((
            keyword_parse("VB_Name").map(|_| Attributes::Name),
            keyword_parse("VB_GlobalNameSpace").map(|_| Attributes::GlobalNameSpace),
            keyword_parse("VB_Creatable").map(|_| Attributes::Creatable),
            keyword_parse("VB_PredeclaredId").map(|_| Attributes::PredeclaredId),
            keyword_parse("VB_Exposed").map(|_| Attributes::Exposed),
            keyword_parse("VB_Description").map(|_| Attributes::Description),
            keyword_parse("VB_Ext_KEY").map(|_| Attributes::ExtKey),
        ))
        .parse_next(input) else {
            return Err(ErrMode::Cut(VB6ErrorKind::UnknownAttribute));
        };

        let _ = (space0, "=", space0).parse_next(input)?;

        match key {
            Attributes::Name => {
                name = match string_parse.parse_next(input) {
                    Ok(name) => Some(name),
                    Err(_) => return Err(ErrMode::Cut(VB6ErrorKind::StringParseError)),
                };

                space0.parse_next(input)?;
                alt((line_comment_parse, line_ending, eof)).parse_next(input)?;
            }
            Attributes::Description => {
                description = match string_parse.parse_next(input) {
                    Ok(description) => Some(description),
                    Err(_) => return Err(ErrMode::Cut(VB6ErrorKind::StringParseError)),
                };

                space0.parse_next(input)?;
                alt((line_comment_parse, line_ending, eof)).parse_next(input)?;
            }
            Attributes::GlobalNameSpace => {
                global_name_space = match alt((
                    literal::<_, _, VB6ErrorKind>("True").map(|_| NameSpace::Global),
                    literal::<_, _, VB6ErrorKind>("False").map(|_| NameSpace::Local),
                ))
                .parse_next(input)
                {
                    Ok(global_name_space) => global_name_space,
                    Err(_) => {
                        return Err(ErrMode::Cut(VB6ErrorKind::Property(
                            PropertyError::InvalidPropertyValueTrueFalse,
                        )))
                    }
                };

                space0.parse_next(input)?;
                alt((line_comment_parse, line_ending, eof)).parse_next(input)?;
            }
            Attributes::Creatable => {
                creatable = match alt((
                    literal::<_, _, VB6ErrorKind>("True").map(|_| Creatable::True),
                    literal::<_, _, VB6ErrorKind>("False").map(|_| Creatable::False),
                ))
                .parse_next(input)
                {
                    Ok(creatable) => creatable,
                    Err(_) => {
                        return Err(ErrMode::Cut(VB6ErrorKind::Property(
                            PropertyError::InvalidPropertyValueTrueFalse,
                        )))
                    }
                };

                space0.parse_next(input)?;
                alt((line_comment_parse, line_ending, eof)).parse_next(input)?;
            }
            Attributes::PredeclaredId => {
                pre_declared_id = match alt((
                    literal::<_, _, VB6ErrorKind>("True").map(|_| PreDeclaredID::True),
                    literal::<_, _, VB6ErrorKind>("False").map(|_| PreDeclaredID::False),
                ))
                .parse_next(input)
                {
                    Ok(pre_declared_id) => pre_declared_id,
                    Err(_) => {
                        return Err(ErrMode::Cut(VB6ErrorKind::Property(
                            PropertyError::InvalidPropertyValueTrueFalse,
                        )))
                    }
                };

                space0.parse_next(input)?;
                alt((line_comment_parse, line_ending, eof)).parse_next(input)?;
            }
            Attributes::Exposed => {
                exposed = match alt((
                    literal::<_, _, VB6ErrorKind>("True").map(|_| Exposed::True),
                    literal::<_, _, VB6ErrorKind>("False").map(|_| Exposed::False),
                ))
                .parse_next(input)
                {
                    Ok(exposed) => exposed,
                    Err(_) => {
                        return Err(ErrMode::Cut(VB6ErrorKind::Property(
                            PropertyError::InvalidPropertyValueTrueFalse,
                        )))
                    }
                };

                space0.parse_next(input)?;
                alt((line_comment_parse, line_ending, eof)).parse_next(input)?;
            }
            Attributes::ExtKey => {
                let Ok(key) = string_parse.parse_next(input) else {
                    return Err(ErrMode::Cut(VB6ErrorKind::StringParseError));
                };

                (space0, ",", space0).parse_next(input)?;

                let Ok(value) = string_parse.parse_next(input) else {
                    return Err(ErrMode::Cut(VB6ErrorKind::StringParseError));
                };

                ext_key.insert(key, value);

                space0.parse_next(input)?;
                alt((line_comment_parse, line_ending, eof)).parse_next(input)?;
            }
        }
    }

    if name.is_none() {
        return Err(ErrMode::Cut(VB6ErrorKind::MissingNameAttribute));
    }

    Ok(VB6FileAttributes {
        name: name.unwrap(),
        global_name_space,
        creatable,
        pre_declared_id,
        exposed,
        description,
        ext_key,
    })
}

pub fn key_value_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<(&'a BStr, &'a BStr)> {
        let checkpoint = input.checkpoint();

        space0.parse_next(input)?;

        let Ok(key) = take_until::<_, _, VB6ErrorKind>(1.., (" ", "\t", divider)).parse_next(input)
        else {
            input.reset(&checkpoint);
            return Err(ErrMode::Cut(VB6ErrorKind::Property(
                PropertyError::NameUnparsable,
            )));
        };

        space0.parse_next(input)?;

        if literal::<_, _, VB6ErrorKind>(divider)
            .parse_next(input)
            .is_err()
        {
            input.reset(&checkpoint);
            return Err(ErrMode::Cut(VB6ErrorKind::NoKeyValueDividerFound));
        }

        space0.parse_next(input)?;

        let Ok(value) =
            alt((string_parse, take_until(1.., (" ", "\t", "\r", "\n")))).parse_next(input)
        else {
            input.reset(&checkpoint);
            return Err(ErrMode::Cut(VB6ErrorKind::KeyValueParseError));
        };

        space0.parse_next(input)?;

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
        if (space0, opt(line_comment_parse), alt((line_ending, eof)))
            .parse_next(input)
            .is_err()
        {
            input.reset(&checkpoint);
            return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
        }

        Ok((key, value))
    }
}

pub fn key_resource_offset_line_parse<'a>(
    input: &mut VB6Stream<'a>,
) -> VB6Result<(&'a BStr, &'a BStr, u32)> {
    let checkpoint = input.checkpoint();

    space0.parse_next(input)?;

    let Ok(key) = take_till::<_, _, VB6ErrorKind>(1.., (' ', '\t', '=')).parse_next(input) else {
        input.reset(&checkpoint);
        return Err(ErrMode::Cut(VB6ErrorKind::Property(
            PropertyError::NameUnparsable,
        )));
    };

    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        input.reset(&checkpoint);
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    let Ok(resource_file_name) = string_parse.parse_next(input) else {
        input.reset(&checkpoint);
        return Err(ErrMode::Cut(VB6ErrorKind::Property(
            PropertyError::ResourceFileNameUnparsable,
        )));
    };

    if literal::<_, _, VB6ErrorKind>(":")
        .parse_next(input)
        .is_err()
    {
        input.reset(&checkpoint);
        return Err(ErrMode::Cut(VB6ErrorKind::NoColonForOffsetSplit));
    }

    let offset_txt = match take_while(1.., ('0'..='9', 'A'..='F')).parse_next(input) {
        Ok(offset_txt) => offset_txt,
        Err(e) => {
            input.reset(&checkpoint);
            return Err(e);
        }
    };

    let offset = match u32::from_str_radix(offset_txt.to_str().unwrap(), 16) {
        Ok(offset) => offset,
        Err(_) => {
            return Err(ErrMode::Cut(VB6ErrorKind::Property(
                PropertyError::OffsetUnparsable,
            )));
        }
    };

    (space0, opt(line_comment_parse)).parse_next(input)?;

    // we have to check for eof here because it's perfectly possible to have a
    // header file that is empty of actual code. This means the last line of the file
    // should be an empty line, but it might be that the file ends at the end of the
    // header attribute section.
    if alt((line_ending::<_, VB6ErrorKind>, eof))
        .parse_next(input)
        .is_err()
    {
        input.reset(&checkpoint);
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    Ok((key, resource_file_name, offset))
}

#[cfg(test)]
mod tests {
    use winnow::stream::StreamIsPartial;

    use super::*;
    use crate::parsers::VB6Stream;

    use super::HeaderKind;

    #[test]
    fn compiled_object_line_valid() {
        let mut input = VB6Stream::new(
            "",
            b"Object={C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0; crviewer.dll\r\n",
        );

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Object".parse_next(&mut input);

        let result = object_parse.parse_next(&mut input).unwrap();

        let expected_uuid = Uuid::parse_str("C4847593-972C-11D0-9567-00A0C9273C2A").unwrap();

        assert_eq!(input.complete(), 0);

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

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Object".parse_next(&mut input);

        let result = object_parse.parse_next(&mut input).unwrap();

        assert_eq!(input.complete(), 0);

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
        let (key, resource_file_name, offset) = key_resource_offset_line_parse
            .parse_next(&mut stream)
            .unwrap();

        assert_eq!(key, "Picture".as_bytes());
        assert_eq!(resource_file_name, "Brightness.frx".as_bytes());
        assert_eq!(offset, 0u32);
    }

    #[test]
    fn test_key_resource_offset_line_with_comment_parse() {
        let input_line = b"      Picture         =   \"Brightness.frx\":0000 'comment\r\n";
        let mut stream = VB6Stream::new("", input_line);
        let (key, resource_file_name, offset) = key_resource_offset_line_parse
            .parse_next(&mut stream)
            .unwrap();

        assert_eq!(key, "Picture".as_bytes());
        assert_eq!(resource_file_name, "Brightness.frx".as_bytes());
        assert_eq!(offset, 0u32);
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
