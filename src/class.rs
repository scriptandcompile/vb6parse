#![warn(clippy::pedantic)]

use bstr::BStr;

use miette::{Diagnostic, NamedSource, Result, SourceOffset, SourceSpan};
use thiserror::Error;

use crate::vb6::{eol_comment_parse, keyword_parse, vb6_parse, VB6Token};
use crate::vb6stream::VB6Stream;
use crate::VB6FileFormatVersion;

use winnow::{
    ascii::{digit1, line_ending, space0, space1},
    combinator::{opt, repeat_till},
    error::{ContextError, ErrMode, ParserError, StrContext},
    token::{literal, take_while},
    PResult, Parser,
};

#[derive(Debug, Error, Diagnostic)]
#[error("A parsing error occured")]
pub struct ErrorInfo {
    #[source_code]
    pub src: NamedSource<String>,
    #[label("oh no")]
    pub location: SourceSpan,
}

impl ErrorInfo {
    pub fn new(
        source: impl Into<String>,
        file_name: impl Into<String>,
        line: usize,
        column: usize,
        len: usize,
    ) -> Self {
        let code = source.into().clone();
        Self {
            src: NamedSource::new(file_name.into(), code.clone()),
            location: SourceSpan::new(SourceOffset::from_location(&code, line, column), len),
        }
    }
}

#[derive(Error, Debug, Diagnostic)]
pub enum ClassParseError {
    #[error("Error parsing header")]
    #[diagnostic(transparent)]
    Header {
        #[label = "A parsing error occured"]
        info: ErrorInfo,
    },
    #[error("Error parsing the VB6 file contents")]
    #[diagnostic(transparent)]
    FileContent {
        #[label = "A parsing error occured"]
        info: ErrorInfo,
    },
}

/// Represents the usage of a file.
/// -1 is 'true' and 0 is 'false' in VB6.
/// `MultiUse` is -1 and `SingleUse` is 0.
/// `MultiUse` is true and `SingleUse` is false.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum FileUsage {
    MultiUse,  // -1 (true)
    SingleUse, // 0 (false)
}

/// Represents the persistability of a file.
/// -1 is 'true' and 0 is 'false' in VB6.
/// `Persistable` is -1 and `NonPersistable` is 0.
/// `Persistable` is true and `NonPersistable` is false.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Persistance {
    Persistable,    // -1 (true)
    NonPersistable, // 0 (false)
}

/// Represents the MTS status of a file.
/// -1 is 'true' and 0 is 'false' in VB6.
/// `MTSObject` is -1 and `NotAnMTSObject` is 0.
/// `MTSObject` is true and `NotAnMTSObject` is false.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum MtsStatus {
    NotAnMTSObject, // 0 (false)
    MTSObject,      // -1 (true)
}

/// Represents the header of a VB6 class file.
/// The header contains the version, multi use, persistable, data binding behavior,
/// data source behavior, and MTS transaction mode.
/// The header also contains the attributes of the class file.
///
/// None of these values are normally visible in the code editor region.
/// They are only visible in the file property explorer.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassHeader<'a> {
    pub version: VB6FileFormatVersion,
    pub multi_use: FileUsage,            // (0/-1) multi use / single use
    pub persistable: Persistance,        // (0/-1) NonParsistable / Persistable
    pub data_binding_behavior: bool,     // (0/-1) false/true - vbNone
    pub data_source_behavior: bool,      // (0/-1) false/true - vbNone
    pub mts_transaction_mode: MtsStatus, // (0/-1) NotAnMTSObject / MTSObject
    pub attributes: VB6FileAttributes<'a>,
}

/// Represents the attributes of a VB6 class file.
/// The attributes contain the name, global name space, creatable, pre-declared id, and exposed.
///
/// None of these values are normally visible in the code editor region.
/// They are only visible in the file property explorer.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassAttributes<'a> {
    pub name: &'a [u8],          // Attribute VB_Name = "Organism"
    pub global_name_space: bool, // (True/False) Attribute VB_GlobalNameSpace = False
    pub creatable: bool,         // (True/False) Attribute VB_Creatable = True
    pub pre_declared_id: bool,   // (True/False) Attribute VB_PredeclaredId = False
    pub exposed: bool,           // (True/False) Attribute VB_Exposed = False
}

/// Represents a VB6 class file.
/// A VB6 class file contains a header and a list of tokens.
///
/// The header contains the version, multi use, persistable, data binding behavior,
/// data source behavior, and MTS transaction mode.
/// The header also contains the attributes of the class file.
///
/// The tokens contain the token stream of the code of the class file.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassFile<'a> {
    pub header: VB6ClassHeader<'a>,
    pub tokens: Vec<VB6Token<'a>>,
}

/// Represents the version of a VB6 class file.
/// The version contains a major and minor version number.
///
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassVersion {
    pub major: u8,
    pub minor: u8,
}

/// Represents the attributes of a VB6 class file.
/// The attributes contain the name, global name space, creatable, pre-declared id, and exposed.
///
/// None of these values are normally visible in the code editor region.
/// They are only visible in the file property explorer.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6FileAttributes<'a> {
    pub name: &'a [u8],          // Attribute VB_Name = "Organism"
    pub global_name_space: bool, // (True/False) Attribute VB_GlobalNameSpace = False
    pub creatable: bool,         // (True/False) Attribute VB_Creatable = True
    pub pre_declared_id: bool,   // (True/False) Attribute VB_PredeclaredId = False
    pub exposed: bool,           // (True/False) Attribute VB_Exposed = False
}

impl<'a> VB6ClassFile<'a> {
    /// Parses a VB6 class file from a byte slice.
    ///
    /// # Arguments
    ///
    /// * `input` The byte slice to parse.
    ///
    /// # Returns
    ///
    /// A result containing the parsed VB6 class file or an error.
    ///
    /// # Errors
    ///
    /// An error will be returned if the input is not a valid VB6 class file.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::class::VB6ClassFile;
    ///
    /// let input = b"VERSION 1.0 CLASS
    /// BEGIN
    ///   MultiUse = -1  'True
    ///   Persistable = 0  'NotPersistable
    ///   DataBindingBehavior = 0  'vbNone
    ///   DataSourceBehavior = 0  'vbNone
    ///   MTSTransactionMode = 0  'NotAnMTSObject
    /// END
    /// Attribute VB_Name = \"Something\"
    /// Attribute VB_GlobalNameSpace = False
    /// Attribute VB_Creatable = True
    /// Attribute VB_PredeclaredId = False
    /// Attribute VB_Exposed = False
    /// ";
    ///
    /// let result = VB6ClassFile::parse(&mut input.as_slice());
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(input: &mut &'a [u8]) -> Result<Self, ClassParseError> {
        let input = &mut VB6Stream::new(input);

        let Ok(header) = class_header_parse(input) else {
            let err_info = ErrorInfo::new(input.stream.to_string(), "", 0, 0, 20);
            return Err(ClassParseError::Header { info: err_info });
        };

        let Ok(tokens) = vb6_parse(input) else {
            let err_info = ErrorInfo::new(input.stream.to_string(), "", 0, 0, 20);
            return Err(ClassParseError::FileContent { info: err_info });
        };

        Ok(VB6ClassFile { header, tokens })
    }
}

/// Parses a VB6 class file header from a byte slice.
///
/// # Arguments
///
/// * `input` The byte slice to parse.
///
/// # Returns
///
/// A result containing the parsed VB6 class file header or an error.
///
/// # Errors
///
/// An error will be returned if the input is not a valid VB6 class file header.
///
fn class_header_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6ClassHeader<'a>> {
    // VERSION #.# CLASS
    // BEGIN
    //  key = value  'comment
    //  ...
    // END

    let version = version_parse.parse_next(input)?;

    space0.parse_next(input)?;

    keyword_parse("BEGIN").parse_next(input)?;

    space0.parse_next(input)?;

    line_ending
        .context(StrContext::Label("Newline expected after BEGIN keyword."))
        .parse_next(input)?;

    let mut multi_use = FileUsage::MultiUse;
    let mut persistable = Persistance::NonPersistable;
    let mut data_binding_behavior = false;
    let mut data_source_behavior = false;
    let mut mts_transaction_mode = MtsStatus::NotAnMTSObject;

    let (collection, _): (Vec<(&BStr, &BStr)>, &BStr) =
        repeat_till(0.., key_value_line_parse("="), keyword_parse("END")).parse_next(input)?;

    for pair in collection.iter() {
        let (key, value) = pair;

        match key.to_ascii_lowercase().as_slice() {
            b"persistable" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value.to_ascii_lowercase().as_slice() == b"-1" {
                    persistable = Persistance::Persistable;
                } else {
                    persistable = Persistance::NonPersistable;
                }
            }
            b"multiuse" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value.to_ascii_lowercase().as_slice() == b"-1" {
                    multi_use = FileUsage::MultiUse;
                } else {
                    multi_use = FileUsage::SingleUse;
                }
            }
            b"databindingbehavior" => {
                // -1 is 'true' and 0 is 'false' in VB6
                data_binding_behavior = value.to_ascii_lowercase().as_slice() == b"-1";
            }
            b"datasourcebehavior" => {
                // -1 is 'true' and 0 is 'false' in VB6
                data_source_behavior = value.to_ascii_lowercase().as_slice() == b"-1";
            }
            b"mtstransactionmode" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value.to_ascii_lowercase().as_slice() == b"-1" {
                    mts_transaction_mode = MtsStatus::MTSObject;
                } else {
                    mts_transaction_mode = MtsStatus::NotAnMTSObject;
                }
            }
            _ => {
                panic!("Unknown key found in class header.");
            }
        }
    }

    line_ending
        .context(StrContext::Label("Newline expected after END."))
        .parse_next(input)?;

    let attributes = attributes_parse.parse_next(input)?;

    Ok(VB6ClassHeader {
        version,
        multi_use,
        persistable,
        data_binding_behavior,
        data_source_behavior,
        mts_transaction_mode,
        attributes,
    })
}

fn attributes_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6FileAttributes<'a>> {
    let _ = space0::<VB6Stream, ContextError>.parse_next(input);

    let mut name = Option::None;
    let mut global_name_space = false;
    let mut creatable = false;
    let mut pre_declared_id = false;
    let mut exposed = false;

    while let Ok((_, (key, value))) =
        (keyword_parse("Attribute"), key_value_parse("=")).parse_next(input)
    {
        line_ending
            .context(StrContext::Label(
                "Newline expected after Class File Attribute line.",
            ))
            .parse_next(input)?;

        match key.to_ascii_lowercase().as_slice() {
            b"vb_name" => {
                name = Some(value);
            }
            b"vb_globalnamespace" => {
                global_name_space = value.to_ascii_lowercase().as_slice() == b"true";
            }
            b"vb_creatable" => {
                creatable = value.to_ascii_lowercase().as_slice() == b"true";
            }
            b"vb_predeclaredid" => {
                pre_declared_id = value.to_ascii_lowercase().as_slice() == b"true";
            }
            b"vb_exposed" => {
                exposed = value.to_ascii_lowercase().as_slice() == b"true";
            }
            _ => {
                panic!("Unknown key found in class attributes.");
            }
        }
    }

    if name.is_none() {
        let error = ParserError::assert(input, "VB_Name attribute not found.");

        return Err(ErrMode::Cut(error));
    }

    Ok(VB6FileAttributes {
        name: name.unwrap(),
        global_name_space,
        creatable,
        pre_declared_id,
        exposed,
    })
}

fn key_value_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> PResult<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> Result<(&'a BStr, &'a BStr), ErrMode<ContextError>> {
        let _ = space0::<VB6Stream<'a>, ContextError>.parse_next(input);

        let key = take_while(1.., ('_', '"', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9'))
            .parse_next(input)?;

        let _ = space0::<VB6Stream<'a>, ContextError>.parse_next(input);

        literal(divider).parse_next(input)?;

        let _ = space0::<VB6Stream<'a>, ContextError>.parse_next(input);

        opt("\"").parse_next(input)?;

        let value =
            take_while(1.., ('_', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9')).parse_next(input)?;

        opt("\"").parse_next(input)?;

        let _ = space0::<VB6Stream<'a>, ContextError>.parse_next(input);

        Ok((key, value))
    }
}

fn key_value_line_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> PResult<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> PResult<(&'a BStr, &'a BStr)> {
        let (key, value) = key_value_parse(divider).parse_next(input)?;

        eol_comment_parse.parse_next(input)?;

        line_ending
            .context(StrContext::Label("newline not found"))
            .parse_next(input)?;

        Ok((key, value))
    }
}

fn version_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6FileFormatVersion> {
    keyword_parse("VERSION")
        .context(StrContext::Label("'VERSION' header element not found."))
        .parse_next(input)?;

    space1.context(StrContext::Label("At least one space is required between the 'VERSION' header element and the header major version number.")).parse_next(input)?;

    let major_digits = digit1
        .context(StrContext::Label("Major version number not found."))
        .parse_next(input)?;

    let major_version =
        u8::from_str_radix(bstr::BStr::new(major_digits).to_string().as_str(), 10).unwrap();

    ".".context(StrContext::Label("Version decimal character not found."))
        .parse_next(input)?;

    let minor_digits = digit1
        .context(StrContext::Label("Minor version number not found."))
        .parse_next(input)?;

    let minor_version =
        u8::from_str_radix(bstr::BStr::new(minor_digits).to_string().as_str(), 10).unwrap();

    space1.context(StrContext::Label("At least one space is required between the header minor version number and the 'CLASS' header element")).parse_next(input)?;

    keyword_parse("CLASS")
        .context(StrContext::Label("'Class' header element not found."))
        .parse_next(input)?;

    space0(input)?;

    line_ending
        .context(StrContext::Label("Newline expected after CLASS keyword."))
        .parse_next(input)?;

    Ok(VB6FileFormatVersion {
        major: major_version,
        minor: minor_version,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_valid_class_file() {
        let input = b"VERSION 1.0 CLASS\r
    BEGIN\r
      MultiUse = -1  'True\r
      Persistable = 0  'NotPersistable\r
      DataBindingBehavior = 0  'vbNone\r
      DataSourceBehavior = 0  'vbNone\r
      MTSTransactionMode = 0  'NotAnMTSObject\r
    END\r
    Attribute VB_Name = \"Something\"\r
    Attribute VB_GlobalNameSpace = False\r
    Attribute VB_Creatable = True\r
    Attribute VB_PredeclaredId = False\r
    Attribute VB_Exposed = False\r
    ";

        let result = VB6ClassFile::parse(&mut input.as_slice());

        assert!(result.is_ok());
    }

    #[test]
    fn parse_invalid_class_file() {
        let input = b"VERSION 1.0 CLASS
    BEGIN
      MultiUse = -1  'True
      Persistable = 0  'NotPersistable
      DataBindingBehavior = 0  'vbNone
      DataSourceBehavior = 0  'vbNone
      MTSTransactionMode = 0  'NotAnMTSObject
    END
    Attribute VB_Name = \"Something\"
    Attribute VB_GlobalNameSpace = False
    Attribute VB_Creatable = True
    Attribute VB_PredeclaredId = False
    Attribute VB_Exposed = False
    ";

        let result = VB6ClassFile::parse(&mut input.as_slice());

        assert!(result.is_err());
    }

    #[test]
    fn parse_valid_class_header() {
        let input = b"MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior = 0  'vbNone
    MTSTransactionMode = 0  'NotAnMTSObject
    ";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = class_header_parse(&mut stream);

        assert!(result.is_ok());
    }

    #[test]
    fn parse_invalid_class_header() {
        let input = b"MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior = 0  'vbNone
    MTSTransactionMode = 0  'NotAnMTSObject
    ";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = class_header_parse(&mut stream);

        assert!(result.is_err());
    }

    #[test]
    fn parse_valid_attributes() {
        let input = b"Attribute VB_Name = \"Something\"
    Attribute VB_GlobalNameSpace = False
    Attribute VB_Creatable = True
    Attribute VB_PredeclaredId = False
    Attribute VB_Exposed = False
    ";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = attributes_parse(&mut stream);

        assert!(result.is_ok());
    }

    #[test]
    fn parse_invalid_attributes() {
        let input = b"Attribute VB_Name = \"Something\"
    Attribute VB_GlobalNameSpace = False
    Attribute VB_Creatable = True
    Attribute VB_PredeclaredId = False
    Attribute VB_Exposed = False
    ";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = attributes_parse(&mut stream);

        assert!(result.is_err());
    }

    #[test]
    fn parse_valid_version() {
        let input = b"VERSION 1.0 CLASS";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = version_parse(&mut stream);

        assert!(result.is_ok());
    }

    #[test]
    fn parse_invalid_version() {
        let input = b"VERSION 1.0 CLASS";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = version_parse(&mut stream);

        assert!(result.is_err());
    }

    #[test]
    fn key_value_parse_valid() {
        let input = b"key = value";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = key_value_parse("=")(&mut stream);

        assert!(result.is_ok());
    }

    #[test]
    fn key_value_parse_invalid() {
        let input = b"key = value";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = key_value_parse(":")(&mut stream);

        assert!(result.is_err());
    }

    #[test]
    fn key_value_line_parse_valid() {
        let input = b"key = value  'comment\n";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = key_value_line_parse("=")(&mut stream);

        assert!(result.is_ok());
    }

    #[test]
    fn key_value_line_parse_invalid() {
        let input = b"key = value  'comment\n";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = key_value_line_parse(":")(&mut stream);

        assert!(result.is_err());
    }

    #[test]
    fn version_parse_valid() {
        let input = b"VERSION 1.0 CLASS";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = version_parse(&mut stream);

        assert!(result.is_ok());
    }

    #[test]
    fn version_parse_invalid() {
        let input = b"VERSION 1.0 CLASS";

        let mut stream = VB6Stream::new(&mut input.as_slice());
        let result = version_parse(&mut stream);

        assert!(result.is_err());
    }
}
