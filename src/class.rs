#![warn(clippy::pedantic)]

use bstr::BStr;
use miette::Result;

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    header::{key_value_line_parse, version_parse, HeaderKind},
    vb6::{keyword_parse, line_comment_parse, vb6_parse, VB6Token},
    vb6stream::VB6Stream,
    VB6FileFormatVersion,
};

use winnow::{
    ascii::{line_ending, space0},
    combinator::{alt, preceded, repeat_till},
    error::ErrMode,
    PResult, Parser,
};

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

/// The properties of a VB6 class file is the list of key/value pairs
/// found between the BEGIN and END lines in the header.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassProperties {
    pub multi_use: FileUsage,            // (0/-1) multi use / single use
    pub persistable: Persistance,        // (0/-1) NonParsistable / Persistable
    pub data_binding_behavior: bool,     // (0/-1) false/true - vbNone
    pub data_source_behavior: bool,      // (0/-1) false/true - vbNone
    pub mts_transaction_mode: MtsStatus, // (0/-1) NotAnMTSObject / MTSObject
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
    pub properties: VB6ClassProperties,
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
    /// let result = VB6ClassFile::parse("class_parse.cls".to_owned(), &mut input.as_slice());
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(file_name: String, input: &mut &'a [u8]) -> Result<Self, VB6Error> {
        let input = &mut VB6Stream::new(file_name, input);

        let Ok(header) = class_header_parse(input) else {
            return Err(input.error(VB6ErrorKind::Header));
        };

        let Ok(tokens) = vb6_parse(input) else {
            return Err(input.error(VB6ErrorKind::FileContent));
        };

        Ok(VB6ClassFile { header, tokens })
    }
}

/// Parses a VB6 class file from the header.
/// The header contains the version, multi use, persistable, data binding behavior,
/// data source behavior, and MTS transaction mode.
/// The header also contains the attributes of the class file.
/// The header is not normally visible in the code editor region.
/// It is only visible in the file property explorer.
///
/// # Arguments
///
/// * `input` The stream to parse.
///
/// # Returns
///
/// A result containing the parsed VB6 class file header or an error.
///
/// # Errors
///
/// An error will be returned if the input is not a valid VB6 class file header.
///
fn class_header_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6ClassHeader<'a>, VB6Error> {
    // VERSION #.# CLASS
    // BEGIN
    //  key = value  'comment
    //  ...
    // END

    let version = version_parse(HeaderKind::Class).parse_next(input)?;

    let properties = properties_parse.parse_next(input)?;

    let attributes = attributes_parse.parse_next(input)?;

    Ok(VB6ClassHeader {
        version,
        properties,
        attributes,
    })
}

fn begin_line_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<(), VB6Error> {
    (space0, keyword_parse("BEGIN"), space0).parse_next(input)?;

    alt(((line_comment_parse, line_ending), (space0, line_ending))).parse_next(input)?;

    Ok(())
}

fn end_line_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<(), VB6Error> {
    (space0, keyword_parse("END"), space0).parse_next(input)?;

    alt(((line_comment_parse, line_ending), (space0, line_ending))).parse_next(input)?;

    Ok(())
}

/// Parses a VB6 class file properties from the header.
/// The properties are the key/value pairs found between the BEGIN and END lines in the header.
/// The properties contain the multi use, persistable, data binding behavior,
/// data source behavior, and MTS transaction mode.
/// The properties are not normally visible in the code editor region.
/// They are only visible in the file property explorer.
///
/// # Arguments
///
/// * `input` The stream to parse.
///
/// # Returns
///
/// A result containing the parsed VB6 class file properties or an error.
fn properties_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6ClassProperties, VB6Error> {
    begin_line_parse.parse_next(input)?;

    let mut multi_use = FileUsage::MultiUse;
    let mut persistable = Persistance::NonPersistable;
    let mut data_binding_behavior = false;
    let mut data_source_behavior = false;
    let mut mts_transaction_mode = MtsStatus::NotAnMTSObject;

    let (collection, _): (Vec<(&BStr, &BStr)>, _) =
        repeat_till(0.., key_value_line_parse("="), end_line_parse).parse_next(input)?;

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

    Ok(VB6ClassProperties {
        multi_use,
        persistable,
        data_binding_behavior,
        data_source_behavior,
        mts_transaction_mode,
    })
}

fn attributes_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<VB6FileAttributes<'a>, VB6Error> {
    let _ = space0::<_, VB6Error>.parse_next(input);

    let mut name = Option::None;
    let mut global_name_space = false;
    let mut creatable = false;
    let mut pre_declared_id = false;
    let mut exposed = false;

    while let Ok((key, value)) =
        preceded(keyword_parse("Attribute"), key_value_line_parse("=")).parse_next(input)
    {
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
        return Err(ErrMode::Cut(input.error(VB6ErrorKind::MissingClassName)));
    }

    Ok(VB6FileAttributes {
        name: name.unwrap(),
        global_name_space,
        creatable,
        pre_declared_id,
        exposed,
    })
}

#[cfg(test)]
mod tests {
    use super::HeaderKind;
    use super::*;

    #[test]
    fn class_file_valid() {
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

Option Explicit
";

        let result = VB6ClassFile::parse("test.vb".to_owned(), &mut input.as_slice());

        assert!(result.is_ok());
    }

    #[test]
    fn class_file_invalid() {
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
                    \r
                    Option Explicit\r";

        let result = VB6ClassFile::parse("test.vb".to_owned(), &mut input.as_slice());

        assert!(result.is_err());
    }

    #[test]
    fn class_header_valid() {
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
Attribute VB_Exposed = False";

        let mut stream = VB6Stream::new("", &mut input.as_slice());
        let result = class_header_parse(&mut stream);

        assert!(result.is_ok());
    }

    #[test]
    fn class_header_invalid() {
        let input = b"MultiUse = -1  'True\r
    Persistable = 0  'NotPersistable\r
    DataBindingBehavior = 0  'vbNone\r
    DataSourceBehavior = 0  'vbNone\r
    MTSTransactionMode = 0  'NotAnMTSObject\r
    ";

        let mut stream = VB6Stream::new("", &mut input.as_slice());
        let result = class_header_parse(&mut stream);

        assert!(result.is_err());
    }

    #[test]
    fn attributes_valid() {
        let input = b"Attribute VB_Name = \"Something\"\r
    Attribute VB_GlobalNameSpace = False\r
    Attribute VB_Creatable = True\r
    Attribute VB_PredeclaredId = False\r
    Attribute VB_Exposed = False\r
    ";

        let mut stream = VB6Stream::new("", &mut input.as_slice());
        let result = attributes_parse(&mut stream);

        assert!(result.is_ok());
    }

    #[test]
    fn attributes_invalid() {
        let input = b"Attribut VB_Name = \"Something\"\r
    Attrbute VB_GlobalNameSpace = False\r
    Attribut VB_Creatable = True\r
    Attriute VB_PredeclaredId = False\r
    Atribute VB_Exposed = False\r
    ";

        let mut stream = VB6Stream::new("", &mut input.as_slice());
        let result = attributes_parse(&mut stream);

        assert!(result.is_err());
    }

    #[test]
    fn version_valid() {
        let input = b"VERSION 1.0 CLASS\r\n";

        let mut stream = VB6Stream::new("", &mut input.as_slice());
        let result = version_parse(HeaderKind::Class).parse_next(&mut stream);

        println!("{:?}", result);

        assert!(result.is_ok());
    }

    #[test]
    fn version_invalid() {
        // Missing the return character and newline character at the end.
        let input = b"VERSION 1.0 CLASS";

        let mut stream = VB6Stream::new("", &mut input.as_slice());
        let result = version_parse(HeaderKind::Class).parse_next(&mut stream);

        assert!(result.is_err());
    }
}
