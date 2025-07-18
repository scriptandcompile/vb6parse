pub mod properties;

use bstr::{BStr, ByteSlice};
use serde::Serialize;
use winnow::{
    ascii::{line_ending, space0},
    combinator::{alt, repeat_till},
    error::ErrMode,
    Parser,
};

use crate::{
    errors::{PropertyError, VB6Error, VB6ErrorKind},
    language::VB6Token,
    parsers::{
        class::properties::*,
        header::{attributes_parse, key_value_line_parse, version_parse, HeaderKind},
        VB6Stream,
    },
    vb6::{keyword_parse, line_comment_parse, vb6_parse, VB6Result},
};

/// Represents a VB6 class file.
/// A VB6 class file contains a header and a list of tokens.
///
/// The header contains the version, multi use, persistable, data binding behavior,
/// data source behavior, and MTS transaction mode.
/// The header also contains the attributes of the class file.
///
/// The tokens contain the token stream of the code of the class file.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6ClassFile<'a> {
    pub header: VB6ClassHeader<'a>,
    pub tokens: Vec<VB6Token<'a>>,
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
    /// use vb6parse::parsers::VB6ClassFile;
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

        let header = match class_header_parse.parse_next(input) {
            Ok(header) => header,
            Err(e) => match e.into_inner() {
                Err(_) => return Err(input.error(VB6ErrorKind::Header)),
                Ok(err) => return Err(input.error(err)),
            },
        };

        let tokens = match vb6_parse.parse_next(input) {
            Ok(tokens) => tokens,
            Err(e) => match e.into_inner() {
                Err(_) => return Err(input.error(VB6ErrorKind::TokenParseError)),
                Ok(err) => return Err(input.error(err)),
            },
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
fn class_header_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ClassHeader<'a>> {
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

fn begin_line_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    let (_, begin, _) = (space0, keyword_parse("BEGIN"), space0).parse_next(input)?;

    alt(((line_comment_parse, line_ending), (space0, line_ending))).parse_next(input)?;

    Ok(begin)
}

fn end_line_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    let (_, keyword, _) = (space0, keyword_parse("END"), space0).parse_next(input)?;

    alt(((line_comment_parse, line_ending), (space0, line_ending))).parse_next(input)?;

    Ok(keyword)
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
fn properties_parse(input: &mut VB6Stream<'_>) -> VB6Result<VB6ClassProperties> {
    begin_line_parse.parse_next(input)?;

    let mut multi_use = FileUsage::MultiUse;
    let mut persistable = Persistance::NonPersistable;
    let mut data_binding_behavior = DataBindingBehavior::None;
    let mut data_source_behavior = DataSourceBehavior::None;
    let mut mts_transaction_mode = MtsStatus::NotAnMTSObject;

    let (collection, _): (Vec<(&BStr, &BStr)>, _) =
        repeat_till(0.., key_value_line_parse("="), end_line_parse).parse_next(input)?;

    for pair in &collection {
        let (key, value) = *pair;

        match key.as_bytes() {
            b"Persistable" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value == "-1" {
                    persistable = Persistance::Persistable;
                } else if value == "0" {
                    persistable = Persistance::NonPersistable;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::Property(
                        PropertyError::InvalidPropertyValueZeroNegOne,
                    )));
                }
            }
            b"MultiUse" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value == "-1" {
                    multi_use = FileUsage::MultiUse;
                } else if value == "0" {
                    multi_use = FileUsage::SingleUse;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::Property(
                        PropertyError::InvalidPropertyValueZeroNegOne,
                    )));
                }
            }
            b"DataBindingBehavior" => {
                if value == "0" {
                    data_binding_behavior = DataBindingBehavior::None;
                } else if value == "1" {
                    data_binding_behavior = DataBindingBehavior::Simple;
                } else if value == "2" {
                    data_binding_behavior = DataBindingBehavior::Complex;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::Property(
                        PropertyError::InvalidPropertyValueZeroNegOne,
                    )));
                }
            }
            b"DataSourceBehavior" => {
                if value == "0" {
                    data_source_behavior = DataSourceBehavior::None;
                } else if value == "1" {
                    data_source_behavior = DataSourceBehavior::DataSource;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::Property(
                        PropertyError::InvalidPropertyValueZeroNegOne,
                    )));
                }
            }
            b"MTSTransactionMode" => {
                if value == "0" {
                    mts_transaction_mode = MtsStatus::NotAnMTSObject;
                } else if value == "1" {
                    mts_transaction_mode = MtsStatus::NoTransactions;
                } else if value == "2" {
                    mts_transaction_mode = MtsStatus::RequiresTransaction;
                } else if value == "3" {
                    mts_transaction_mode = MtsStatus::UsesTransaction;
                } else if value == "4" {
                    mts_transaction_mode = MtsStatus::RequiresNewTransaction;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::Property(
                        PropertyError::InvalidPropertyValueZeroNegOne,
                    )));
                }
            }
            _ => {
                return Err(ErrMode::Cut(VB6ErrorKind::Property(
                    PropertyError::UnknownProperty,
                )));
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
Attribute VB_Ext_KEY = \"SavedWithClassBuilder6\" ,\"Yes\"
Attribute VB_Ext_KEY = \"Saved\" ,\"False\"

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
                    Attribute VB_Description = \"Description text\"\r
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
