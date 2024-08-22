#![warn(clippy::pedantic)]

use bstr::{BStr, ByteSlice};

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    language::VB6Token,
    parsers::{
        header::{key_value_line_parse, version_parse, HeaderKind, VB6FileFormatVersion},
        VB6Stream,
    },
    vb6::{keyword_parse, line_comment_parse, vb6_parse, VB6Result},
};

use winnow::{
    ascii::{line_ending, space0},
    combinator::{alt, preceded, repeat_till},
    error::ErrMode,
    Parser,
};

/// Represents the COM usage of a class file.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum FileUsage {
    // In a COM object a MultiUse class object will be created for all clients.
    // This value is stored as -1 (true) in the file.
    MultiUse,
    // In a COM object a SingleUse class object will be created for each client.
    // This value is stored as 0 (false) in the file.
    SingleUse, // 0 (false)
}

/// Represents the persistability of a file.
/// Only available when the class is part of an ActiveX DLL project that is both
/// public and creatable.
///
/// Determines whether the class can be saved to disk.
///
/// If it is Persistable, four procedures: InitProperties, ReadProperties, and
/// WriteProperties events, and the PropertyChanged method are automatically
/// added to the class module.
///
/// Without these procedures, the class cannot be saved to disk.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Persistance {
    // The class property cannot be saved to a file in a property bag.
    // This value is stored as 0 (false) in the file.
    NonPersistable,
    // The class property can be saved to a file in a property bag.
    // This value is stored as 1 (true) in the file.
    Persistable,
}

/// Represents the MTS status of a file.
///
/// Only available when the class is part of an ActiveX DLL project. This should
/// be set to values other than NotAnMTSObject (0) if the class is to be used as
/// a Microsoft Transaction Server component.
///
/// Maps directly to the MTS transaction mode attribute in Microsoft Transaction
/// Server.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum MtsStatus {
    // This class is not an MTS component.
    // This value is stored as 0 in the file.
    // This is the default value.
    NotAnMTSObject,
    // This class is an MTS component but does not support transactions.
    // This value is stored as 1 in the file.
    NoTransactions,
    // This class is an MTS component and requires a transaction.
    // This value is stored as 2 in the file.
    RequiresTransaction,
    // This class is an MTS component and uses a transaction.
    // This value is stored as 3 in the file.
    UsesTransaction,
    // This class is an MTS component and requires a new transaction.
    // This value is stored as 4 in the file.
    RequiresNewTransaction,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DataSourceBehavior {
    // The class does not support acting as a Data Source.
    // This value is stored as 0 in the file.
    None,
    // The class supports acting as a Data Source.
    // This value is stored as 1 in the file.
    DataSource,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DataBindingBehavior {
    // The class does not support data binding.
    // This value is stored as 0 in the file.
    None,
    // The class supports simple data binding.
    // This value is stored as 1 in the file.
    Simple,
    // The class supports complex data binding.
    // This value is stored as 2 in the file.
    Complex,
}

/// The properties of a VB6 class file is the list of key/value pairs
/// found between the BEGIN and END lines in the header.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassProperties {
    // (0/-1) multi use / single use
    pub multi_use: FileUsage,
    // (0/1) NonParsistable / Persistable
    pub persistable: Persistance,
    // (0/1/2) vbNone / vbSimple / vbComplex
    pub data_binding_behavior: DataBindingBehavior,
    // (0/1) vbNone / vbDataSource
    pub data_source_behavior: DataSourceBehavior,
    // (0/1/2/3/4) NotAnMTSObject / NoTransactions / RequiresTransaction / UsesTransaction / RequiresNewTransaction
    pub mts_transaction_mode: MtsStatus,
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
fn properties_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ClassProperties> {
    begin_line_parse.parse_next(input)?;

    let mut multi_use = FileUsage::MultiUse;
    let mut persistable = Persistance::NonPersistable;
    let mut data_binding_behavior = DataBindingBehavior::None;
    let mut data_source_behavior = DataSourceBehavior::None;
    let mut mts_transaction_mode = MtsStatus::NotAnMTSObject;

    let (collection, _): (Vec<(&BStr, &BStr)>, _) =
        repeat_till(0.., key_value_line_parse("="), end_line_parse).parse_next(input)?;

    for pair in collection.iter() {
        let (key, value) = *pair;

        match key.as_bytes() {
            b"Persistable" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value == "-1" {
                    persistable = Persistance::Persistable;
                } else if value == "0" {
                    persistable = Persistance::NonPersistable;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueZeroNegOne));
                }
            }
            b"MultiUse" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value == "-1" {
                    multi_use = FileUsage::MultiUse;
                } else if value == "0" {
                    multi_use = FileUsage::SingleUse;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueZeroNegOne));
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
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueZeroNegOne));
                }
            }
            b"DataSourceBehavior" => {
                if value == "0" {
                    data_source_behavior = DataSourceBehavior::None;
                } else if value == "1" {
                    data_source_behavior = DataSourceBehavior::DataSource;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueZeroNegOne));
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
                    mts_transaction_mode = MtsStatus::RequiresNewTransaction
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueZeroNegOne));
                }
            }
            _ => {
                return Err(ErrMode::Cut(VB6ErrorKind::UnknownProperty));
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

fn attributes_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6FileAttributes<'a>> {
    let _ = space0::<_, VB6Error>.parse_next(input);

    let mut name = Option::None;
    let mut global_name_space = false;
    let mut creatable = false;
    let mut pre_declared_id = false;
    let mut exposed = false;

    while let Ok((key, value)) =
        preceded(keyword_parse("Attribute"), key_value_line_parse("=")).parse_next(input)
    {
        match key.as_bytes() {
            b"VB_Name" => {
                name = Some(value);
            }
            b"VB_GlobalNameSpace" => {
                if value == "True" {
                    global_name_space = true;
                } else if value == "False" {
                    global_name_space = false;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueTrueFalse));
                }
            }
            b"VB_Creatable" => {
                if value == "True" {
                    creatable = true;
                } else if value == "False" {
                    creatable = false;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueTrueFalse));
                }
            }
            b"VB_PredeclaredId" => {
                if value == "True" {
                    pre_declared_id = true;
                } else if value == "False" {
                    pre_declared_id = false;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueTrueFalse));
                }
            }
            b"VB_Exposed" => {
                if value == "True" {
                    exposed = true;
                } else if value == "False" {
                    exposed = false;
                } else {
                    return Err(ErrMode::Cut(VB6ErrorKind::InvalidPropertyValueTrueFalse));
                }
            }
            _ => {
                return Err(ErrMode::Cut(VB6ErrorKind::UnknownProperty));
            }
        }
    }

    if name.is_none() {
        return Err(ErrMode::Cut(VB6ErrorKind::MissingClassName));
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