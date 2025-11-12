pub mod properties;

use std::collections::HashMap;

use serde::Serialize;

use crate::{
    errors::{PropertyError, VB6ClassErrorKind},
    language::VB6Token,
    parsers::{
        class::properties::{
            DataBindingBehavior, DataSourceBehavior, FileUsage, MtsStatus, Persistence,
            VB6ClassHeader, VB6ClassProperties,
        },
        header::{
            self, attributes_parse, version_parse, Creatable, Exposed, HeaderKind, NameSpace,
            PreDeclaredID, VB6FileAttributes, VB6FileFormatVersion,
        },
    },
    vb6code::tokenize,
    ParseResult, SourceFile, SourceStream, VB6Tokenizer,
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
    pub tokens: Vec<(&'a str, VB6Token)>,
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
    /// use vb6parse::SourceFile;
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
    ///
    /// let result = SourceFile::decode_with_replacement("class_parse.cls", input);
    ///
    /// let source_file = match result {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => panic!("Failed to decode source file 'class_parse.cls': {e:?}"),
    /// };
    ///
    ///
    /// let result = VB6ClassFile::parse(&source_file);
    ///
    /// assert!(result.has_result());
    /// ```
    #[must_use]
    pub fn parse(source_file: &'a SourceFile) -> ParseResult<'a, Self, VB6ClassErrorKind<'a>> {
        let mut input = source_file.get_source_stream();

        let mut failures = vec![];

        let version_result = version_header_parse(&mut input, HeaderKind::Class);

        let Some(version) = version_result.result else {
            for failure in version_result.failures {
                failures.push(failure);
            }

            return ParseResult {
                result: None,
                failures,
            };
        };

        let properties_result = properties_parse(&mut input);

        let Some(properties) = properties_result.result else {
            for failure in properties_result.failures {
                failures.push(failure);
            }

            return ParseResult {
                result: None,
                failures,
            };
        };

        let header = VB6ClassHeader {
            version,
            attributes: VB6FileAttributes {
                name: b" ".into(),
                global_name_space: NameSpace::default(),
                creatable: Creatable::default(),
                pre_declared_id: PreDeclaredID::default(),
                exposed: Exposed::default(),
                description: None,
                ext_key: HashMap::new(),
            },
            properties,
        };

        let code_parse_result = tokenize(&mut input);

        if code_parse_result.has_failures() {
            for failure in code_parse_result.failures {
                failures.push(failure.into());
            }
        }

        let Some(tokens) = code_parse_result.result else {
            return ParseResult {
                result: None,
                failures,
            };
        };

        ParseResult {
            result: Some(VB6ClassFile { header, tokens }),
            failures,
        }
    }
}

fn version_header_parse<'a>(
    input: &mut SourceStream<'a>,
    header_kind: HeaderKind,
) -> ParseResult<'a, VB6FileFormatVersion, VB6ClassErrorKind<'a>> {
    let mut failures = vec![];

    // eat any whitespaces before the version keyword.
    let _ = input.take_ascii_whitespaces();

    let version_start_offset = input.offset();
    let Some(version_keyword) = input.take_matching_text("VERSION") else {
        let error = input.generate_error(VB6ClassErrorKind::VersionKeywordMissing);
        failures.push(error);

        return ParseResult {
            result: None,
            failures,
        };
    };

    // Not really an error, more of a warning issue, but not correctly using
    // full upper case on 'VERSION' could mean the file won't be compatible
    // with Microsoft VB6 IDE.
    if version_keyword != "VERSION" {
        let error = input.generate_error_at(
            version_start_offset,
            VB6ClassErrorKind::VersionKeywordNotFullyUppercase {
                version_text: version_keyword,
            },
        );
        failures.push(error);
    }

    // eat any whitespaces after the version keyword.
    // If there isn't a space between 'VERSION' and the major version it should
    // be a warning since we can continue reasonably well even if it will not be
    // fully compatible with Microsoft's VB6 IDE.
    if input.take_ascii_whitespaces().is_none() {
        let error = input.generate_error(
            VB6ClassErrorKind::WhitespaceMissingBetweenVersionAndMajorVersionNumber,
        );
        failures.push(error);
    }

    let major_digit_offset = input.offset();
    // sadly, we can't really continue if we can't get the version numbers since
    // we can't be sure this is even a VB6 class file.
    let Some(major_version_digits) = input.take_ascii_digits() else {
        let error = input.generate_error(VB6ClassErrorKind::UnableToParseMajorVersionNumber);
        failures.push(error);

        return ParseResult {
            result: None,
            failures,
        };
    };

    // If we can't convert the major digits text into a number we have some weird
    // issue and need to bail out at this point. This *should* always work as far
    // as I can tell, but this covers our bases.
    let Ok(major) = major_version_digits.parse::<u8>() else {
        let error = input.generate_error_at(
            major_digit_offset,
            VB6ClassErrorKind::UnableToConvertMajorVersionNumber,
        );
        failures.push(error);

        return ParseResult {
            result: None,
            failures,
        };
    };

    // Technically, the major/minor versions are supposed to be
    // separated by a period (1.0 for example). But we want to be as liberal
    // in acceptance as possible so we also support whitespaces between the
    // digits.
    if input
        .take(".", crate::Comparator::CaseInsensitive)
        .is_none()
    {
        let error = input
            .generate_error(VB6ClassErrorKind::MissingPeriodDividerBetweenMajorAndMinorVersion);
        failures.push(error);

        let whitespace_divider_offset = input.offset();
        if input.take_ascii_whitespaces().is_some() {
            let error = input.generate_error_at(
                whitespace_divider_offset,
                VB6ClassErrorKind::WhitespaceDividerBetweenMajorAndMinorVersionNumbers,
            );
            failures.push(error);
        }
    };

    let minor_digit_offset = input.offset();
    // sadly, we can't really continue if we can't get the version numbers since
    // we can't be sure this is even a VB6 class file.
    let Some(minor_version_digits) = input.take_ascii_digits() else {
        let error = input.generate_error(VB6ClassErrorKind::UnableToParseMinorVersionNumber);
        failures.push(error);

        return ParseResult {
            result: None,
            failures,
        };
    };

    // If we can't convert the minor digits text into a number we have some weird
    // issue and need to bail out at this point. This *should* always work as far
    // as I can tell, but this covers our bases.
    let Ok(minor) = minor_version_digits.parse::<u8>() else {
        let error = input.generate_error_at(
            minor_digit_offset,
            VB6ClassErrorKind::UnableToConvertMinorVersionNumber,
        );
        failures.push(error);

        return ParseResult {
            result: None,
            failures,
        };
    };

    let whitespace_offset = input.offset();
    match input.take_ascii_whitespaces() {
        None => {
            let error = input.generate_error(VB6ClassErrorKind::MissingWhitespaceAfterMinorVersion);
            failures.push(error);
        }
        Some(whitespace) => {
            if whitespace != " " {
                let error = input.generate_error_at(
                    whitespace_offset,
                    VB6ClassErrorKind::IncorrectWhitespaceAfterMinorVersion,
                );
                failures.push(error);
            }
        }
    }

    let match_text = match header_kind {
        HeaderKind::Class => "CLASS",
        HeaderKind::Form => "FORM",
    };

    let match_text_start_offset = input.offset();
    let Some(match_text_keyword) = input.take_matching_text(match_text) else {
        let error = match header_kind {
            HeaderKind::Class => input.generate_error(VB6ClassErrorKind::ClassKeywordMissing),
            // TODO: Correct this to a 'Form' keyword missing error message when I get a chance.
            HeaderKind::Form => input.generate_error(VB6ClassErrorKind::ClassKeywordMissing),
        };
        failures.push(error);

        return ParseResult {
            result: None,
            failures,
        };
    };

    // Not really an error, more of a warning issue, but not correctly using
    // full upper case on 'CLASS' or 'FORM' could mean the file won't be compatible
    // with Microsoft VB6 IDE.
    if match_text_keyword != match_text {
        let error = match header_kind {
            HeaderKind::Class => input.generate_error_at(
                match_text_start_offset,
                VB6ClassErrorKind::ClassKeywordNotFullyUppercase {
                    class_text: match_text_keyword,
                },
            ),
            // TODO: Correct this to a 'Form' keyword not fully uppercase error message when I get a chance.
            HeaderKind::Form => input.generate_error_at(
                match_text_start_offset,
                VB6ClassErrorKind::ClassKeywordNotFullyUppercase {
                    class_text: match_text_keyword,
                },
            ),
        };
        failures.push(error);
    }

    let _ = input.take_ascii_whitespaces();

    let _ = input.take_newline();

    ParseResult {
        result: Some(VB6FileFormatVersion { major, minor }),
        failures,
    }
}

/// Parses a VB6 class file properties from the header, including the
/// BEGIN and END lines.
/// The properties are the key/value pairs found between the BEGIN and END lines
/// in the header.
/// The properties contain the multi use, persistability, data binding behavior,
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
fn properties_parse<'a>(
    input: &mut SourceStream<'a>,
) -> ParseResult<'a, VB6ClassProperties, VB6ClassErrorKind<'a>> {
    let mut failures = vec![];

    let class_properties = VB6ClassProperties::default();

    // eat any whitespaces before the 'BEGIN' keyword.
    let _ = input.take_ascii_whitespaces();

    let begin_start_offset = input.offset();
    let Some(begin_keyword) = input.take_matching_text("BEGIN") else {
        let error = input.generate_error(VB6ClassErrorKind::BeginKeywordMissing);
        failures.push(error);

        return ParseResult {
            result: None,
            failures,
        };
    };

    // Not really an error, more of a warning issue, but not correctly using
    // full upper case on 'BEGIN' could mean the file won't be compatible
    // with Microsoft VB6 IDE.
    if begin_keyword != "BEGIN" {
        let error = input.generate_error_at(
            begin_start_offset,
            VB6ClassErrorKind::BeginKeywordNotFullyUppercase {
                begin_text: begin_keyword,
            },
        );
        failures.push(error);
    }

    // eat any whitespace after the 'BEGIN'
    let _ = input.take_ascii_whitespaces();

    let possible_comment = input.peek(1);

    // We want to eat to the end of the line any comments and move to the next line.
    // if it's a carriage return or newline, we skip over it.
    if possible_comment.is_some_and(|single_character| {
        single_character == "'" || single_character == "\r" || single_character == "\n"
    }) {
        let _ = input.take_until_newline();
    } else {
        let error = input.generate_error(VB6ClassErrorKind::BeginKeywordShouldBeStandAlone);
        failures.push(error);

        return ParseResult {
            result: None,
            failures,
        };
    }

    ParseResult {
        result: Some(class_properties),
        failures,
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
// fn class_header_parse<'a>(input: &mut SourceStream<'a>) -> ParseResult<VB6ClassHeader<'a>> {
//     // VERSION #.# CLASS
//     // BEGIN
//     //  key = value  'comment
//     //  ...
//     // END

//     let version = version_parse(HeaderKind::Class).parse_next(input)?;

//     let properties = properties_parse.parse_next(input)?;

//     let attributes = attributes_parse.parse_next(input)?;

//     Ok(VB6ClassHeader {
//         version,
//         properties,
//         attributes,
//     })
// }

// fn begin_line_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
//     let (_, begin, _) = (space0, keyword_parse("BEGIN"), space0).parse_next(input)?;

//     alt(((line_comment_parse, line_ending), (space0, line_ending))).parse_next(input)?;

//     Ok(begin)
// }

// fn end_line_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
//     let (_, keyword, _) = (space0, keyword_parse("END"), space0).parse_next(input)?;

//     alt(((line_comment_parse, line_ending), (space0, line_ending))).parse_next(input)?;

//     Ok(keyword)
// }

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
// fn properties_parse(input: &mut VB6Stream<'_>) -> VB6Result<VB6ClassProperties> {
//     begin_line_parse.parse_next(input)?;

//     let mut multi_use = FileUsage::MultiUse;
//     let mut persistable = Persistence::NotPersistable;
//     let mut data_binding_behavior = DataBindingBehavior::None;
//     let mut data_source_behavior = DataSourceBehavior::None;
//     let mut mts_transaction_mode = MtsStatus::NotAnMTSObject;

//     let (collection, _): (Vec<(&BStr, &BStr)>, _) =
//         repeat_till(0.., key_value_line_parse("="), end_line_parse).parse_next(input)?;

//     for pair in &collection {
//         let (key, value) = *pair;

//         match key.as_bytes() {
//             b"Persistable" => {
//                 // -1 is 'true' and 0 is 'false' in VB6
//                 if value == "-1" {
//                     persistable = Persistence::Persistable;
//                 } else if value == "0" {
//                     persistable = Persistence::NotPersistable;
//                 } else {
//                     return Err(ErrMode::Cut(VB6ErrorKind::Property(
//                         PropertyError::InvalidPropertyValueZeroNegOne,
//                     )));
//                 }
//             }
//             b"MultiUse" => {
//                 // -1 is 'true' and 0 is 'false' in VB6
//                 if value == "-1" {
//                     multi_use = FileUsage::MultiUse;
//                 } else if value == "0" {
//                     multi_use = FileUsage::SingleUse;
//                 } else {
//                     return Err(ErrMode::Cut(VB6ErrorKind::Property(
//                         PropertyError::InvalidPropertyValueZeroNegOne,
//                     )));
//                 }
//             }
//             b"DataBindingBehavior" => {
//                 if value == "0" {
//                     data_binding_behavior = DataBindingBehavior::None;
//                 } else if value == "1" {
//                     data_binding_behavior = DataBindingBehavior::Simple;
//                 } else if value == "2" {
//                     data_binding_behavior = DataBindingBehavior::Complex;
//                 } else {
//                     return Err(ErrMode::Cut(VB6ErrorKind::Property(
//                         PropertyError::InvalidPropertyValueZeroNegOne,
//                     )));
//                 }
//             }
//             b"DataSourceBehavior" => {
//                 if value == "0" {
//                     data_source_behavior = DataSourceBehavior::None;
//                 } else if value == "1" {
//                     data_source_behavior = DataSourceBehavior::DataSource;
//                 } else {
//                     return Err(ErrMode::Cut(VB6ErrorKind::Property(
//                         PropertyError::InvalidPropertyValueZeroNegOne,
//                     )));
//                 }
//             }
//             b"MTSTransactionMode" => {
//                 if value == "0" {
//                     mts_transaction_mode = MtsStatus::NotAnMTSObject;
//                 } else if value == "1" {
//                     mts_transaction_mode = MtsStatus::NoTransactions;
//                 } else if value == "2" {
//                     mts_transaction_mode = MtsStatus::RequiresTransaction;
//                 } else if value == "3" {
//                     mts_transaction_mode = MtsStatus::UsesTransaction;
//                 } else if value == "4" {
//                     mts_transaction_mode = MtsStatus::RequiresNewTransaction;
//                 } else {
//                     return Err(ErrMode::Cut(VB6ErrorKind::Property(
//                         PropertyError::InvalidPropertyValueZeroNegOne,
//                     )));
//                 }
//             }
//             _ => {
//                 return Err(ErrMode::Cut(VB6ErrorKind::Property(
//                     PropertyError::UnknownProperty,
//                 )));
//             }
//         }
//     }

//     Ok(VB6ClassProperties {
//         multi_use,
//         persistable,
//         data_binding_behavior,
//         data_source_behavior,
//         mts_transaction_mode,
//     })
// }

#[cfg(test)]
mod tests {
    //use super::HeaderKind;
    use super::*;

    #[test]
    fn class_file_valid() {
        let class_bytes = r#"VERSION 1.0 CLASS
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
"#;

        let result = SourceFile::decode_with_replacement("test.cls", &class_bytes.as_bytes());

        let source_file = match result {
            Ok(source_file) => source_file,
            Err(e) => panic!("Failed to decode source file 'test.cls': {e:?}"),
        };

        let result = VB6ClassFile::parse(&source_file);

        if result.has_failures() {
            for failure in result.failures {
                failure.print();
            }

            panic!("Class parse had failures");
        }

        assert!(result.has_result());
    }

    #[test]
    fn class_file_invalid() {
        // These should be '\r\n', or at worst '\n', but not '\r'.
        // '\r' is not even remotely a valid line ending.
        let class_bytes = b"VERSION 1.0 CLASS\r
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

        let result = SourceFile::decode_with_replacement("test.cls", class_bytes);

        let source_file = match result {
            Ok(source_file) => source_file,
            Err(e) => panic!("Failed to decode source file 'test.cls': {e:?}"),
        };

        let result = VB6ClassFile::parse(&source_file);

        assert!(result.has_failures());
    }

    //     #[test]
    //     fn class_header_valid() {
    //         let input = b"VERSION 1.0 CLASS\r
    // BEGIN\r
    //     MultiUse = -1  'True\r
    //     Persistable = 0  'NotPersistable\r
    //     DataBindingBehavior = 0  'vbNone\r
    //     DataSourceBehavior = 0  'vbNone\r
    //     MTSTransactionMode = 0  'NotAnMTSObject\r
    // END\r
    // Attribute VB_Name = \"Something\"\r
    // Attribute VB_GlobalNameSpace = False\r
    // Attribute VB_Creatable = True\r
    // Attribute VB_PredeclaredId = False\r
    // Attribute VB_Exposed = False";

    //         assert!(result.is_ok());
    //     }

    //     #[test]
    //     fn class_header_invalid() {
    //         let input = b"MultiUse = -1  'True\r
    //     Persistable = 0  'NotPersistable\r
    //     DataBindingBehavior = 0  'vbNone\r
    //     DataSourceBehavior = 0  'vbNone\r
    //     MTSTransactionMode = 0  'NotAnMTSObject\r
    //     ";

    //         let mut stream = VB6Stream::new("", &mut input.as_slice());
    //         let result = class_header_parse(&mut stream);

    //         assert!(result.is_err());
    //     }

    //     #[test]
    //     fn attributes_valid() {
    //         let input = b"Attribute VB_Name = \"Something\"\r
    //     Attribute VB_GlobalNameSpace = False\r
    //     Attribute VB_Creatable = True\r
    //     Attribute VB_PredeclaredId = False\r
    //     Attribute VB_Exposed = False\r
    //     ";

    //         let mut stream = VB6Stream::new("", &mut input.as_slice());
    //         let result = attributes_parse(&mut stream);

    //         assert!(result.is_ok());
    //     }

    //     #[test]
    //     fn attributes_invalid() {
    //         let input = b"Attribut VB_Name = \"Something\"\r
    //     Attrbute VB_GlobalNameSpace = False\r
    //     Attribut VB_Creatable = True\r
    //     Attriute VB_PredeclaredId = False\r
    //     Atribute VB_Exposed = False\r
    //     ";

    //         let mut stream = VB6Stream::new("", &mut input.as_slice());
    //         let result = attributes_parse(&mut stream);

    //         assert!(result.is_err());
    //     }

    #[test]
    fn version_valid() {
        let class_bytes = b"VERSION 1.0 CLASS\r\n";

        let result = SourceFile::decode_with_replacement("test.cls", class_bytes);

        let source_file = match result {
            Ok(source_file) => source_file,
            Err(e) => panic!("Failed to decode source file 'test.cls': {e:?}"),
        };

        let mut source_stream = source_file.get_source_stream();

        let result = version_header_parse(&mut source_stream, HeaderKind::Class);

        assert!(result.has_result());
        assert!(!result.has_failures());
    }

    #[test]
    fn version_invalid() {
        // 'VERSION' isn't correct
        let class_bytes = b"VERION 1.0 CLASS";

        let result = SourceFile::decode_with_replacement("test.cls", class_bytes);

        let source_file = match result {
            Ok(source_file) => source_file,
            Err(e) => panic!("Failed to decode source file 'test.cls': {e:?}"),
        };

        let mut source_stream = source_file.get_source_stream();

        let result = version_header_parse(&mut source_stream, HeaderKind::Class);

        assert!(!result.has_result());
        assert!(result.has_failures());
    }
}
