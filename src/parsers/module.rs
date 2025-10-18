use crate::{
    errors::VB6ModuleErrorKind, language::VB6Token, parsers::ParseResult, sourcefile::SourceFile,
    sourcestream::Comparator, vb6_code_tokenize,
};

use serde::Serialize;

/// Represents a VB6 module file.
/// A VB6 module files contain a header and a list of tokens.
///
/// The tokens contain the token stream of the code of the class file.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6ModuleFile<'a> {
    pub name: &'a [u8], // Attribute VB_Name = "Module1"
    pub tokens: Vec<VB6Token<'a>>,
}

impl<'a> VB6ModuleFile<'a> {
    /// Parses a VB6 module file from a byte slice.
    ///
    /// # Arguments
    ///
    /// * `input` The byte slice to parse.
    ///
    /// # Returns
    ///
    /// A result containing the parsed VB6 module file or an error.
    ///
    /// # Errors
    ///
    /// An error will be returned if the input is not a valid VB6 module file.
    ///
    /// # Panics
    ///
    /// This function will panic if the source code is not a valid module file.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::*;
    ///
    /// let input = b"Attribute VB_Name = \"Module1\"
    /// Option Explicit
    ///
    /// Private Sub Class_Initialize()
    /// End Sub
    /// ";
    ///
    /// let source_file = match SourceFile::decode_with_replacement("module.bas", input) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode module source code.");
    ///     }
    /// };
    ///
    /// let result = VB6ModuleFile::parse(&source_file);
    ///
    /// if result.has_failures() {
    ///     for failure in result.failures {
    ///         failure.print();
    ///     }
    ///     panic!("Module parse had failures");
    /// }
    ///
    /// let module_file = result.unwrap();
    ///
    /// assert_eq!(module_file.name, "Module1".as_bytes());
    /// assert_eq!(module_file.tokens.len(), 18);
    /// ```
    pub fn parse(
        source_file: &'a SourceFile,
    ) -> ParseResult<'a, VB6ModuleFile<'a>, VB6ModuleErrorKind> {
        let mut failures = vec![];
        let mut input = source_file.get_source_stream();

        // Eat however many spaces starts the files. It doesn't matter how many
        // whitespaces it has, zero or many.
        let _ = input.take_ascii_whitespaces();

        // Grab the Attribute keyword. If we don't find it, we should output an error
        // but keep trying to read on.
        if input
            .take("Attribute", Comparator::CaseInsensitive)
            .is_none()
        {
            let error = input.generate_error(VB6ModuleErrorKind::AttributeKeywordMissing);
            failures.push(error);
        };

        // Eat however many spaces sits between the attribute and the VB_Name keyword. It doesn't matter how many
        // whitespaces it has as long as we have at least one.
        if input.take_ascii_whitespaces().is_none() {
            let error = input.generate_error(VB6ModuleErrorKind::MissingWhitespaceInHeader);
            failures.push(error);
        }

        // Grab the attribute VB_Name keyword. If we don't find it, we should output an error
        // but keep trying to read on.
        if input.take("VB_Name", Comparator::CaseInsensitive).is_none() {
            let error = input.generate_error(VB6ModuleErrorKind::VBNameAttributeMissing);
            failures.push(error);
        };

        // Eat however many spaces starts the files. It doesn't matter how many
        // whitespaces it has, zero or many.
        let _ = input.take_ascii_whitespaces();

        // Grab the equality symbol. If we don't find it, we should output an error
        // but keep trying to read on.
        if input.take("=", Comparator::CaseInsensitive).is_none() {
            let error = input.generate_error(VB6ModuleErrorKind::EqualMissing);
            failures.push(error);
        };

        // Eat however many spaces starts the files. It doesn't matter how many
        // whitespaces it has, zero or many.
        let _ = input.take_ascii_whitespaces();

        // Grab the quote symbol. If we don't find it, we should output an error
        // but keep trying to read on.
        if input.take("\"", Comparator::CaseInsensitive).is_none() {
            let error = input.generate_error(VB6ModuleErrorKind::VBNameAttributeValueUnquoted);
            failures.push(error);
        };

        match input.take_until("\"", Comparator::CaseInsensitive) {
            None => {
                // Well, it looks like we don't have a quoted value even if it might have a single quote at the start.
                let Some((vb_name_value, _)) = input.take_until_newline() else {
                    let error =
                        input.generate_error(VB6ModuleErrorKind::VBNameAttributeValueUnquoted);
                    failures.push(error);

                    let parse_result = vb6_code_tokenize(&mut input);

                    if parse_result.has_failures() {
                        for failure in parse_result.failures {
                            failures.push(failure.into());
                        }
                    }
                    input.take_newline();

                    return ParseResult {
                        result: None,
                        failures,
                    };
                };

                let parse_result = vb6_code_tokenize(&mut input);

                if parse_result.has_failures() {
                    for failure in parse_result.failures {
                        failures.push(failure.into());
                    }
                }

                match parse_result.result {
                    Some(tokens) => {
                        return ParseResult {
                            result: Some(VB6ModuleFile {
                                name: vb_name_value.as_bytes(),
                                tokens,
                            }),
                            failures,
                        };
                    }
                    None => {
                        return ParseResult {
                            result: None,
                            failures,
                        };
                    }
                }
            }
            Some((vb_name_value, _)) => {
                // eat the quote character we found.
                let _ = input.take_count(1);
                // we might have whitespaces after the quoted value.
                // we don't care about them so just eat them and then the newline.
                let _ = input.take_ascii_whitespaces();
                let _ = input.take_newline();

                // looks like we have a fully quoted value.
                let parse_result = vb6_code_tokenize(&mut input);

                if parse_result.has_failures() {
                    for failure in parse_result.failures {
                        failures.push(failure.into());
                    }
                }

                match parse_result.result {
                    Some(tokens) => {
                        return ParseResult {
                            result: Some(VB6ModuleFile {
                                name: vb_name_value.as_bytes(),
                                tokens,
                            }),
                            failures,
                        };
                    }
                    None => {
                        return ParseResult {
                            result: None,
                            failures,
                        };
                    }
                }
            }
        }
    }
}
