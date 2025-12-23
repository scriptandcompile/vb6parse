//! VB6 Module File Parser Module
//!
//! This module provides functionality to parse VB6 module files (.bas).
//! It defines the `ModuleFile` struct representing a VB6 module file,
//! along with methods to parse the file and extract relevant information
//! such as the module name and its concrete syntax tree (CST).
//!

use std::fmt::Display;

use crate::{
    errors::ModuleErrorKind,
    parsers::{cst::serialize_cst, cst::ConcreteSyntaxTree, ParseResult, SyntaxKind},
    sourcefile::SourceFile,
    sourcestream::Comparator,
    tokenize::tokenize,
};

use serde::Serialize;

/// Represents a VB6 module file.
/// A VB6 module file contains a header and a concrete syntax tree.
///
/// The CST contains the parsed structure of the module code.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct ModuleFile {
    /// The name of the module.
    pub name: String, // Attribute VB_Name = "Module1"
    /// The concrete syntax tree of the module file.
    #[serde(serialize_with = "serialize_cst")]
    pub cst: ConcreteSyntaxTree,
}

impl Display for ModuleFile {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "VB6 Module File: {}", self.name)
    }
}

impl ModuleFile {
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
    /// let result = ModuleFile::parse(&source_file);
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
    /// assert_eq!(module_file.name, "Module1");
    /// assert!(module_file.cst.child_count() > 0);
    /// ```
    #[must_use]
    pub fn parse(source_file: &SourceFile) -> ParseResult<'_, ModuleFile, ModuleErrorKind> {
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
            let error = input.generate_error(ModuleErrorKind::AttributeKeywordMissing);
            failures.push(error);
        }

        // Eat however many spaces sits between the attribute and the VB_Name keyword. It doesn't matter how many
        // whitespaces it has as long as we have at least one.
        if input.take_ascii_whitespaces().is_none() {
            let error = input.generate_error(ModuleErrorKind::MissingWhitespaceInHeader);
            failures.push(error);
        }

        // Grab the attribute VB_Name keyword. If we don't find it, we should output an error
        // but keep trying to read on.
        if input.take("VB_Name", Comparator::CaseInsensitive).is_none() {
            let error = input.generate_error(ModuleErrorKind::VBNameAttributeMissing);
            failures.push(error);
        }

        // Eat however many spaces starts the files. It doesn't matter how many
        // whitespaces it has, zero or many.
        let _ = input.take_ascii_whitespaces();

        // Grab the equality symbol. If we don't find it, we should output an error
        // but keep trying to read on.
        if input.take("=", Comparator::CaseInsensitive).is_none() {
            let error = input.generate_error(ModuleErrorKind::EqualMissing);
            failures.push(error);
        }

        // Eat however many spaces starts the files. It doesn't matter how many
        // whitespaces it has, zero or many.
        let _ = input.take_ascii_whitespaces();

        // Grab the quote symbol. If we don't find it, we should output an error
        // but keep trying to read on.
        if input.take("\"", Comparator::CaseInsensitive).is_none() {
            let error = input.generate_error(ModuleErrorKind::VBNameAttributeValueUnquoted);
            failures.push(error);
        }

        match input.take_until("\"", Comparator::CaseInsensitive) {
            None => {
                // Well, it looks like we don't have a quoted value even if it might have a single quote at the start.
                let Some((vb_name_value, _)) = input.take_until_newline() else {
                    let error = input.generate_error(ModuleErrorKind::VBNameAttributeValueUnquoted);
                    failures.push(error);

                    return ParseResult {
                        result: None,
                        failures,
                    };
                };

                // Parse the entire source file as CST
                let mut stream = source_file.get_source_stream();
                let token_result = tokenize(&mut stream);

                if token_result.has_failures() {
                    for failure in token_result.failures {
                        failures.push(failure.into());
                    }
                }

                match token_result.result {
                    Some(tokens) => {
                        let cst = crate::parsers::cst::parse(tokens);

                        // Filter out nodes that are already extracted to avoid duplication
                        let filtered_cst = cst.without_kinds(&[SyntaxKind::AttributeStatement]);

                        ParseResult {
                            result: Some(ModuleFile {
                                name: vb_name_value.to_string(),
                                cst: filtered_cst,
                            }),
                            failures,
                        }
                    }
                    None => ParseResult {
                        result: None,
                        failures,
                    },
                }
            }
            Some((vb_name_value, _)) => {
                // Eat the quote character we found.
                let _ = input.take_count(1);
                // We might have whitespaces after the quoted value.
                // We don't care about them so just eat them and then the newline.
                let _ = input.take_ascii_whitespaces();
                let _ = input.take_newline();

                // Looks like we have a fully quoted value.
                // Parse the remaining source file as CST
                let token_result = tokenize(&mut input);

                if token_result.has_failures() {
                    for failure in token_result.failures {
                        failures.push(failure.into());
                    }
                }

                match token_result.result {
                    Some(tokens) => {
                        let cst = crate::parsers::cst::parse(tokens);

                        // Filter out nodes that are already extracted to avoid duplication
                        let filtered_cst = cst.without_kinds(&[SyntaxKind::AttributeStatement]);

                        ParseResult {
                            result: Some(ModuleFile {
                                name: vb_name_value.to_string(),
                                cst: filtered_cst,
                            }),
                            failures,
                        }
                    }
                    None => ParseResult {
                        result: None,
                        failures,
                    },
                }
            }
        }
    }
}
