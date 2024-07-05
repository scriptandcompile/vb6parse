#![warn(clippy::pedantic)]

use crate::vb6::{vb6_parse, VB6Token};

use winnow::{
    ascii::{line_ending, space0, space1, Caseless},
    error::{ContextError, ErrMode},
    token::{literal, take_until},
    Parser,
};

/// Represents a VB6 module file.
/// A VB6 module file contains a header and a list of tokens.
///
/// The tokens contain the token stream of the code of the class file.
#[derive(Debug, PartialEq, Eq, Clone)]
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
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::module::VB6ModuleFile;
    ///
    /// let input = b"Attribute VB_Name = \"Module1\"
    /// Option Explicit
    ///
    /// Private Sub Class_Initialize()
    /// End Sub
    /// ";
    ///
    /// let result = VB6ModuleFile::parse(input);
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(input: &'a [u8]) -> Result<Self, ErrMode<ContextError>> {
        let mut input = input;

        (
            space0,
            literal(Caseless("Attribute")),
            space1,
            literal(Caseless("VB_Name")),
            space0,
            literal("="),
            space0,
        )
            .parse_next(&mut input)?;

        let name = (
            literal("\""),
            take_until(0.., "\""),
            literal("\""),
            space0,
            line_ending,
        )
            .recognize()
            .parse_next(&mut input)?;

        let tokens = vb6_parse(&mut input)?;

        Ok(VB6ModuleFile { name, tokens })
    }
}
