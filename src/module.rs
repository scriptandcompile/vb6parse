#![warn(clippy::pedantic)]

use crate::{
    errors::VB6Error,
    vb6::{keyword_parse, vb6_parse, VB6Token},
    vb6stream::VB6Stream,
};

use winnow::{
    ascii::{line_ending, space0, space1},
    token::take_until,
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
    /// let result = VB6ModuleFile::parse("module.bas".to_owned(), input);
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(file_name: String, input: &'a [u8]) -> Result<Self, VB6Error> {
        let mut input = VB6Stream::new(file_name, input);

        match (
            space0,
            keyword_parse("Attribute"),
            space1,
            keyword_parse("VB_Name"),
            space0,
            "=",
            space0,
        )
            .parse_next(&mut input)
        {
            Ok(_) => {}
            Err(e) => {
                return Err(e.into_inner().unwrap());
            }
        }

        let name = match ("\"", take_until(0.., "\""), "\"", space0, line_ending)
            .recognize()
            .parse_next(&mut input)
        {
            Ok(name) => name,
            Err(e) => {
                return Err(e.into_inner().unwrap());
            }
        };

        let tokens = match vb6_parse(&mut input) {
            Ok(tokens) => tokens,
            Err(e) => {
                return Err(e.into_inner().unwrap());
            }
        };

        Ok(VB6ModuleFile { name, tokens })
    }
}
