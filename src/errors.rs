//! Module containing the error types used in the VB6 parser.
//! It defines the `ErrorDetails` type, which is used to represent
//! errors that occur during parsing. The `ErrorDetails` type contains
//! information about the error, including the file name, source code,
//! source offset, column, line number, and the kind of error.
use core::convert::From;

use std::borrow::Cow;
use std::error::Error;
use std::fmt::Debug;

use ariadne::{Label, Report, ReportKind, Source};

/// Contains detailed information about an error that occurred during parsing.
/// This struct contains the source name, source content, error offset,
/// line start and end positions, and the kind of error.
///
/// Generic parameter `T` represents the type of error kind.
/// It must implement the `ToString` trait to allow for error message formatting.
///
/// Example usage:
/// ```rust
/// use vb6parse::errors::ErrorDetails;
/// use vb6parse::errors::CodeErrorKind;
/// use std::borrow::Cow;
///
/// let error_details = ErrorDetails {
///     source_name: "example.cls".to_string(),
///     source_content: Cow::Borrowed("Some VB6 code here..."),
///     error_offset: 10,
///     line_start: 1,
///     line_end: 1,
///     kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
/// };
/// error_details.print();
/// ```
#[derive(Debug, Clone)]
pub struct ErrorDetails<'a, T>
where
    T: ToString + Debug,
{
    /// The name of the source file where the error occurred.
    pub source_name: String,
    /// The content of the source file where the error occurred.
    pub source_content: Cow<'a, str>,
    /// The offset in the source content where the error occurred.
    pub error_offset: usize,
    /// The starting line number of the error.
    pub line_start: usize,
    /// The ending line number of the error.
    pub line_end: usize,
    /// The kind of error that occurred.
    pub kind: T,
}

impl<T> ErrorDetails<'_, T>
where
    T: ToString + Debug,
{
    /// Print the `ErrorDetails` using ariadne for formatting
    ///
    /// Example usage:
    /// ```rust
    /// use vb6parse::errors::ErrorDetails;
    /// use vb6parse::errors::CodeErrorKind;
    /// use std::borrow::Cow;
    ///
    /// let error_details = ErrorDetails {
    /// source_name: "example.cls".to_string(),
    ///   source_content: Cow::Borrowed("Some VB6 code here..."),
    ///   error_offset: 10,
    ///   line_start: 1,
    ///   line_end: 1,
    ///   kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    /// };
    /// error_details.print();
    /// ```
    pub fn print(&self) {
        let cache = (
            self.source_name.clone(),
            Source::from(self.source_content.to_string()),
        );

        let report = Report::build(
            ReportKind::Error,
            (self.source_name.clone(), self.line_start..=self.line_end),
        )
        .with_message(self.kind.to_string())
        .with_label(
            Label::new((
                self.source_name.clone(),
                self.error_offset..=self.error_offset,
            ))
            .with_message("error here"),
        )
        .finish()
        .print(cache);

        if let Some(e) = report.err() {
            eprint!("Error attempting to build ErrorDetails print message {e:?}");
        }
    }

    /// Eprint the `ErrorDetails` using ariadne for formatting
    ///
    /// Example usage:
    /// ```rust
    /// use vb6parse::errors::ErrorDetails;
    /// use vb6parse::errors::CodeErrorKind;
    /// use std::borrow::Cow;
    /// let error_details = ErrorDetails {
    ///     source_name: "example.cls".to_string(),
    ///     source_content: Cow::Borrowed("Some VB6 code here..."),
    ///     error_offset: 10,
    ///     line_start: 1,
    ///     line_end: 1,
    ///     kind: CodeErrorKind::UnknownToken {
    ///         token: "???".to_string(),
    ///     },
    /// };
    /// error_details.eprint();
    /// ```
    pub fn eprint(&self) {
        let cache = (
            self.source_name.clone(),
            Source::from(self.source_content.to_string()),
        );

        let report = Report::build(
            ReportKind::Error,
            (self.source_name.clone(), self.line_start..=self.line_end),
        )
        .with_message(format!("{:?}", self.kind))
        .with_label(
            Label::new((
                self.source_name.clone(),
                self.error_offset..=self.error_offset,
            ))
            .with_message("error here"),
        )
        .finish()
        .eprint(cache);

        if let Some(e) = report.err() {
            eprint!("Error attempting to build ErrorDetails eprint message {e:?}");
        }
    }

    /// Convert the `ErrorDetails` into a string using ariadne for formatting
    ///
    /// # Errors
    /// This function will return an error if there is an issue converting the
    /// formatted report into a UTF-8 string.
    pub fn print_to_string(&self) -> Result<String, Box<dyn Error>> {
        let cache = (
            self.source_name.clone(),
            Source::from(self.source_content.to_string()),
        );

        let mut buf = Vec::new();

        let _ = Report::build(
            ReportKind::Error,
            (self.source_name.clone(), self.line_start..=self.line_end),
        )
        .with_message(self.kind.to_string())
        .with_label(
            Label::new((
                self.source_name.clone(),
                self.error_offset..=self.error_offset,
            ))
            .with_message("error here"),
        )
        .finish()
        .write(cache, &mut buf);

        let text = String::from_utf8(buf.clone())?;

        Ok(text)
    }
}

/// Errors related to source file parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum SourceFileErrorKind {
    /// Indicates that the source file is malformed.
    #[error("Unable to parse source file: {message}")]
    MalformedSource {
        /// The error message describing the issue.
        message: String,
    },
}

/// Errors related to code parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum CodeErrorKind {
    /// Indicates that a variable name exceeds the maximum allowed length in VB6.
    #[error("Variable names in VB6 have a maximum length of 255 characters.")]
    VariableNameTooLong,

    /// Indicates that an unknown token was encountered during parsing.
    #[error("Unknown token '{token}' found.")]
    UnknownToken {
        /// The unknown token that was encountered.
        token: String,
    },

    /// Indicates that an unexpected end of the code stream was encountered.
    #[error("Unexpected end of code stream.")]
    UnexpectedEndOfStream,
}

/// Errors related to class parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ClassErrorKind<'a> {
    /// Indicates that the 'VERSION' keyword is missing from the class file header.
    #[error("The 'VERSION' keyword is missing from the class file header.")]
    VersionKeywordMissing,

    /// Indicates that the 'BEGIN' keyword is missing from the class file header.
    #[error("The 'BEGIN' keyword is missing from the class file header.")]
    BeginKeywordMissing,

    /// Indicates that the 'Class' keyword is missing from the class file header.
    #[error("The 'Class' keyword is missing from the class file header.")]
    ClassKeywordMissing,

    /// Indicates that there is missing whitespace between the 'VERSION' keyword and the major version number.
    #[error(
        "After the 'VERSION' keyword there should be a space before the major version number."
    )]
    WhitespaceMissingBetweenVersionAndMajorVersionNumber,

    /// Indicates that the 'VERSION' keyword is not fully uppercase.
    #[error("The 'VERSION' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    VersionKeywordNotFullyUppercase {
        /// The text of the 'VERSION' keyword as found in the source.
        version_text: &'a str,
    },

    /// Indicates that the 'CLASS' keyword is not fully uppercase.
    #[error("The 'CLASS' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    ClassKeywordNotFullyUppercase {
        /// The text of the 'CLASS' keyword as found in the source.
        class_text: &'a str,
    },

    /// Indicates that the 'BEGIN' keyword is not fully uppercase.
    #[error("The 'BEGIN' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    BeginKeywordNotFullyUppercase {
        /// The text of the 'BEGIN' keyword as found in the source.
        begin_text: &'a str,
    },

    /// Indicates that the 'END' keyword is not fully uppercase.
    #[error(
        "The 'END' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE."
    )]
    EndKeywordNotFullyUppercase {
        /// The text of the 'END' keyword as found in the source.
        end_text: &'a str,
    },

    /// Indicates that the 'BEGIN' keyword should be on its own line.
    #[error("The 'BEGIN' keyword should stand alone on its own line.")]
    BeginKeywordShouldBeStandAlone,

    /// Indicates that the 'END' keyword should be on its own line.
    #[error("The 'END' keyword should stand alone on its own line.")]
    EndKeywordShouldBeStandAlone,

    /// Indicates that the major version number could not be parsed.
    #[error("Unable to parse the major version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToParseMajorVersionNumber,

    /// Indicates that the major version text could not be converted to a number.
    #[error("Unable to convert the major version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToConvertMajorVersionNumber,

    /// Indicates that the minor version number could not be parsed.
    #[error("Unable to parse the minor version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToParseMinorVersionNumber,

    /// Indicates that the minor version text could not be converted to a number.
    #[error("Unable to convert the minor version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToConvertMinorVersionNumber,

    /// Indicates that the period divider between major and minor version digits is missing.
    #[error("The '.' divider between major and minor version digits is missing.")]
    MissingPeriodDividerBetweenMajorAndMinorVersion,

    /// Indicates that there is missing whitespace between minor version digits and 'CLASS' keyword.
    #[error("Missing whitespace between minor version digits and 'CLASS' keyword. This may not be compliant with Microsoft's VB6 IDE.")]
    MissingWhitespaceAfterMinorVersion,

    /// Indicates that there is incorrect whitespace between minor version digits and 'CLASS' keyword.
    #[error("Between the minor version digits and the 'CLASS' keyword should be a single ASCII space. This may not be compliant with Microsoft's VB6 IDE.")]
    IncorrectWhitespaceAfterMinorVersion,

    /// Indicates that whitespace was used to divide between major and minor version numbers.
    #[error("Whitespace was used to divide between major and minor version information. This may not be compliant with Microsoft's VB6 IDE.")]
    WhitespaceDividerBetweenMajorAndMinorVersionNumbers,

    /// Indicates that there was an error parsing VB6 tokens.
    #[error("There was an error parsing the VB6 tokens.")]
    ClassTokenError {
        /// The underlying code error that occurred.
        code_error: CodeErrorKind,
    },

    /// Indicates that there was an error parsing the CST.
    #[error("CST parsing error: {0}")]
    CSTError(String),
}

impl<'a> From<ErrorDetails<'a, CodeErrorKind>> for ErrorDetails<'a, ClassErrorKind<'a>> {
    fn from(value: ErrorDetails<'a, CodeErrorKind>) -> Self {
        ErrorDetails {
            source_content: value.source_content,
            source_name: value.source_name,
            error_offset: value.error_offset,
            line_start: value.line_start,
            line_end: value.line_end,
            kind: ClassErrorKind::ClassTokenError {
                code_error: value.kind,
            },
        }
    }
}

/// Errors related to module parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ModuleErrorKind {
    /// Indicates that the 'Attribute' keyword is missing from the module file header.
    #[error("The 'Attribute' keyword is missing from the module file header.")]
    AttributeKeywordMissing,

    /// Indicates that there is missing whitespace between the 'Attribute' keyword and the '`VB_Name`' attribute.
    #[error("The 'Attribute' keyword and the 'VB_Name' attribute must be separated by at least one ASCII whitespace character.")]
    MissingWhitespaceInHeader,

    /// Indicates that the '`VB_Name`' attribute is missing from the module file header.
    #[error("The 'VB_Name' attribute is missing from the module file header.")]
    VBNameAttributeMissing,

    /// Indicates that the '`VB_Name`' attribute is missing the equal symbol.
    #[error("The 'VB_Name' attribute is missing the equal symbol from the module file header.")]
    EqualMissing,

    /// Indicates that the '`VB_Name`' attribute value is unquoted.
    #[error("The 'VB_Name' attribute is unquoted.")]
    VBNameAttributeValueUnquoted,

    /// Indicates that there was an error parsing VB6 tokens.
    #[error("There was an error parsing the VB6 tokens.")]
    ModuleTokenError {
        /// The underlying code error that occurred.
        code_error: CodeErrorKind,
    },
}

impl<'a> From<ErrorDetails<'a, CodeErrorKind>> for ErrorDetails<'a, ModuleErrorKind> {
    fn from(value: ErrorDetails<'a, CodeErrorKind>) -> Self {
        ErrorDetails {
            source_content: value.source_content,
            source_name: value.source_name,
            error_offset: value.error_offset,
            line_start: value.line_start,
            line_end: value.line_end,
            kind: ModuleErrorKind::ModuleTokenError {
                code_error: value.kind,
            },
        }
    }
}

/// Errors related to project parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ProjectErrorKind<'a> {
    /// Indicates that a section header was expected but was not terminated properly.
    #[error("A section header was expected but was not terminated with a ']' character.")]
    UnterminatedSectionHeader,

    /// Indicates that a property name was not found in a property line.
    #[error("Project property line invalid. Expected a Property Name followed by an equal sign '=' and a Property Value.")]
    PropertyNameNotFound,

    /// Indicates that a property value was not found in a property line.
    #[error("'Type' property line invalid. Only the values 'Exe', 'OleDll', 'Control', or 'OleExe' are valid.")]
    ProjectTypeUnknown,

    /// Indicates that the 'Designer' line is invalid.
    #[error("'Designer' line is invalid. Expected a designer path after the equal sign '='. Found a newline or the end of the file instead.")]
    DesignerFileNotFound,

    /// Indicates that the 'Reference' line is invalid for a project-based reference.
    #[error("'Reference' line is invalid. The line started with '*\\G' indicating a compiled reference. Expected a closing '}}' after the UUID, but found a newline or the end of the file instead.")]
    ReferenceCompiledUuidMissingMatchingBrace,

    /// Indicates that the 'Reference' line is invalid for a compiled reference with an invalid UUID.
    #[error("'Reference' line is invalid. The line started with '*\\G' indicating a compiled reference but the contents of the '{{' and '}}' was not a valid UUID.")]
    ReferenceCompiledUuidInvalid,

    /// Indicates that the 'Reference' line is invalid for a project-based reference.
    #[error("'Reference' line is invalid. Expected a reference path but found a newline or the end of the file instead.")]
    ReferenceProjectPathNotFound,

    /// Indicates that the 'Reference' line is invalid for a project-based reference with an invalid path.
    #[error("'Reference' line is invalid. Expected a reference path to begin with '*\\A' followed by the path to the reference project file ending with a quote '\"' character. Found '{value}' instead.")]
    ReferenceProjectPathInvalid {
        /// The invalid path value that was found.
        value: &'a str,
    },

    /// Indicates that the 'Reference' line is invalid for a compiled reference missing 'unknown1'.
    #[error("'Reference' line is invalid. Expected a compiled reference 'unknown1' value after the UUID, between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledUnknown1Missing,

    /// Indicates that the 'Reference' line is invalid for a compiled reference missing 'unknown2'.
    #[error("'Reference' line is invalid. Expected a compiled reference 'unknown2' value after the UUID and 'unknown1', between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledUnknown2Missing,

    /// Indicates that the 'Reference' line is invalid for a compiled reference missing 'path'.
    #[error("'Reference' line is invalid. Expected a compiled reference 'path' value after the UUID, 'unknown1', and 'unknown2', between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledPathNotFound,

    /// Indicates that the 'Reference' line is invalid for a compiled reference with an invalid 'path'.
    #[error("'Reference' line is invalid. Expected a compiled reference 'description' value after the UUID, 'unknown1', 'unknown2', and 'path', but found a newline or the end of the file instead.")]
    ReferenceCompiledDescriptionNotFound,

    /// Indicates that the 'Reference' line is invalid for a compiled reference with an invalid 'description'.
    #[error("'Reference' line is invalid. Compiled reference description contains a '#' character, which is not allowed. The description must be a valid ASCII string without any '#' characters.")]
    ReferenceCompiledDescriptionInvalid,

    /// Indicates that the 'Object' line is invalid for a project-based object.
    #[error("'Object' line is invalid. Project based objects lines must be quoted strings and begin with '*\\A' followed by the path to the object project file ending with a quote '\"' character. Found a newline or the end of the file instead.")]
    ObjectProjectPathNotFound,

    /// Indicates that the 'Object' line is invalid for a compiled object missing opening brace.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{'. Found a newline or the end of the file instead.")]
    ObjectCompiledMissingOpeningBrace,

    /// Indicates that the 'Object' line is invalid for a compiled object missing matching closing brace.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID. Expected a closing '}}' after the UUID, but found a newline or the end of the file instead.")]
    ObjectCompiledUuidMissingMatchingBrace,

    /// Indicates that the 'Object' line is invalid for a compiled object with an invalid UUID.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID, and end with '}}'. The UUID was not valid. Expected a valid UUID in the format 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' containing only ASCII characters.")]
    ObjectCompiledUuidInvalid,

    /// Indicates that the 'Object' line is invalid for a compiled object missing '#' character.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by a '#' character, and then a version number. Expected a '#' character, but found a newline or the end of the file instead.")]
    ObjectCompiledVersionMissing,

    /// Indicates that the 'Object' line is invalid for a compiled object with an invalid version number.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by a '#' character, and then a version number. The version number was not valid. Expected a valid version number in the format 'x.x'. The version number must contain only '.' or the characters \"0\"..\"9\". Invalid character found instead.")]
    ObjectCompiledVersionInvalid,

    /// Indicates that the 'Object' line is invalid for a compiled object missing 'unknown1'.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by '#', a version number, followed by another '#', then an 'unknown1' value followed by \"; \". Expected \"; \", but found a newline or the end of the file instead.")]
    ObjectCompiledUnknown1Missing,

    /// Indicates that the 'Object' line is invalid for a compiled object missing the file name.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by '#', a version number, followed by another '#', then an 'unknown1' value followed by \"; \", and ending with the object's file name. Expected the object's file name, but found a newline or the end of the file instead.")]
    ObjectCompiledFileNameNotFound,

    /// Indicates that the 'Module' line is invalid.
    #[error("'Module' line is invalid. Expected a module name followed by a \"; \". Found a newline or the end of the file instead.")]
    ModuleNameNotFound,

    /// Indicates that the 'Module' line is invalid.
    #[error("'Module' line is invalid. Expected a module name followed by a \"; \", followed by the module file name. Found a newline or the end of the file instead.")]
    ModuleFileNameNotFound,

    /// Indicates that the 'Class' line is invalid.
    #[error("'Class' line is invalid. Expected a class name followed by a \"; \". Found a newline or the end of the file instead.")]
    ClassNameNotFound,

    /// Indicates that the 'Class' line is invalid.
    #[error("'Class' line is invalid. Expected a class name followed by a \"; \", followed by the class file name. Found a newline or the end of the file instead.")]
    ClassFileNameNotFound,

    /// Indicates that a parameter line is invalid because the value is missing.
    #[error("'{parameter_line_name}' line is invalid. Expected a '{parameter_line_name}' path after the equal sign '='. Found a newline or the end of the file instead.")]
    PathValueNotFound {
        /// The name of the parameter line that is invalid.
        parameter_line_name: &'a str,
    },

    /// Indicates that a parameter line is invalid because the value is missing.
    #[error("'{parameter_line_name}' line is invalid. Expected a quoted '{parameter_line_name}' value after the equal sign '='. Found a newline or the end of the file instead.")]
    ParameterValueNotFound {
        /// The name of the parameter line that is invalid.
        parameter_line_name: &'a str,
    },

    /// Indicates that a parameter line is invalid because the value is missing a closing quote.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing an opening quote. Expected a closing quote for the '{parameter_line_name}' value.")]
    ParameterValueMissingOpeningQuote {
        /// The name of the parameter line that is invalid.
        parameter_line_name: &'a str,
    },

    /// Indicates that a parameter line is invalid because the value is missing a matching quote.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing a matching quote. Expected a closing quote for the '{parameter_line_name}' value.")]
    ParameterValueMissingMatchingQuote {
        /// The name of the parameter line that is invalid.
        parameter_line_name: &'a str,
    },

    /// Indicates that a parameter line is invalid because the value is missing both opening and closing quotes.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing both opening and closing quotes. Expected a quoted '{parameter_line_name}' value after the equal sign '='.")]
    ParameterValueMissingQuotes {
        /// The name of the parameter line that is invalid.
        parameter_line_name: &'a str,
    },

    /// Indicates that a parameter line is invalid because the value is not valid.
    #[error("'{parameter_line_name}' line is invalid. '{invalid_value}' is not a valid value for '{parameter_line_name}'. Only {valid_value_message} are valid values for '{parameter_line_name}'.")]
    ParameterValueInvalid {
        /// The name of the parameter line that is invalid.
        parameter_line_name: &'a str,
        /// The invalid value that was found.
        invalid_value: &'a str,
        /// A message describing the valid values for the parameter line.
        valid_value_message: String,
    },

    /// Indicates that the '`DllBaseAddress`' line is invalid.
    #[error("'DllBaseAddress' line is invalid. Expected a hex address after the equal sign '='. Found a newline or the end of the file instead.")]
    DllBaseAddressNotFound,

    /// Indicates that the '`DllBaseAddress`' line is invalid.
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address beginning with '&h' after the equal sign '='.")]
    DllBaseAddressMissingHexPrefix,

    /// Indicates that the '`DllBaseAddress`' line is invalid.
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address after the equal sign '=' beginning with '&h'. Unable to parse hex value '{hex_value}'.")]
    DllBaseAddressUnparsable {
        /// The hex value that could not be parsed.
        hex_value: &'a str,
    },

    /// Indicates that the '`DllBaseAddress`' line is invalid.
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address after the equal sign '=' beginning with '&h'. Unable to parse empty hex value.")]
    DllBaseAddressUnparsableEmpty,

    /// Indicates that a parameter line is unknown.
    #[error("'{parameter_line_name}' line is unknown.")]
    ParameterLineUnknown {
        /// The name of the unknown parameter line.
        parameter_line_name: &'a str,
    },
}

/// Errors related to property parsing.
#[derive(thiserror::Error, PartialEq, Eq, Debug, Clone)]
pub enum PropertyError {
    /// Indicates that the Appearance property has an invalid value.
    #[error("Appearance can only be a 0 (Flat) or a 1 (ThreeD)")]
    AppearanceInvalid,

    /// Indicates that the `BorderStyle` property has an invalid value.
    #[error("BorderStyle can only be a 0 (None) or 1 (FixedSingle)")]
    BorderStyleInvalid,

    /// Indicates that the `ClipControls` property has an invalid value.
    #[error("ClipControls can only be a 0 (false) or a 1 (true)")]
    ClipControlsInvalid,

    /// Indicates that the `DragMode` property has an invalid value.
    #[error("DragMode can only be 0 (Manual) or 1 (Automatic)")]
    DragModeInvalid,

    /// Indicates that the Enabled property has an invalid value.
    #[error("Enabled can only be 0 (false) or a 1 (true)")]
    EnabledInvalid,

    /// Indicates that the `MousePointer` property has an invalid value.
    #[error("MousePointer can only be 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom)")]
    MousePointerInvalid,

    /// Indicates that the `OLEDropMode` property has an invalid value.
    #[error("OLEDropMode can only be 0 (None), or 1 (Manual)")]
    OLEDropModeInvalid,

    /// Indicates that the `RightToLeft` property has an invalid value.
    #[error("RightToLeft can only be 0 (false) or a 1 (true)")]
    RightToLeftInvalid,

    /// Indicates that the Visible property has an invalid value.
    #[error("Visible can only be 0 (false) or a 1 (true)")]
    VisibleInvalid,

    /// Indicates that the property is unknown in the header file.
    #[error("Unknown property in header file")]
    UnknownProperty,

    /// Indicates that the property value is invalid for any value except 0 or -1.
    #[error("Invalid property value. Only 0 or -1 are valid for this property")]
    InvalidPropertyValueZeroNegOne,

    /// Indicates that the property name could not be parsed.
    #[error("Unable to parse the property name")]
    NameUnparsable,

    /// Indicates that the resource file name could not be parsed.
    #[error("Unable to parse the resource file name")]
    ResourceFileNameUnparsable,

    /// Indicates that the offset into the resource file could not be parsed.
    #[error("Unable to parse the offset into the resource file for property")]
    OffsetUnparsable,

    /// Indicates that the property value is invalid for any value except True or False.
    #[error("Invalid property value. Only True or False are valid for this property")]
    InvalidPropertyValueTrueFalse,
}

/// Errors related to form parsing.
#[derive(thiserror::Error, Debug)]
pub enum FormErrorKind {
    /// Indicates that the 'VERSION' keyword is missing from the form file header.
    #[error("The 'VERSION' keyword is missing from the form file header.")]
    VersionKeywordMissing,

    /// Indicates that the 'Begin' keyword is missing from the form file header.
    #[error("The 'Begin' keyword is missing from the form file header.")]
    BeginKeywordMissing,

    /// Indicates that the 'Form' keyword is missing from the form file header.
    #[error("The Form is missing from the form file header.")]
    FormMissing,

    /// Indicates that there was an error parsing VB6 tokens.
    #[error("There was an error parsing the VB6 tokens.")]
    TokenError {
        /// The underlying code error that occurred.
        code_error: CodeErrorKind,
    },

    /// Indicates that there was an error parsing the Property component of the CST.
    #[error("Property parsing error")]
    Property(#[from] PropertyError),

    /// Indicates that there was an error parsing the resource file.
    #[error("Resource file parsing error")]
    ResourceFile(#[from] std::io::Error),

    /// Indicates that there was an error reading the source file.
    #[error("Error reading the source file")]
    SourceFileError(std::io::Error),

    /// Indicates that the form file likely uses a non-English character set.
    #[error("The file contains more than a significant number of non-ASCII characters. This file was likely saved in a non-English character set. The vb6parse crate currently does not support non-english vb6 files.")]
    LikelyNonEnglishCharacterSet,

    /// Indicates that there was an error parsing reference extra section.
    #[error("The reference line has too many elements")]
    ReferenceExtraSections,

    /// Indicates that the reference line is missing required sections.
    #[error("The reference line has too few elements")]
    ReferenceMissingSections,

    /// Indicates that first line of the project file was not a project 'Type' entry.
    #[error("The first line of a VB6 project file must be a project 'Type' entry.")]
    FirstLineNotProject,

    /// Indicates that the line type is unknown.
    #[error("Line type is unknown.")]
    LineTypeUnknown,

    /// Indicates that the project type is unknown.
    #[error("Project type is not Exe, OleDll, Control, or OleExe")]
    ProjectTypeUnknown,

    /// Indicates that the project lacks a version number.
    #[error("Project lacks a version number.")]
    NoVersion,

    /// Indicates their was a parse error while processing an Object line.
    #[error("Project parse error while processing an Object line.")]
    NoObjects,

    /// Indicates there was a parse error while processing a Form line.
    #[error("Form parse error. No Form found in form file.")]
    NoForm,

    /// Indicates there was a parse error while processing Form attributes.
    #[error("Parse error while processing Form attributes.")]
    AttributeParseError,

    /// Indicates there was a parse error while attempting to parse Form Designer tokens.
    #[error("Parse error while attempting to parse Form tokens.")]
    TokenParseError,

    /// Indicates there was a parse error while attempting to find the BEGIN element.
    #[error("Project parse error, failure to find BEGIN element.")]
    NoBegin,

    /// Indicates an error trying to find a line ending.
    #[error("Project line entry is not ended with a recognized line ending.")]
    NoLineEnding,

    /// Indicates an error trying to parse a Uuid.
    #[error("Unable to parse the Uuid")]
    UnableToParseUuid,

    /// Indicates an error trying to find a semicolon split within a form line.
    #[error("Unable to find a semicolon ';' in this line.")]
    NoSemicolonSplit,

    /// Indicates an error trying to find an equal sign split within a line.
    #[error("Unable to find an equal '=' in this line.")]
    NoEqualSplit,

    /// Indicates an error trying to parse the resource file name.
    #[error("While trying to parse the offset into the resource file, no colon ':' was found.")]
    NoColonForOffsetSplit,

    /// Indicates an error trying to find a key value divider within a line.
    #[error("No key value divider found in the line.")]
    NoKeyValueDividerFound,

    /// Indicates an unknown parser error.
    #[error("Unknown parser error")]
    Unparsable,

    /// Indicates that the major version could not be parsed.
    #[error("Major version is not a number.")]
    MajorVersionUnparsable,

    /// Indicates that the `DllBaseAddress` could not be parsed.
    #[error("Unable to parse hex address from DllBaseAddress key")]
    DllBaseAddressUnparsable,

    /// Indicates that the Startup object could not be parsed.
    #[error("The Startup object is not a valid parameter. Must be a quoted startup method/object, \"(None)\", !(None)!, \"\", or \"!!\"")]
    StartupUnparsable,

    /// Indicates that the Name parameter could not be parsed.
    #[error("The Name parameter is invalid. Must be a quoted name, \"(None)\", !(None)!, \"\", or \"!!\"")]
    NameUnparsable,

    /// Indicates that the `CommandLine` parameter could not be parsed.
    #[error("The CommandLine parameter is invalid. Must be a quoted command line, \"(None)\", !(None)!, \"\", or \"!!\"")]
    CommandLineUnparsable,

    /// Indicates that the `HelpContextId` parameter could not be parsed.
    #[error("The HelpContextId parameter is not a valid parameter line. Must be a quoted help context id, \"(None)\", !(None)!, \"\", or \"!!\"")]
    HelpContextIdUnparsable,

    /// Indicates that the Minor version could not be parsed.
    #[error("Minor version is not a number.")]
    MinorVersionUnparsable,

    /// Indicates that the Revision version could not be parsed.
    #[error("Revision version is not a number.")]
    RevisionVersionUnparsable,

    /// Indicates that the `ThreadingModel` value could not be parsed.
    #[error("Unable to parse the value after ThreadingModel key")]
    ThreadingModelUnparsable,

    /// Indicates that the `ThreadingModel` value is invalid.
    #[error("ThreadingModel can only be 0 (Apartment Threaded), or 1 (Single Threaded)")]
    ThreadingModelInvalid,

    /// Indicates that no property name was found after the `BeginProperty` keyword.
    #[error("No property name found after BeginProperty keyword.")]
    NoPropertyName,

    /// Indicates that the `RelatedDoc` property line could not be parsed.
    #[error("Unable to parse the RelatedDoc property line.")]
    RelatedDocLineUnparsable,

    /// Indicates that the `AutoIncrement` property line could not be parsed.
    #[error("AutoIncrement can only be a 0 (false) or a -1 (true)")]
    AutoIncrementUnparsable,

    /// Indicates that the `CompatibilityMode` property line could not be parsed.
    #[error("CompatibilityMode can only be a 0 (CompatibilityMode::NoCompatibility), 1 (CompatibilityMode::Project), or 2 (CompatibilityMode::CompatibleExe)")]
    CompatibilityModeUnparsable,

    /// Indicates that the `NoControlUpgrade` property line could not be parsed.
    #[error("NoControlUpgrade can only be a 0 (UpgradeControls::Upgrade) or a 1 (UpgradeControls::NoUpgrade)")]
    NoControlUpgradeUnparsable,

    /// Indicates that the `ServerSupportFiles` property line could not be parsed.
    #[error("ServerSupportFiles can only be a 0 (false) or a -1 (true)")]
    ServerSupportFilesUnparsable,

    /// Indicates that the Comment line could not be parsed.
    #[error("Comment line was unparsable")]
    CommentUnparsable,

    /// Indicates that the `PropertyPage` line could not be parsed.
    #[error("PropertyPage line was unparsable")]
    PropertyPageUnparsable,

    /// Indicates that the `CompilationType` property line could not be parsed.
    #[error("CompilationType can only be a 0 (false) or a -1 (true)")]
    CompilationTypeUnparsable,

    /// Indicates that the `OptimizationType` property line could not be parsed.
    #[error("OptimizationType can only be a 0 (FastCode) or 1 (SmallCode), or 2 (NoOptimization)")]
    OptimizationTypeUnparsable,

    /// Indicates that the FavorPentiumPro(tm) property line could not be parsed.
    #[error("FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)")]
    FavorPentiumProUnparsable,

    /// Indicates that the `DesignerLine` property line could not be parsed.
    #[error("Designer line is unparsable")]
    DesignerLineUnparsable,

    /// Indicates that the Form line could not be parsed.
    #[error("Form line is unparsable")]
    FormLineUnparsable,

    /// Indicates that the `UserControl` line could not be parsed.
    #[error("UserControl line is unparsable")]
    UserControlLineUnparsable,

    /// Indicates that the `UserDocument` line could not be parsed.
    #[error("UserDocument line is unparsable")]
    UserDocumentLineUnparsable,

    /// Indicates that the Period divider in the version number is missing.
    #[error("Period expected in version number")]
    PeriodExpectedInVersionNumber,

    /// Indicates that the `CodeViewDebugInfo` property line could not be parsed.
    #[error("CodeViewDebugInfo can only be a 0 (false) or a -1 (true)")]
    CodeViewDebugInfoUnparsable,

    /// Indicates that the `NoAliasing` property line could not be parsed.
    #[error("NoAliasing can only be a 0 (false) or a -1 (true)")]
    NoAliasingUnparsable,

    /// Indicates that the `RemoveUnusedControlInfo` property line could not be parsed.
    #[error("RemoveUnusedControlInfo can only be 0 (UnusedControlInfo::Retain) or -1 (UnusedControlInfo::Remove)")]
    UnusedControlInfoUnparsable,

    /// Indicates that the `BoundsCheck` property line could not be parsed.
    #[error("BoundsCheck can only be a 0 (false) or a -1 (true)")]
    BoundsCheckUnparsable,

    /// Indicates that the `OverflowCheck` property line could not be parsed.
    #[error("OverflowCheck can only be a 0 (false) or a -1 (true)")]
    OverflowCheckUnparsable,

    /// Indicates that the `FlPointCheck` property line could not be parsed.
    #[error("FlPointCheck can only be a 0 (false) or a -1 (true)")]
    FlPointCheckUnparsable,

    /// Indicates that the `FDIVCheck` property line could not be parsed.
    #[error("FDIVCheck can only be a 0 (PentiumFDivBugCheck::CheckPentiumFDivBug) or a -1 (PentiumFDivBugCheck::NoPentiumFDivBugCheck)")]
    FDIVCheckUnparsable,

    /// Indicates that the `UnroundedFP` property line could not be parsed.
    #[error("UnroundedFP can only be a 0 (UnroundedFloatingPoint::DoNotAllow) or a -1 (UnroundedFloatingPoint::Allow)")]
    UnroundedFPUnparsable,

    /// Indicates that the `StartMode` property line could not be parsed.
    #[error("StartMode can only be a 0 (StartMode::StandAlone) or a 1 (StartMode::Automation)")]
    StartModeUnparsable,

    /// Indicates that the Unattended property line could not be parsed.
    #[error("Unattended can only be a 0 (Unattended::False) or a -1 (Unattended::True)")]
    UnattendedUnparsable,

    /// Indicates that the Retained property line could not be parsed.
    #[error(
        "Retained can only be a 0 (Retained::UnloadOnExit) or a 1 (Retained::RetainedInMemory)"
    )]
    RetainedUnparsable,

    /// Indicates that the `ShortCut` property line could not be parsed.
    #[error("Unable to parse the ShortCut property.")]
    ShortCutUnparsable,

    /// Indicates that the `DebugStartup` property line could not be parsed.
    #[error("DebugStartup can only be a 0 (false) or a -1 (true)")]
    DebugStartupOptionUnparsable,

    /// Indicates that the `UseExistingBrowser` property line could not be parsed.
    #[error("UseExistingBrowser can only be a 0 (UseExistingBrowser::DoNotUse) or a -1 (UseExistingBrowser::Use)")]
    UseExistingBrowserUnparsable,

    /// Indicates that the `AutoRefresh` property line could not be parsed.
    #[error("AutoRefresh can only be a 0 (false) or a -1 (true)")]
    AutoRefreshUnparsable,

    /// Indicates that the `ConnectionType` property line could not be parsed.
    #[error("Data control Connection type is not valid.")]
    ConnectionTypeUnparsable,

    /// Indicates that the `ThreadPerObject` property line could not be parsed.
    #[error("Thread Per Object is not a number.")]
    ThreadPerObjectUnparsable,

    /// Indicates that an attribute in the class header file is unknown.
    #[error("Unknown attribute in class header file. Must be one of: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_Ext_KEY")]
    UnknownAttribute,

    /// Indicates that there was an error parsing the header of the VB6 file.
    #[error("Error parsing header")]
    Header,

    /// Indicates that there is no name in the attribute section of the VB6 file.
    #[error("No name in the attribute section of the VB6 file")]
    MissingNameAttribute,

    /// Indicates that a keyword was not found where expected.
    #[error("Keyword not found")]
    KeywordNotFound,

    /// Indicates that there was an error parsing a true/false value from a header.
    #[error("Error parsing true/false from header. Must be a 0 (false), -1 (true), or 1 (true)")]
    TrueFalseOneZeroNegOneUnparsable,

    /// Indicates that there was an error parsing the contents of the VB6 file.
    #[error("Error parsing the VB6 file contents")]
    FileContent,

    /// Indicates that the Max Threads property line could not be parsed.
    #[error("Max Threads is not a number.")]
    MaxThreadsUnparsable,

    /// Indicates that no `EndProperty` was found after a `BeginProperty`.
    #[error("No EndProperty found after BeginProperty")]
    NoEndProperty,

    /// Indicates that there was no line ending after the `EndProperty`.
    #[error("No line ending after EndProperty")]
    NoLineEndingAfterEndProperty,

    /// Indicates that there was no namespace after the Begin keyword.
    #[error("Expected namespace after Begin keyword")]
    NoNamespaceAfterBegin,

    /// Indicates that there was no dot after the namespace.
    #[error("No dot found after namespace")]
    NoDotAfterNamespace,

    /// Indicates that there was no User Control name after the namespace and dot.
    #[error("No User Control name found after namespace and '.'")]
    NoUserControlNameAfterDot,

    /// Indicates that there was no ASCII space after the User Control kind.
    #[error("No space after Control kind")]
    NoSpaceAfterControlKind,

    /// Indicates that there was no Control name after the Control kind.
    #[error("No control name found after Control kind")]
    NoControlNameAfterControlKind,

    /// Indicates that there was no line ending after the Control name.
    #[error("No line ending after Control name")]
    NoLineEndingAfterControlName,

    /// Indicates that an unknown token was encountered.
    #[error("Unknown token")]
    UnknownToken,

    /// Indicates that the title text could not be parsed.
    #[error("Title text was unparsable")]
    TitleUnparsable,

    /// Indicates that a hex color value could not be parsed.
    #[error("Unable to parse hex color value")]
    HexColorParseError,

    /// Indicates that an unknown control kind was encountered in the control list.
    #[error("Unknown control in control list")]
    UnknownControlKind,

    /// Indicates that the property name is not a valid ASCII string.
    #[error("Property name is not a valid ASCII string")]
    PropertyNameAsciiConversionError,

    /// Indicates that the property value is unterminated.
    #[error("String is unterminated")]
    UnterminatedString,

    /// Indicates that there was an error parsing a VB6 string.
    #[error("Unable to parse VB6 string.")]
    StringParseError,

    /// Indicates that the property value is not a valid ASCII string.
    #[error("Property value is not a valid ASCII string")]
    PropertyValueAsciiConversionError,

    /// Indicates that the key value pair format is incorrect.
    #[error("Key value pair format is incorrect")]
    KeyValueParseError,

    /// Indicates that the namespace is not a valid ASCII string.
    #[error("Namespace is not a valid ASCII string")]
    NamespaceAsciiConversionError,

    /// Indicates that the user control name is not a valid ASCII string.
    #[error("Control kind is not a valid ASCII string")]
    ControlKindAsciiConversionError,

    /// Indicates that the qualified control name is not a valid ASCII string.
    #[error("Qualified control name is not a valid ASCII string")]
    QualifiedControlNameAsciiConversionError,

    /// Indicates that the variable name is too long for VB6.
    #[error("Variable names must be less than 255 characters in VB6.")]
    VariableNameTooLong,

    /// Indicates an internal parser error that should be reported to the developers.
    #[error("Internal Parser Error - please report this issue to the developers.")]
    InternalParseError,
}
