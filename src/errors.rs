//! This module contains the error types used in the VB6 parser.
//! It defines the `VB6Error` type, which is used to represent
//! errors that occur during parsing. The `VB6Error` type contains
//! information about the error, including the file name, source code,
//! source offset, column, line number, and the kind of error.
use core::convert::From;

use std::borrow::Cow;
use std::error::Error;
use std::fmt::Debug;

use ariadne::{Label, Report, ReportKind, Source};

/// Represents detailed information about an error that occurred during parsing.
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
pub struct ErrorDetails<'a, T> {
    pub source_name: String,
    pub source_content: Cow<'a, str>,
    pub error_offset: usize,
    pub line_start: usize,
    pub line_end: usize,
    pub kind: T,
}

impl<T> ErrorDetails<'_, T>
where
    T: ToString,
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
        .with_message(self.kind.to_string())
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

/// Represents errors related to source file parsing.
/// This enum defines various kinds of source file errors
/// that can occur during the parsing process.
/// Each variant includes a descriptive error message.
#[derive(thiserror::Error, Debug, Clone)]
pub enum SourceFileErrorKind {
    #[error("Unable to parse source file: {message}")]
    MalformedSource { message: String },
}

/// Represents errors related to code parsing.
/// This enum defines various kinds of code errors
/// that can occur during the parsing process.
/// Each variant includes a descriptive error message.
#[derive(thiserror::Error, Debug, Clone)]
pub enum CodeErrorKind {
    #[error("Variable names in VB6 have a maximum length of 255 characters.")]
    VariableNameTooLong,

    #[error("Unknown token '{token}' found.")]
    UnknownToken { token: String },

    #[error("Unexpected end of code stream.")]
    UnexpectedEndOfStream,
}

#[derive(thiserror::Error, Debug, Clone)]
pub enum ClassErrorKind<'a> {
    #[error("The 'VERSION' keyword is missing from the class file header.")]
    VersionKeywordMissing,

    #[error("The 'BEGIN' keyword is missing from the class file header.")]
    BeginKeywordMissing,

    #[error("The 'Class' keyword is missing from the class file header.")]
    ClassKeywordMissing,

    #[error("The 'Form' keyword is missing from the class file header.")]
    FormKeywordMissing,

    #[error(
        "After the 'VERSION' keyword there should be a space before the major version number."
    )]
    WhitespaceMissingBetweenVersionAndMajorVersionNumber,

    #[error("The 'VERSION' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    VersionKeywordNotFullyUppercase { version_text: &'a str },

    #[error("The 'CLASS' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    ClassKeywordNotFullyUppercase { class_text: &'a str },

    #[error("The 'Form' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    FormKeywordNotFullyUppercase { form_text: &'a str },

    #[error("The 'BEGIN' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    BeginKeywordNotFullyUppercase { begin_text: &'a str },

    #[error(
        "The 'END' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE."
    )]
    EndKeywordNotFullyUppercase { begin_text: &'a str },

    #[error("The 'BEGIN' keyword should stand alone on its own line.")]
    BeginKeywordShouldBeStandAlone,

    #[error("The 'END' keyword should stand alone on its own line.")]
    EndKeywordShouldBeStandAlone,

    #[error("Unable to parse the major version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToParseMajorVersionNumber,

    #[error("Unable to convert the major version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToConvertMajorVersionNumber,

    #[error("Unable to parse the minor version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToParseMinorVersionNumber,

    #[error("Unable to convert the minor version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToConvertMinorVersionNumber,

    #[error("The '.' divider between major and minor version digits is missing.")]
    MissingPeriodDividerBetweenMajorAndMinorVersion,

    #[error("Missing whitespace between minor version digits and 'CLASS' keyword. This may not be compliant with Microsoft's VB6 IDE.")]
    MissingWhitespaceAfterMinorVersion,

    #[error("Between the minor version digits and the 'CLASS' keyword should be a single ASCII space. This may not be compliant with Microsoft's VB6 IDE.")]
    IncorrectWhitespaceAfterMinorVersion,

    #[error("Whitespace was used to divide between major and minor version information. This may not be compliant with Microsoft's VB6 IDE.")]
    WhitespaceDividerBetweenMajorAndMinorVersionNumbers,

    #[error("There was an error parsing the VB6 tokens.")]
    ClassTokenError { code_error: CodeErrorKind },

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

/// Represents errors related to module parsing.
/// This enum defines various kinds of module errors
/// that can occur during the parsing process.
/// Each variant includes a descriptive error message.
#[derive(thiserror::Error, Debug, Clone)]
pub enum ModuleErrorKind {
    #[error("The 'Attribute' keyword is missing from the module file header.")]
    AttributeKeywordMissing,

    #[error("The 'Attribute' keyword and the 'VB_Name' attribute must be separated by at least one ASCII whitespace character.")]
    MissingWhitespaceInHeader,

    #[error("The 'VB_Name' attribute is missing from the module file header.")]
    VBNameAttributeMissing,

    #[error("The 'VB_Name' attribute is missing the equal symbol from the module file header.")]
    EqualMissing,

    #[error("The 'VB_Name' attribute is unquoted.")]
    VBNameAttributeValueUnquoted,

    #[error("There was an error parsing the VB6 tokens.")]
    ModuleTokenError { code_error: CodeErrorKind },
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

/// Represents errors related to project parsing.
/// This enum defines various kinds of project errors
/// that can occur during the parsing process.
/// Each variant includes a descriptive error message.
#[derive(thiserror::Error, Debug, Clone)]
pub enum ProjectErrorKind<'a> {
    #[error("A section header was expected but was not terminated with a ']' character.")]
    UnterminatedSectionHeader,

    #[error("Project property line invalid. Expected a Property Name followed by an equal sign '=' and a Property Value.")]
    PropertyNameNotFound,

    #[error("'Type' property line invalid. Only the values 'Exe', 'OleDll', 'Control', or 'OleExe' are valid.")]
    ProjectTypeUnknown,

    #[error("'Designer' line is invalid. Expected a designer path after the equal sign '='. Found a newline or the end of the file instead.")]
    DesignerFileNotFound,

    #[error("'Reference' line is invalid. The line started with '*\\G' indicating a compiled reference. Expected a closing '}}' after the UUID, but found a newline or the end of the file instead.")]
    ReferenceCompiledUuidMissingMatchingBrace,

    #[error("'Reference' line is invalid. The line started with '*\\G' indicating a compiled reference but the contents of the '{{' and '}}' was not a valid UUID.")]
    ReferenceCompiledUuidInvalid,

    #[error("'Reference' line is invalid. Expected a reference path but found a newline or the end of the file instead.")]
    ReferenceProjectPathNotFound,

    #[error("'Reference' line is invalid. Expected a reference path to begin with '*\\A' followed by the path to the reference project file ending with a quote '\"' character. Found '{value}' instead.")]
    ReferenceProjectPathInvalid { value: &'a str },

    #[error("'Reference' line is invalid. Expected a compiled reference 'unknown1' value after the UUID, between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledUnknown1Missing,

    #[error("'Reference' line is invalid. Expected a compiled reference 'unknown2' value after the UUID and 'unknown1', between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledUnknown2Missing,

    #[error("'Reference' line is invalid. Expected a compiled reference 'path' value after the UUID, 'unknown1', and 'unknown2', between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledPathNotFound,

    #[error("'Reference' line is invalid. Expected a compiled reference 'description' value after the UUID, 'unknown1', 'unknown2', and 'path', but found a newline or the end of the file instead.")]
    ReferenceCompiledDescriptionNotFound,

    #[error("'Reference' line is invalid. Compiled reference description contains a '#' character, which is not allowed. The description must be a valid ASCII string without any '#' characters.")]
    ReferenceCompiledDescriptionInvalid,

    #[error("'Object' line is invalid. Project based objects lines must be quoted strings and begin with '*\\A' followed by the path to the object project file ending with a quote '\"' character. Found a newline or the end of the file instead.")]
    ObjectProjectPathNotFound,

    #[error("'Object' line is invalid. Compiled object lines must begin with '{{'. Found a newline or the end of the file instead.")]
    ObjectCompiledMissingOpeningBrace,

    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID. Expected a closing '}}' after the UUID, but found a newline or the end of the file instead.")]
    ObjectCompiledUuidMissingMatchingBrace,

    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID, and end with '}}'. The UUID was not valid. Expected a valid UUID in the format 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' containing only ASCII characters.")]
    ObjectCompiledUuidInvalid,

    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by a '#' character, and then a version number. Expected a '#' character, but found a newline or the end of the file instead.")]
    ObjectCompiledVersionMissing,

    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by a '#' character, and then a version number. The version number was not valid. Expected a valid version number in the format 'x.x'. The version number must contain only '.' or the characters \"0\"..\"9\". Invalid character found instead.")]
    ObjectCompiledVersionInvalid,

    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by '#', a version number, followed by another '#', then an 'unknown1' value followed by \"; \". Expected \"; \", but found a newline or the end of the file instead.")]
    ObjectCompiledUnknown1Missing,

    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by '#', a version number, followed by another '#', then an 'unknown1' value followed by \"; \", and ending with the object's file name. Expected the object's file name, but found a newline or the end of the file instead.")]
    ObjectCompiledFileNameNotFound,

    #[error("'Module' line is invalid. Expected a module name followed by a \"; \". Found a newline or the end of the file instead.")]
    ModuleNameNotFound,

    #[error("'Module' line is invalid. Expected a module name followed by a \"; \", followed by the module file name. Found a newline or the end of the file instead.")]
    ModuleFileNameNotFound,

    #[error("'Class' line is invalid. Expected a class name followed by a \"; \". Found a newline or the end of the file instead.")]
    ClassNameNotFound,

    #[error("'Class' line is invalid. Expected a class name followed by a \"; \", followed by the class file name. Found a newline or the end of the file instead.")]
    ClassFileNameNotFound,

    #[error("'{parameter_line_name}' line is invalid. Expected a '{parameter_line_name}' path after the equal sign '='. Found a newline or the end of the file instead.")]
    PathValueNotFound { parameter_line_name: &'a str },

    #[error("'{parameter_line_name}' line is invalid. Expected a quoted '{parameter_line_name}' value after the equal sign '='. Found a newline or the end of the file instead.")]
    ParameterValueNotFound { parameter_line_name: &'a str },

    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing an opening quote. Expected a closing quote for the '{parameter_line_name}' value.")]
    ParameterValueMissingOpeningQuote { parameter_line_name: &'a str },

    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing a matching quote. Expected a closing quote for the '{parameter_line_name}' value.")]
    ParameterValueMissingMatchingQuote { parameter_line_name: &'a str },

    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing both opening and closing quotes. Expected a quoted '{parameter_line_name}' value after the equal sign '='.")]
    ParameterValueMissingQuotes { parameter_line_name: &'a str },

    #[error("'{parameter_line_name}' line is invalid. '{invalid_value}' is not a valid value for '{parameter_line_name}'. Only {valid_value_message} are valid values for '{parameter_line_name}'.")]
    ParameterValueInvalid {
        parameter_line_name: &'a str,
        invalid_value: &'a str,
        valid_value_message: String,
    },

    #[error("'DllBaseAddress' line is invalid. Expected a hex address after the equal sign '='. Found a newline or the end of the file instead.")]
    DllBaseAddressNotFound,

    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address beginning with '&h' after the equal sign '='.")]
    DllBaseAddressMissingHexPrefix,

    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address after the equal sign '=' beginning with '&h'. Unable to parse hex value '{hex_value}'.")]
    DllBaseAddressUnparsable { hex_value: &'a str },

    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address after the equal sign '=' beginning with '&h'. Unable to parse empty hex value.")]
    DllBaseAddressUnparsableEmpty,

    #[error("'{parameter_line_name}' line is unknown.")]
    ParameterLineUnknown { parameter_line_name: &'a str },
}

#[derive(thiserror::Error, PartialEq, Eq, Debug)]
pub enum PropertyError {
    #[error("Appearance can only be a 0 (Flat) or a 1 (ThreeD)")]
    AppearanceInvalid,

    #[error("BorderStyle can only be a 0 (None) or 1 (FixedSingle)")]
    BorderStyleInvalid,

    #[error("ClipControls can only be a 0 (false) or a 1 (true)")]
    ClipControlsInvalid,

    #[error("DragMode can only be 0 (Manual) or 1 (Automatic)")]
    DragModeInvalid,

    #[error("Enabled can only be 0 (false) or a 1 (true)")]
    EnabledInvalid,

    #[error("MousePointer can only be 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom)")]
    MousePointerInvalid,

    #[error("OLEDropMode can only be 0 (None), or 1 (Manual)")]
    OLEDropModeInvalid,

    #[error("RightToLeft can only be 0 (false) or a 1 (true)")]
    RightToLeftInvalid,

    #[error("Visible can only be 0 (false) or a 1 (true)")]
    VisibleInvalid,

    #[error("Unknown property in header file")]
    UnknownProperty,

    #[error("Invalid property value. Only 0 or -1 are valid for this property")]
    InvalidPropertyValueZeroNegOne,

    #[error("Unable to parse the property name")]
    NameUnparsable,

    #[error("Unable to parse the resource file name")]
    ResourceFileNameUnparsable,

    #[error("Unable to parse the offset into the resource file for property")]
    OffsetUnparsable,

    #[error("Invalid property value. Only True or False are valid for this property")]
    InvalidPropertyValueTrueFalse,
}

/// Represents errors related to form parsing.
/// This enum defines various kinds of form errors
/// that can occur during the parsing process.
#[derive(thiserror::Error, Debug)]
pub enum FormErrorKind {
    #[error("The 'VERSION' keyword is missing from the form file header.")]
    VersionKeywordMissing,

    #[error("The 'Begin' keyword is missing from the form file header.")]
    BeginKeywordMissing,

    #[error("The Form is missing from the form file header.")]
    FormMissing,

    #[error("There was an error parsing the VB6 tokens.")]
    TokenError { code_error: CodeErrorKind },

    #[error("Property parsing error")]
    Property(#[from] PropertyError),

    #[error("Resource file parsing error")]
    ResourceFile(#[from] std::io::Error),

    #[error("Error reading the source file")]
    SourceFileError(std::io::Error),

    #[error("The file contains more than a significant number of non-ASCII characters. This file was likely saved in a non-English character set. The vb6parse crate currently does not support non-english vb6 files.")]
    LikelyNonEnglishCharacterSet,

    #[error("The reference line has too many elements")]
    ReferenceExtraSections,

    #[error("The reference line has too few elements")]
    ReferenceMissingSections,

    #[error("The first line of a VB6 project file must be a project 'Type' entry.")]
    FirstLineNotProject,

    #[error("Line type is unknown.")]
    LineTypeUnknown,

    #[error("Project type is not Exe, OleDll, Control, or OleExe")]
    ProjectTypeUnknown,

    #[error("Project lacks a version number.")]
    NoVersion,

    #[error("Project parse error while processing an Object line.")]
    NoObjects,

    #[error("Form parse error. No Form found in form file.")]
    NoForm,

    #[error("Parse error while processing Form attributes.")]
    AttributeParseError,

    #[error("Parse error while attempting to parse Form tokens.")]
    TokenParseError,

    #[error("Project parse error, failure to find BEGIN element.")]
    NoBegin,

    #[error("Project line entry is not ended with a recognized line ending.")]
    NoLineEnding,

    #[error("Unable to parse the Uuid")]
    UnableToParseUuid,

    #[error("Unable to find a semicolon ';' in this line.")]
    NoSemicolonSplit,

    #[error("Unable to find an equal '=' in this line.")]
    NoEqualSplit,

    #[error("While trying to parse the offset into the resource file, no colon ':' was found.")]
    NoColonForOffsetSplit,

    #[error("No key value divider found in the line.")]
    NoKeyValueDividerFound,

    #[error("Unknown parser error")]
    Unparsable,

    #[error("Major version is not a number.")]
    MajorVersionUnparsable,

    #[error("Unable to parse hex address from DllBaseAddress key")]
    DllBaseAddressUnparsable,

    #[error("The Startup object is not a valid parameter. Must be a quoted startup method/object, \"(None)\", !(None)!, \"\", or \"!!\"")]
    StartupUnparsable,

    #[error("The Name parameter is invalid. Must be a quoted name, \"(None)\", !(None)!, \"\", or \"!!\"")]
    NameUnparsable,

    #[error("The CommandLine parameter is invalid. Must be a quoted command line, \"(None)\", !(None)!, \"\", or \"!!\"")]
    CommandLineUnparsable,

    #[error("The HelpContextId parameter is not a valid parameter line. Must be a quoted help context id, \"(None)\", !(None)!, \"\", or \"!!\"")]
    HelpContextIdUnparsable,

    #[error("Minor version is not a number.")]
    MinorVersionUnparsable,

    #[error("Revision version is not a number.")]
    RevisionVersionUnparsable,

    #[error("Unable to parse the value after ThreadingModel key")]
    ThreadingModelUnparsable,

    #[error("ThreadingModel can only be 0 (Apartment Threaded), or 1 (Single Threaded)")]
    ThreadingModelInvalid,

    #[error("No property name found after BeginProperty keyword.")]
    NoPropertyName,

    #[error("Unable to parse the RelatedDoc property line.")]
    RelatedDocLineUnparsable,

    #[error("AutoIncrement can only be a 0 (false) or a -1 (true)")]
    AutoIncrementUnparsable,

    #[error("CompatibilityMode can only be a 0 (CompatibilityMode::NoCompatibility), 1 (CompatibilityMode::Project), or 2 (CompatibilityMode::CompatibleExe)")]
    CompatibilityModeUnparsable,

    #[error("NoControlUpgrade can only be a 0 (UpgradeControls::Upgrade) or a 1 (UpgradeControls::NoUpgrade)")]
    NoControlUpgradeUnparsable,

    #[error("ServerSupportFiles can only be a 0 (false) or a -1 (true)")]
    ServerSupportFilesUnparsable,

    #[error("Comment line was unparsable")]
    CommentUnparsable,

    #[error("PropertyPage line was unparsable")]
    PropertyPageUnparsable,

    #[error("CompilationType can only be a 0 (false) or a -1 (true)")]
    CompilationTypeUnparsable,

    #[error("OptimizationType can only be a 0 (FastCode) or 1 (SmallCode), or 2 (NoOptimization)")]
    OptimizationTypeUnparsable,

    #[error("FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)")]
    FavorPentiumProUnparsable,

    #[error("Designer line is unparsable")]
    DesignerLineUnparsable,

    #[error("Form line is unparsable")]
    FormLineUnparsable,

    #[error("UserControl line is unparsable")]
    UserControlLineUnparsable,

    #[error("UserDocument line is unparsable")]
    UserDocumentLineUnparsable,

    #[error("Period expected in version number")]
    PeriodExpectedInVersionNumber,

    #[error("CodeViewDebugInfo can only be a 0 (false) or a -1 (true)")]
    CodeViewDebugInfoUnparsable,

    #[error("NoAliasing can only be a 0 (false) or a -1 (true)")]
    NoAliasingUnparsable,

    #[error("RemoveUnusedControlInfo can only be 0 (UnusedControlInfo::Retain) or -1 (UnusedControlInfo::Remove)")]
    UnusedControlInfoUnparsable,

    #[error("BoundsCheck can only be a 0 (false) or a -1 (true)")]
    BoundsCheckUnparsable,

    #[error("OverflowCheck can only be a 0 (false) or a -1 (true)")]
    OverflowCheckUnparsable,

    #[error("FlPointCheck can only be a 0 (false) or a -1 (true)")]
    FlPointCheckUnparsable,

    #[error("FDIVCheck can only be a 0 (PentiumFDivBugCheck::CheckPentiumFDivBug) or a -1 (PentiumFDivBugCheck::NoPentiumFDivBugCheck)")]
    FDIVCheckUnparsable,

    #[error("UnroundedFP can only be a 0 (UnroundedFloatingPoint::DoNotAllow) or a -1 (UnroundedFloatingPoint::Allow)")]
    UnroundedFPUnparsable,

    #[error("StartMode can only be a 0 (StartMode::StandAlone) or a 1 (StartMode::Automation)")]
    StartModeUnparsable,

    #[error("Unattended can only be a 0 (Unattended::False) or a -1 (Unattended::True)")]
    UnattendedUnparsable,

    #[error(
        "Retained can only be a 0 (Retained::UnloadOnExit) or a 1 (Retained::RetainedInMemory)"
    )]
    RetainedUnparsable,

    #[error("Unable to parse the ShortCut property.")]
    ShortCutUnparsable,

    #[error("DebugStartup can only be a 0 (false) or a -1 (true)")]
    DebugStartupOptionUnparsable,

    #[error("UseExistingBrowser can only be a 0 (UseExistingBrowser::DoNotUse) or a -1 (UseExistingBrowser::Use)")]
    UseExistingBrowserUnparsable,

    #[error("AutoRefresh can only be a 0 (false) or a -1 (true)")]
    AutoRefreshUnparsable,

    #[error("Data control Connection type is not valid.")]
    ConnectionTypeUnparsable,

    #[error("Thread Per Object is not a number.")]
    ThreadPerObjectUnparsable,

    #[error("Unknown attribute in class header file. Must be one of: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_Ext_KEY")]
    UnknownAttribute,

    #[error("Error parsing header")]
    Header,

    #[error("No name in the attribute section of the VB6 file")]
    MissingNameAttribute,

    #[error("Keyword not found")]
    KeywordNotFound,

    #[error("Error parsing true/false from header. Must be a 0 (false), -1 (true), or 1 (true)")]
    TrueFalseOneZeroNegOneUnparsable,

    #[error("Error parsing the VB6 file contents")]
    FileContent,

    #[error("Max Threads is not a number.")]
    MaxThreadsUnparsable,

    #[error("No EndProperty found after BeginProperty")]
    NoEndProperty,

    #[error("No line ending after EndProperty")]
    NoLineEndingAfterEndProperty,

    #[error("Expected namespace after Begin keyword")]
    NoNamespaceAfterBegin,

    #[error("No dot found after namespace")]
    NoDotAfterNamespace,

    #[error("No User Control name found after namespace and '.'")]
    NoUserControlNameAfterDot,

    #[error("No space after Control kind")]
    NoSpaceAfterControlKind,

    #[error("No control name found after Control kind")]
    NoControlNameAfterControlKind,

    #[error("No line ending after Control name")]
    NoLineEndingAfterControlName,

    #[error("Unknown token")]
    UnknownToken,

    #[error("Title text was unparsable")]
    TitleUnparsable,

    #[error("Unable to parse hex color value")]
    HexColorParseError,

    #[error("Unknown control in control list")]
    UnknownControlKind,

    #[error("Property name is not a valid ASCII string")]
    PropertyNameAsciiConversionError,

    #[error("String is unterminated")]
    UnterminatedString,

    #[error("Unable to parse VB6 string.")]
    StringParseError,

    #[error("Property value is not a valid ASCII string")]
    PropertyValueAsciiConversionError,

    #[error("Key value pair format is incorrect")]
    KeyValueParseError,

    #[error("Namespace is not a valid ASCII string")]
    NamespaceAsciiConversionError,

    #[error("Control kind is not a valid ASCII string")]
    ControlKindAsciiConversionError,

    #[error("Qualified control name is not a valid ASCII string")]
    QualifiedControlNameAsciiConversionError,

    #[error("Variable names must be less than 255 characters in VB6.")]
    VariableNameTooLong,

    #[error("Internal Parser Error - please report this issue to the developers.")]
    InternalParseError,
}
