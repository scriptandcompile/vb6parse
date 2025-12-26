//! Errors related to VB6 project file (.vbp) parsing.
//!
//! This module contains error types for issues that occur during:
//! - Project file section parsing
//! - Reference and object line parsing
//! - Module, class, and form references
//! - Compilation settings and parameters

/// Errors related to project file parsing.
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
