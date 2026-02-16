//! Project file (.vbp) parsing errors.

/// Errors that can occur during project file parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ProjectError {
    /// A section header was not terminated properly.
    #[error("A section header was expected but was not terminated with a ']' character.")]
    UnterminatedSectionHeader,

    /// Property name was not found in a property line.
    #[error("Project property line invalid. Expected a Property Name followed by an equal sign '=' and a Property Value.")]
    PropertyNameNotFound,

    /// Project type is unknown.
    #[error("'Type' property line invalid. Only the values 'Exe', 'OleDll', 'Control', or 'OleExe' are valid.")]
    TypeUnknown,

    /// Designer file not found.
    #[error("'Designer' line is invalid. Expected a designer path after the equal sign '='. Found a newline or the end of the file instead.")]
    DesignerFileNotFound,

    /// Reference compiled UUID missing matching brace.
    #[error("'Reference' line is invalid. The line started with '*\\G' indicating a compiled reference. Expected a closing '}}' after the UUID, but found a newline or the end of the file instead.")]
    ReferenceCompiledUuidMissingMatchingBrace,

    /// Reference compiled UUID is invalid.
    #[error("'Reference' line is invalid. The line started with '*\\G' indicating a compiled reference but the contents of the '{{' and '}}' was not a valid UUID.")]
    ReferenceCompiledUuidInvalid,

    /// Reference project path not found.
    #[error("'Reference' line is invalid. Expected a reference path but found a newline or the end of the file instead.")]
    ReferenceProjectPathNotFound,

    /// Reference project path is invalid.
    #[error("'Reference' line is invalid. Expected a reference path to begin with '*\\A' followed by the path to the reference project file ending with a quote '\"' character. Found '{value}' instead.")]
    ReferenceProjectPathInvalid {
        /// The invalid path value that was found.
        value: String,
    },

    /// Reference compiled unknown1 missing.
    #[error("'Reference' line is invalid. Expected a compiled reference 'unknown1' value after the UUID, between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledUnknown1Missing,

    /// Reference compiled unknown2 missing.
    #[error("'Reference' line is invalid. Expected a compiled reference 'unknown2' value after the UUID and 'unknown1', between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledUnknown2Missing,

    /// Reference compiled path not found.
    #[error("'Reference' line is invalid. Expected a compiled reference 'path' value after the UUID, 'unknown1', and 'unknown2', between '#' characters but found a newline or the end of the file instead.")]
    ReferenceCompiledPathNotFound,

    /// Reference compiled description not found.
    #[error("'Reference' line is invalid. Expected a compiled reference 'description' value after the UUID, 'unknown1', 'unknown2', and 'path', but found a newline or the end of the file instead.")]
    ReferenceCompiledDescriptionNotFound,

    /// Reference compiled description is invalid.
    #[error("'Reference' line is invalid. Compiled reference description contains a '#' character, which is not allowed. The description must be a valid ASCII string without any '#' characters.")]
    ReferenceCompiledDescriptionInvalid,

    /// Object project path not found.
    #[error("'Object' line is invalid. Project based objects lines must be quoted strings and begin with '*\\A' followed by the path to the object project file ending with a quote '\"' character. Found a newline or the end of the file instead.")]
    ObjectProjectPathNotFound,

    /// Object compiled missing opening brace.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{'. Found a newline or the end of the file instead.")]
    ObjectCompiledMissingOpeningBrace,

    /// Object compiled UUID missing matching brace.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID. Expected a closing '}}' after the UUID, but found a newline or the end of the file instead.")]
    ObjectCompiledUuidMissingMatchingBrace,

    /// Object compiled UUID is invalid.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID, and end with '}}'. The UUID was not valid. Expected a valid UUID in the format 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' containing only ASCII characters.")]
    ObjectCompiledUuidInvalid,

    /// Object compiled version missing.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by a '#' character, and then a version number. Expected a '#' character, but found a newline or the end of the file instead.")]
    ObjectCompiledVersionMissing,

    /// Object compiled version is invalid.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by a '#' character, and then a version number. The version number was not valid. Expected a valid version number in the format 'x.x'. The version number must contain only '.' or the characters \"0\"..\"9\". Invalid character found instead.")]
    ObjectCompiledVersionInvalid,

    /// Object compiled unknown1 missing.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by '#', a version number, followed by another '#', then an 'unknown1' value followed by \"; \". Expected \"; \", but found a newline or the end of the file instead.")]
    ObjectCompiledUnknown1Missing,

    /// Object compiled file name not found.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by '#', a version number, followed by another '#', then an 'unknown1' value followed by \"; \", and ending with the object's file name. Expected the object's file name, but found a newline or the end of the file instead.")]
    ObjectCompiledFileNameNotFound,

    /// Module name not found.
    #[error("'Module' line is invalid. Expected a module name followed by a \"; \". Found a newline or the end of the file instead.")]
    ModuleNameNotFound,

    /// Module file name not found.
    #[error("'Module' line is invalid. Expected a module name followed by a \"; \", followed by the module file name. Found a newline or the end of the file instead.")]
    ModuleFileNameNotFound,

    /// Class name not found.
    #[error("'Class' line is invalid. Expected a class name followed by a \"; \". Found a newline or the end of the file instead.")]
    ClassNameNotFound,

    /// Class file name not found.
    #[error("'Class' line is invalid. Expected a class name followed by a \"; \", followed by the class file name. Found a newline or the end of the file instead.")]
    ClassFileNameNotFound,

    /// Path value not found.
    #[error("'{parameter_line_name}' line is invalid. Expected a '{parameter_line_name}' path after the equal sign '='. Found a newline or the end of the file instead.")]
    PathValueNotFound {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value not found.
    #[error("'{parameter_line_name}' line is invalid. Expected a quoted '{parameter_line_name}' value after the equal sign '='. Found a newline or the end of the file instead.")]
    ParameterValueNotFound {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value missing opening quote.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing an opening quote. Expected a closing quote for the '{parameter_line_name}' value.")]
    ParameterValueMissingOpeningQuote {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value missing matching quote.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing a matching quote. Expected a closing quote for the '{parameter_line_name}' value.")]
    ParameterValueMissingMatchingQuote {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value missing quotes.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing both opening and closing quotes. Expected a quoted '{parameter_line_name}' value after the equal sign '='.")]
    ParameterValueMissingQuotes {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value is invalid.
    #[error("'{parameter_line_name}' line is invalid. '{invalid_value}' is not a valid value for '{parameter_line_name}'. Only {valid_value_message} are valid values for '{parameter_line_name}'.")]
    ParameterValueInvalid {
        /// The parameter line name.
        parameter_line_name: String,
        /// The invalid value that was found.
        invalid_value: String,
        /// Valid value message.
        valid_value_message: String,
    },

    /// `DllBaseAddress` not found.
    #[error("'DllBaseAddress' line is invalid. Expected a hex address after the equal sign '='. Found a newline or the end of the file instead.")]
    DllBaseAddressNotFound,

    /// `DllBaseAddress` missing hex prefix.
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address beginning with '&h' after the equal sign '='.")]
    DllBaseAddressMissingHexPrefix,

    /// `DllBaseAddress` unparsable.
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address after the equal sign '=' beginning with '&h'. Unable to parse hex value '{hex_value}'.")]
    DllBaseAddressUnparsable {
        /// The hex value that couldn't be parsed.
        hex_value: String,
    },

    /// `DllBaseAddress` unparsable (empty).
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address after the equal sign '=' beginning with '&h'. Unable to parse empty hex value.")]
    DllBaseAddressUnparsableEmpty,

    /// Parameter line is unknown.
    #[error("'{line}' line is unknown.")]
    ParameterLineUnknown {
        /// The parameter line that is unknown.
        line: String,
    },
}
