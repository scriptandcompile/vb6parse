/// This module is used by the `project` module to build error messages when parsing a project file.
///
use crate::errors::{DiagnosticLabel, ParserContext, ProjectError};
use crate::io::SourceStream;

use std::fmt::Debug;
use std::marker::PhantomData;

use strum::{EnumMessage, IntoEnumIterator};

/// Formats all valid values for an enum type as a string.
///
/// Returns a comma-separated list of valid enum values in the format:
/// ```text
/// 'numeric value' "message"
/// ```
/// for each variant, with the final variant
/// being appended with:
/// ```text
/// ", and 'numeric value' "message"
/// ```
/// This makes it slightly nicer to read.
///
/// Long live the Oxford comma!
///
/// # Example
/// For an enum with values 0, 1, 2 this should return:
/// `'0' "No Compatibility", '1' "Project Compatibility", and '2' "Compatible Exe Mode"`
pub fn format_valid_enum_values<T>() -> String
where
    T: IntoEnumIterator + EnumMessage + Debug + Into<i16> + Copy,
{
    match T::iter()
        .map(|v| {
            let numeric: i16 = v.into();
            format!("'{:?}' {:#?}", numeric, v.get_message().unwrap_or(""))
        })
        .collect::<Vec<_>>()
        .split_last()
    {
        Some((last, elements)) => {
            format!("{}, and {}", elements.join(", "), last)
        } // we shoiuld never get a 'None' here since all
        // the enums should have multiple variants with values, but...
        None => String::new(),
    }
}

/// Represents different kinds of parameter parsing errors.
pub enum ParameterErrorKind<'a, T> {
    /// Missing value at EOF for a required parameter
    MissingValueEof,
    /// Missing value at EOF for an optional parameter
    OptionalMissingValueEof,
    /// Missing value at EOF for a parameter with a default value
    MissingValueEofWithDefault(PhantomData<T>),
    /// Missing opening quote
    MissingOpeningQuote { value: &'a str },
    /// Missing closing quote
    MissingClosingQuote { value: &'a str },
    /// Missing both value and closing quote (only has opening quote)
    MissingValueAndClosingQuote {
        value: &'a str,
        _phantom: PhantomData<T>,
    },
    /// Missing both quotes with default
    MissingQuotesWithDefault {
        value: &'a str,
        _phantom: PhantomData<T>,
    },
    /// Invalid value for enum
    InvalidValue {
        value: &'a str,
        _phantom: PhantomData<T>,
    },
    /// Missing value and quotes
    MissingValueAndQuotes(PhantomData<T>),
    /// Empty parameter value
    EmptyValue,
    /// Missing both quotes (without default)
    MissingBothQuotes,
    /// Property name not found (no '=' delimiter)
    PropertyNameNotFound,
}

/// Reports a parameter error based on the error kind.
///
/// This consolidated function replaces multiple similar error reporting functions
/// by using an enum to determine which specific error to report.
pub fn report_parameter_error<'a, T>(
    ctx: &mut ParserContext<'a>,
    input: &SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
    kind: &ParameterErrorKind<'a, T>,
) where
    T: TryFrom<&'a str, Error = String>
        + IntoEnumIterator
        + EnumMessage
        + Debug
        + Into<i16>
        + Default
        + Copy,
{
    match kind {
        ParameterErrorKind::MissingValueEof => {
            let value_span = input.span_range(parameter_start - 1, parameter_start);
            ctx.error(
                value_span,
                ProjectError::ParameterValueNotFound {
                    parameter_line_name: line_type.to_string(),
                },
            );
        }
        ParameterErrorKind::OptionalMissingValueEof => {
            let value_span = input.span_range(parameter_start - 1, parameter_start);
            let valid_value_message = "Text string values are valid here as well as !None!, (None), !(None)!, \"(None)\", \"!None!\", or \"!(None)!\" to indicate no value is selected.".to_string();
            ctx.error_with(
                value_span,
                ProjectError::ParameterWithDefaultValueNotFoundEOF {
                    parameter_line_name: line_type.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' must have a double quoted value and end with a newline."),
            ))
            .with_note(format!("{line_type}=\"!None!\""))
            .emit(ctx);
        }
        ParameterErrorKind::MissingValueEofWithDefault(_) => {
            let value_span = input.span_range(parameter_start - 1, parameter_start);
            let valid_value_message = format_valid_enum_values::<T>();
            ctx.error_with(
                value_span,
                ProjectError::ParameterWithDefaultValueNotFoundEOF {
                    parameter_line_name: line_type.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' must have a double qouted value and end with a newline."),
            ))
            .with_note(format!("{line_type}=\"{}\"", T::default().into()))
            .emit(ctx);
        }
        ParameterErrorKind::MissingOpeningQuote { value } => {
            let value_span = input.span_range(parameter_start, parameter_start + value.len());
            ctx.error_with(
                value_span,
                ProjectError::ParameterValueMissingOpeningQuote {
                    parameter_line_name: line_type.to_string(),
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be surrounded by double quotes."),
            ))
            .with_note(format!("{line_type}=\"{value}"))
            .emit(ctx);
        }
        ParameterErrorKind::MissingClosingQuote { value } => {
            let value_span = input.span_range(parameter_start, parameter_start + value.len());
            ctx.error_with(
                value_span,
                ProjectError::ParameterValueMissingClosingQuote {
                    parameter_line_name: line_type.to_string(),
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be surrounded by double quotes."),
            ))
            .with_note(format!("{line_type}={value}\""))
            .emit(ctx);
        }
        ParameterErrorKind::MissingValueAndClosingQuote { value, .. } => {
            let value_span = input.span_range(parameter_start, parameter_start + value.len());
            let valid_value_message = format_valid_enum_values::<T>();
            let default_value = T::default().into();
            let note_message = format!("{line_type}=\"{default_value}\"");

            ctx.error_with(
                value_span,
                ProjectError::ParameterValueMissingClosingQuoteAndValue {
                    parameter_line_name: line_type.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be surrounded by double quotes."),
            ))
            .with_note(note_message)
            .emit(ctx);
        }
        ParameterErrorKind::MissingQuotesWithDefault { value, .. } => {
            let valid_value_message = format_valid_enum_values::<T>();
            let note_message = if T::try_from(value).is_ok() {
                format!("{line_type}=\"{value}\"")
            } else {
                let default_value = T::default().into();
                format!("{line_type}=\"{default_value}\"")
            };

            let value_span = input.span_at(parameter_start);
            ctx.error_with(
                value_span,
                ProjectError::ParameterValueMissingQuotes {
                    parameter_line_name: line_type.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be contained within double qoutes."),
            ))
            .with_note(note_message)
            .emit(ctx);
        }
        ParameterErrorKind::InvalidValue { value, .. } => {
            let valid_value_message = format_valid_enum_values::<T>();
            let value_span = input.span_at(parameter_start + 1);
            ctx.error_with(
                value_span,
                ProjectError::ParameterValueInvalid {
                    parameter_line_name: line_type.to_string(),
                    invalid_value: value.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(value_span, "invalid value"))
            .with_note("Change the quoted value to one of the valid values.")
            .emit(ctx);
        }
        ParameterErrorKind::MissingValueAndQuotes(_) => {
            let valid_value_message = format_valid_enum_values::<T>();
            let default_value = T::default().into();
            let note_message = format!("{line_type}=\"{default_value}\"");

            let value_span = input.span_at(parameter_start);
            ctx.error_with(
                value_span,
                ProjectError::ParameterWithDefaultValueNotFound {
                    parameter_line_name: line_type.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!(
                    "'{line_type}' value must be one of the valid values contained within double qoutes."
                ),
            ))
            .with_note(note_message)
            .emit(ctx);
        }
        ParameterErrorKind::EmptyValue => {
            let value_span = input.span_at(parameter_start);
            ctx.error(
                value_span,
                ProjectError::ParameterValueNotFound {
                    parameter_line_name: line_type.to_string(),
                },
            );
        }
        ParameterErrorKind::MissingBothQuotes => {
            let value_span = input.span_at(parameter_start);
            ctx.error(
                value_span,
                ProjectError::ParameterWithoutDefaultValueMissingQuotes {
                    parameter_line_name: line_type.to_string(),
                },
            );
        }
        ParameterErrorKind::PropertyNameNotFound => {
            let value_span = input.span_at(parameter_start);
            ctx.error(value_span, ProjectError::PropertyNameNotFound);
        }
    }
}

// Helper dummy type for non-generic error functions
#[derive(Debug, Copy, Clone)]
pub struct DummyEnumType;

impl Default for DummyEnumType {
    fn default() -> Self {
        DummyEnumType
    }
}

impl From<DummyEnumType> for i16 {
    fn from(_val: DummyEnumType) -> Self {
        0
    }
}

impl IntoEnumIterator for DummyEnumType {
    type Iterator = std::iter::Empty<Self>;
    fn iter() -> Self::Iterator {
        std::iter::empty()
    }
}

impl EnumMessage for DummyEnumType {
    fn get_message(&self) -> Option<&'static str> {
        None
    }
    fn get_detailed_message(&self) -> Option<&'static str> {
        None
    }
    fn get_documentation(&self) -> Option<&'static str> {
        None
    }
    fn get_serializations(&self) -> &'static [&'static str] {
        &[]
    }
}

impl<'a> TryFrom<&'a str> for DummyEnumType {
    type Error = String;
    fn try_from(_: &'a str) -> Result<Self, Self::Error> {
        Ok(DummyEnumType)
    }
}

#[cfg(test)]
mod tests {
    use crate::errors::{ErrorKind, ParserContext, ProjectError, Severity};
    use crate::files::project::properties::*;
    use crate::io::{Comparator, SourceStream};
    use assert_matches::assert_matches;

    #[test]
    fn no_optional_value_eof() {
        use crate::files::project::parse_optional_quoted_value;
        use crate::io::{Comparator, SourceStream};

        let mut input = SourceStream::new("", "Startup=");

        let parameter_name = input.take("Startup", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result = parse_optional_quoted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_eq!(result, None);
        assert_eq!(errors[0].line_start, 0);
        assert_eq!(errors[0].line_end, 8);
        assert_eq!(errors[0].error_offset, 7);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'Startup' must have a double quoted value and end with a newline."
        );
        assert_eq!(errors[0].labels[0].span.line_start, 0);
        assert_eq!(errors[0].labels[0].span.line_end, 8);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(errors[0].labels[0].span.offset, 7);
        assert_eq!(errors[0].notes.len(), 1);
        assert_eq!(errors[0].notes[0], "Startup=\"!None!\"");
    }

    #[test]
    fn compatibility_mode_eof_after_equal() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterWithDefaultValueNotFoundEOF { .. })
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 15);
        assert_eq!(errors[0].labels[0].span.offset, 14);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' must have a double qouted value and end with a newline."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_with_only_start_quote() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingClosingQuoteAndValue { .. })
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);
        assert_eq!(errors[0].labels[0].span.line_end, 18);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be surrounded by double quotes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_is_invalid() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"5\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueInvalid { .. })
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 21);
        assert_eq!(errors[0].labels[0].span.offset, 16);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(errors[0].labels[0].message, "invalid value");
        assert_eq!(
            errors[0].notes[0],
            "Change the quoted value to one of the valid values."
        );
    }

    #[test]
    fn compatibility_mode_without_quotes() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=0\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingQuotes { .. })
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 18);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be contained within double qoutes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"0\"");
    }

    #[test]
    fn compatibility_mode_invalid_without_quotes() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=5\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingQuotes { .. })
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 18);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be contained within double qoutes."
        ); // Since the unqouted value is invalid, we should show a note with the default for 'CompatibleMode'
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_without_value() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterWithDefaultValueNotFound { .. })
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 16);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be one of the valid values contained within double qoutes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_without_end_quote() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"1\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingClosingQuote { .. })
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);
        assert_eq!(errors[0].labels[0].span.line_end, 19);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 2);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be surrounded by double quotes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_without_start_quote() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=2\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingOpeningQuote { .. })
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);
        assert_eq!(errors[0].labels[0].span.line_end, 19);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 2);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be surrounded by double quotes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"2\"");
    }

    #[test]
    fn property_name_not_found() {
        use crate::files::project::parse_property_name;

        let mut input = SourceStream::new("", "SomePropertyWithoutEquals\n");

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result = parse_property_name(&mut ctx, &mut input);

        let errors = ctx.errors();

        errors[0].print();

        assert_eq!(errors.len(), 1);
        assert_matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::PropertyNameNotFound)
        );
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(result, None);
        assert_eq!(errors[0].line_start, 0);
        assert_eq!(errors[0].line_end, 25);
        assert_eq!(errors[0].error_offset, 0);
        assert_eq!(errors[0].labels.len(), 0);
        assert_eq!(errors[0].notes.len(), 0);
    }
}
