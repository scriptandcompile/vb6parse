/// This module is used by the `project` module to build error messages when parsing a project file.
///
use crate::errors::{DiagnosticLabel, ParserContext, ProjectError};
use crate::io::SourceStream;

use std::fmt::Debug;

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

pub fn parameter_optional_missing_value_eof_error<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
) {
    // The input ends right after the equal!
    // weird error and indicates the system is basically done, but still need to
    // spit out a reasonable error message.
    let value_span = input.span_range(parameter_start - 1, parameter_start);
    // We don't have a value so we want the valid values.
    let valid_value_message = "Text string values are valid here as well as !None!, (None), !(None)!, \"(None)\", \"!None!\", or \"!(None)!\" to indicate no value is selected.".to_string();
    let error = ctx
        .error_with(
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
        .with_note(format!("{line_type}=\"!None!\""));
    ctx.push_error(error);
}

pub fn parameter_with_default_missing_value_eof_error<'a, T>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
) where
    T: 'a
        + TryFrom<&'a str, Error = String>
        + IntoEnumIterator
        + EnumMessage
        + Debug
        + Into<i16>
        + Default
        + Copy,
{
    // The input ends right after the equal!
    // weird error and indicates the system is basically done, but still need to
    // spit out a reasonable error message.
    let value_span = input.span_range(parameter_start - 1, parameter_start);
    // We don't have a value so we want the valid values.
    let valid_value_message = format_valid_enum_values::<T>();
    let error = ctx
        .error_with(
            value_span,
            ProjectError::ParameterWithDefaultValueNotFoundEOF {
                parameter_line_name: line_type.to_string(),
                valid_value_message,
            },
        )
        .with_label(DiagnosticLabel::new(
            value_span,
            format!("'{line_type}' must have a double qouted value and end with a newline."),
        )) // only a start quote in the note since we already have the end quote value.
        .with_note(format!("{line_type}=\"{}\"", T::default().into()));
    ctx.push_error(error);
}

pub fn parameter_missing_value_opening_quote_error<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
    parameter_value: &str,
) {
    // The value ends with a quote but does not start with one.
    // This is an error, so we return an error.
    let value_span = input.span_range(parameter_start, parameter_start + parameter_value.len());
    let error = ctx
        .error_with(
            value_span,
            ProjectError::ParameterValueMissingOpeningQuote {
                parameter_line_name: line_type.to_string(),
            },
        )
        .with_label(DiagnosticLabel::new(
            value_span,
            format!("'{line_type}' value must be surrounded by double quotes."),
        )) // only a start quote in the note since we already have the end quote value.
        .with_note(format!("{line_type}=\"{parameter_value}"));
    ctx.push_error(error);
}

pub fn parameter_with_default_missing_value_and_closing_quote_error<'a, T>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
    parameter_value: &str,
) where
    T: 'a
        + TryFrom<&'a str, Error = String>
        + IntoEnumIterator
        + EnumMessage
        + Debug
        + Into<i16>
        + Default
        + Copy,
{
    // The value starts with a quote and is only a single character wide. This means the entire
    // parameter value consists of a single double qoute character: '"'
    let value_span = input.span_range(parameter_start, parameter_start + parameter_value.len());
    // We do not have a valid parameter value, so we return an error.
    let valid_value_message = format_valid_enum_values::<T>();
    let default_value = T::default().into();
    let note_message = format!("{line_type}=\"{default_value}\"");

    let error = ctx
        .error_with(
            value_span,
            ProjectError::ParameterValueMissingClosingQuoteAndValue {
                parameter_line_name: line_type.to_string(),
                valid_value_message,
            },
        )
        .with_label(DiagnosticLabel::new(
            value_span,
            format!("'{line_type}' value must be surrounded by double quotes."),
        )) // only an end quote in the note since we already have the start quote value.
        .with_note(note_message);
    ctx.push_error(error);
}

pub(crate) fn parameter_missing_value_and_closing_quote_error<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
    parameter_value: &str,
) {
    // The value ends with a quote but does not start with one.
    // This is an error, so we return an error.
    let value_span = input.span_range(parameter_start, parameter_start + parameter_value.len());
    let error = ctx
        .error_with(
            value_span,
            ProjectError::ParameterValueMissingClosingQuote {
                parameter_line_name: line_type.to_string(),
            },
        )
        .with_label(DiagnosticLabel::new(
            value_span,
            format!("'{line_type}' value must be surrounded by double quotes."),
        )) // only an end quote in the note since we already have the start quote value.
        .with_note(format!("{line_type}={parameter_value}\""));
    ctx.push_error(error);
}

pub(crate) fn parameter_with_default_missing_quotes_error<'a, T>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
    parameter_value: &'a str,
) where
    T: 'a
        + TryFrom<&'a str, Error = String>
        + IntoEnumIterator
        + EnumMessage
        + Debug
        + Into<i16>
        + Default
        + Copy,
{
    // The value does not start or end with a quote but there *is* a number here.
    // this is not the same as not having an start or end and having a length of zero.
    // this is likely something like 'CompatibleMode=1' and needs to mention the
    // double qouting.

    // We do not have a valid parameter value, so we return an error.
    let valid_value_message = format_valid_enum_values::<T>();

    // We have a value, but it's not qouted. If the value makes sense
    // for this conversion, we should have the note show the qouted values.
    // if it's an invalid value, show the user an example with the default
    // value.
    let note_message = if T::try_from(parameter_value).is_ok() {
        format!("{line_type}=\"{parameter_value}\"")
    } else {
        let default_value = T::default().into();
        format!("{line_type}=\"{default_value}\"")
    };

    let value_span = input.span_at(parameter_start);
    let error = ctx
        .error_with(
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
        .with_note(note_message);
    ctx.push_error(error);
}

pub(crate) fn parameter_with_default_invalid_value_error<'a, T>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
    parameter_value: &'a str,
) where
    T: 'a
        + TryFrom<&'a str, Error = String>
        + IntoEnumIterator
        + EnumMessage
        + Debug
        + Into<i16>
        + Default
        + Copy,
{
    // We have a parameter value that is invalid, so we return an error.
    let valid_value_message = format_valid_enum_values::<T>();

    let value_span = input.span_at(parameter_start + 1);
    let error = ctx
        .error_with(
            value_span,
            ProjectError::ParameterValueInvalid {
                parameter_line_name: line_type.to_string(),
                invalid_value: parameter_value.to_string(),
                valid_value_message,
            },
        )
        .with_label(DiagnosticLabel::new(value_span, "invalid value"))
        .with_note("Change the quoted value to one of the valid values.");
    ctx.push_error(error);
}

pub(crate) fn parameter_with_default_missing_value_and_quotes_error<'a, T>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
    parameter_start: usize,
) where
    T: 'a
        + TryFrom<&'a str, Error = String>
        + IntoEnumIterator
        + EnumMessage
        + Debug
        + Into<i16>
        + Default
        + Copy,
{
    // The value does not start or end with a quote but there *is* a number here.
    // this is not the same as not having an start or end and having a length of zero.
    // this is likely something like 'CompatibleMode=' and needs to show the default vale.

    // We do not have a valid parameter value, so we return an error.
    let valid_value_message = format_valid_enum_values::<T>();

    // We don't have a value or qoutes.
    // show the user an example with the default value.

    let default_value = T::default().into();
    let note_message = format!("{line_type}=\"{default_value}\"");

    let value_span = input.span_at(parameter_start);
    let error = ctx
        .error_with(
            value_span,
            ProjectError::ParameterWithDefaultValueNotFound {
                parameter_line_name: line_type.to_string(),
                valid_value_message,
            },
        )
        .with_label(DiagnosticLabel::new(
            value_span,
            format!("'{line_type}' value must be one of the valid values contained within double qoutes."),
        ))
        .with_note(note_message);
    ctx.push_error(error);
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
}
