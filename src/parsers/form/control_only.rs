//! Control-only parsing for VB6 Form files.
//!
//! This module provides a fast path for extracting just the VERSION header and
//! root control structure from a VB6 Form file, without parsing the code sections
//! into a full CST. This is useful for scenarios where:
//!
//! - Only UI/control information is needed (layout, properties, control hierarchy)
//! - Code analysis can be deferred or skipped entirely
//! - Performance is critical and full CST generation is too expensive
//! - Streaming/partial parsing is desired for large projects
//!
//! # Example
//!
//! ```rust
//! use vb6parse::{SourceFile, tokenize, FormFile};
//!
//! let source_bytes = b"VERSION 5.00\nBegin VB.Form Form1\n   Caption = \"Test\"\nEnd\n";
//! let source = SourceFile::decode_with_replacement("test.frm", source_bytes).unwrap();
//! let mut source_stream = source.source_stream();
//! let result = tokenize(&mut source_stream);
//! let (token_stream, _failures) = result.unpack();
//!
//! if let Some(ts) = token_stream {
//!     // Parse VERSION + control only (fast path)
//!     let result = FormFile::parse_control_only(ts);
//!     let (parse_result, failures) = result.unpack();
//!
//!     if let Some((version, control, _remaining_tokens)) = parse_result {
//!         if let Some(v) = version {
//!             println!("VERSION {}.{}", v.major, v.minor);
//!         }
//!
//!         if let Some(ctrl) = control {
//!             println!("Control: {}", ctrl.name);
//!         }
//!     }
//! }
//! ```

use crate::{
    errors::FormErrorKind,
    language::Control,
    parsers::{header::FileFormatVersion, ParseResult},
    TokenStream,
};

/// Result of control-only parsing containing:
/// - Optional VERSION (5.00, etc.) - may be absent in older files
/// - Optional Control (Form/MDIForm/UserControl) - None if parsing failed
/// - Remaining `TokenStream` positioned after control block
///
/// Both fields are optional to support partial success:
/// - VERSION may be missing in older .frm files
/// - Control may fail to parse while VERSION succeeds
/// - Failures are collected in the `ParseResult` wrapper
pub type ControlOnlyResult<'a> = (Option<FileFormatVersion>, Option<Control>, TokenStream<'a>);

/// Parses only the VERSION header and root control from a `TokenStream`.
///
/// This function consumes the `TokenStream` and parses:
/// 1. VERSION statement (if present at current position)
/// 2. The root control's BEGIN...END block
///
/// It returns a new `TokenStream` positioned after the control block,
/// allowing the caller to continue parsing attributes, objects, or code
/// sections if needed.
///
/// This is faster than full `FormFile` parsing because it:
/// - Skips CST construction for VERSION/control (direct extraction)
/// - Stops parsing after control block (doesn't parse code sections)
/// - Zero-copy design (moves `TokenStream` ownership)
///
/// # Arguments
///
/// * `token_stream` - `TokenStream` to parse from (consumed)
///
/// # Returns
///
/// * `ParseResult<ControlOnlyResult, FormErrorKind>` containing:
///   - `Option<FileFormatVersion>` - VERSION if found, None otherwise
///   - `Option<Control>` - Parsed root control, None if parsing failed
///   - `TokenStream` - Remaining tokens positioned after control block
///   - Failures vector with any warnings/errors encountered
///
/// # Example
///
/// ```rust
/// use vb6parse::{SourceFile, tokenize};
/// use vb6parse::parsers::form::control_only::parse_control_from_tokens;
///
/// let source_bytes = b"VERSION 5.00\nBegin VB.Form Form1\nEnd\n";
/// let source = SourceFile::decode_with_replacement("test.frm", source_bytes).unwrap();
/// let mut source_stream = source.source_stream();
/// let result = tokenize(&mut source_stream);
/// let (token_stream, _failures) = result.unpack();
///
/// if let Some(ts) = token_stream {
///     let result = parse_control_from_tokens(ts);
///     let (parse_result, failures) = result.unpack();
///     
///     if !failures.is_empty() {
///         for failure in &failures {
///             eprintln!("Warning: {:?}", failure);
///         }
///     }
///
///     if let Some((version, control, _remaining)) = parse_result {
///         // Use version and control here
///     }
/// }
/// ```
#[must_use]
pub fn parse_control_from_tokens(
    token_stream: TokenStream<'_>,
) -> ParseResult<'_, ControlOnlyResult<'_>, FormErrorKind> {
    // Convert TokenStream to tokens vector
    let tokens = token_stream.into_tokens();

    // Create parser in direct extraction mode
    let mut parser = crate::parsers::cst::Parser::new_direct_extraction(tokens, 0);

    // Parse VERSION directly (no CST overhead)
    let (version_opt, version_failures) = parser.parse_version_direct().unpack();

    // Parse control directly (no CST overhead)
    let (control_opt, control_failures) = parser.parse_properties_block_to_control().unpack();

    // Collect all failures
    let mut failures = Vec::new();
    failures.extend(version_failures);
    failures.extend(control_failures);

    // Get remaining tokens
    let remaining_tokens = parser.into_tokens();
    let remaining_stream = TokenStream::from_tokens(remaining_tokens);

    // Return result tuple
    ParseResult::new(Some((version_opt, control_opt, remaining_stream)), failures)
}
