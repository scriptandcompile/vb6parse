//! This module provides the `SourceFile` struct, which represents a VB6
//! source file along with its content and filename. It includes methods to
//! read a source file from disk and decode its content using Windows-1252
//! encoding, with options for handling invalid characters.
//!
//! The `SourceFile` struct is essential for parsing VB6 source files,
//! as it provides the necessary functionality to read and decode the
//! source code before further processing.
//!
//! # Example
//! ```no_run
//! use vb6parse::SourceFile;
//!
//! let source_file = SourceFile::from_file("path/to/module.bas").unwrap();
//! ```
//!
//! # Errors
//! The methods in this module return `ErrorDetails` when errors occur
//! during file reading or decoding
//!
//! # Encoding
//! The library assumes that VB6 source files are encoded in Windows-1252.
//! If the source file contains invalid characters, the library can either
//! replace them with a placeholder or return an error, depending on the
//! method used.
//!
//! # See Also
//! - [`SourceStream`]: for low-level character stream
//! - [`ErrorDetails`]: for error handling details

use std::borrow::Cow;
use std::fmt::Display;
use std::fs;
use std::path::Path;

use crate::errors::{ErrorDetails, SourceFileErrorKind};
use crate::parsers::SourceStream;

use encoding_rs::{mem::utf8_latin1_up_to, CoderResult, WINDOWS_1252};

/// Represents a VB6 source file with its content and filename.
/// This struct provides methods to read and decode source files
/// using Windows-1252 encoding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceFile {
    /// The content of the source file as a `String`.
    file_content: String,
    /// The name of the source file.
    pub file_name: String,
}

impl Display for SourceFile {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "SourceFile {{ file name: '{}', content len: {} }}",
            self.file_name,
            self.file_content.len()
        )
    }
}

impl SourceFile {
    /// Creates a `SourceFile` by reading from a file path.
    ///
    /// This method reads the file at the given path, decodes it using Windows-1252 encoding
    /// with replacement for invalid characters, and extracts the filename from the path.
    ///
    /// # Arguments
    ///
    /// * `path` - A path to the file to read
    ///
    /// # Returns
    ///
    /// Returns a `Result` containing either:
    /// - `Ok(SourceFile)` - Successfully read and decoded file
    /// - `Err(ErrorDetails)` - Error reading the file or decoding its contents
    ///
    /// # Errors
    ///
    /// This function will return an error if:
    /// - The file cannot be read
    /// - The file content cannot be decoded
    ///
    /// # Example
    ///
    /// ```no_run
    /// use vb6parse::SourceFile;
    ///
    /// let source_file = SourceFile::from_file("path/to/module.bas").unwrap();
    /// ```
    pub fn from_file<P: AsRef<Path>>(
        path: P,
    ) -> Result<Self, ErrorDetails<'static, SourceFileErrorKind>> {
        let path = path.as_ref();

        // Read the file contents
        let bytes = fs::read(path).map_err(|io_err| ErrorDetails {
            kind: SourceFileErrorKind::MalformedSource {
                message: format!("Failed to read file: {io_err}"),
            },
            error_offset: 0,
            source_content: Cow::Borrowed(""),
            source_name: path.display().to_string(),
            line_start: 0,
            line_end: 0,
        })?;

        // Extract the filename from the path
        let file_name = path
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("unknown")
            .to_string();

        // Decode the file using decode_with_replacement
        Self::decode_with_replacement(file_name, &bytes).map_err(|err| ErrorDetails {
            kind: err.kind,
            error_offset: err.error_offset,
            source_content: Cow::Owned(err.source_content.into_owned()),
            source_name: err.source_name,
            line_start: err.line_start,
            line_end: err.line_end,
        })
    }

    /// Creates a `SourceFile` from a file name and source code string.
    ///
    /// # Arguments
    ///
    /// * `file_name` - The name of the source file
    /// * `source_code` - The source code as a string
    ///
    /// # Returns
    ///
    /// Returns a `SourceFile` instance.
    #[must_use]
    pub fn from_string(file_name: impl Into<String>, source_code: impl Into<String>) -> Self {
        SourceFile {
            file_name: file_name.into(),
            file_content: source_code.into(),
        }
    }

    /// Decodes the source code using Windows-1252 encoding with replacement for invalid characters.
    ///
    /// # Arguments
    ///
    /// * `file_name` - The name of the source file
    /// * `source_code` - The byte slice containing the source code to decode
    ///
    /// # Returns
    ///
    /// Returns a `Result` containing either:
    /// - `Ok(SourceFile)` - Successfully decoded source file
    /// - `Err(ErrorDetails)` - Error decoding the source code
    ///
    /// # Errors
    ///
    /// This function will return an error if the source code contains invalid characters
    /// that cannot be replaced. Any character that is representable within the Windows-1252
    /// encoding will be decoded successfully and will be replaced with the Unicode equivalent.
    ///
    /// A good example of invalid characters would be any chinese characters, as they are not
    /// representable within the Windows-1252 encoding.
    pub fn decode_with_replacement(
        file_name: impl Into<String>,
        source_code: &[u8],
    ) -> Result<Self, ErrorDetails<'_, SourceFileErrorKind>> {
        Self::decode_internal(file_name, source_code, true)
    }

    fn decode_internal(
        file_name: impl Into<String>,
        source_code: &[u8],
        allow_replacement: bool,
    ) -> Result<Self, ErrorDetails<'_, SourceFileErrorKind>> {
        let mut decoder = WINDOWS_1252.new_decoder();

        let Some(max_len) = decoder.max_utf8_buffer_length(source_code.len()) else {
            return Err(ErrorDetails {
                kind: SourceFileErrorKind::MalformedSource {
                    message: "Failed to decode the source code. '{file_name}' was empty.".into(),
                },
                error_offset: 0,
                source_content: "".into(),
                source_name: file_name.into().clone(),
                line_start: 0,
                line_end: 0,
            });
        };

        let file_name = file_name.into();
        let mut source_file = SourceFile {
            file_name: file_name.clone(),
            file_content: String::with_capacity(max_len),
        };

        let last = true;
        let (coder_result, attempted_decode_len, all_processed) =
            decoder.decode_to_string(source_code, &mut source_file.file_content, last);

        if source_file.file_content.len() == source_code.len() {
            // It looks like we actually succeeded even if the coder_result might be
            // confused at that.
            return Ok(source_file);
        }

        if (!all_processed && !allow_replacement) || coder_result == CoderResult::OutputFull {
            let mut decoded_len = utf8_latin1_up_to(source_code);
            let mut error_offset = decoded_len - 1;

            // Looks like we actually succeeded even if the coder_result might be
            // confused at that.
            if attempted_decode_len == decoded_len {
                return Ok(source_file);
            }

            let text_up_to_error = if let Ok(v) = str::from_utf8(&source_code[0..decoded_len]) {
                v.to_owned()
            } else {
                // For some reason, even though this should never happen
                // we ended up here. Oh well. Report that things failed at
                // the start of the file since we can't pinpoint the exact
                // location.
                error_offset = 0;
                decoded_len = 0;
                String::new()
            };

            let details = ErrorDetails {
                kind: SourceFileErrorKind::MalformedSource {
                    message: format!(
                        r"Failed to decode the source file. '{file_name}' may not use latin-1 (Windows-1252) code page. 
Currently, only latin-1 source code is supported."
                    ),
                },
                source_content: Cow::Owned(text_up_to_error),
                source_name: file_name,
                error_offset,
                line_start: 0,
                line_end: decoded_len,
            };

            return Err(details);
        }

        Ok(source_file)
    }

    /// Decodes the source code using Windows-1252 encoding without allowing replacement for invalid characters.
    ///
    /// # Arguments
    ///
    /// * `file_name` - The name of the source file
    /// * `source_code` - The byte slice containing the source code to decode
    ///
    /// # Returns
    ///
    /// Returns a `Result` containing either:
    /// - `Ok(SourceFile)` - Successfully decoded source file
    /// - `Err(ErrorDetails)` - Error decoding the source code
    ///
    /// # Errors
    ///
    /// This function will return an error if the source code contains any invalid characters
    /// that cannot be represented in Windows-1252 encoding. All characters in the source code
    /// must be valid Windows-1252 characters for successful decoding. A good example of invalid
    /// characters would be any chinese characters, as they are not representable within the
    /// Windows-1252 encoding.
    pub fn decode(
        file_name: impl Into<String>,
        source_code: &[u8],
    ) -> Result<Self, ErrorDetails<'_, SourceFileErrorKind>> {
        Self::decode_internal(file_name, source_code, false)
    }

    /// Creates a `SourceStream` from the `SourceFile`.
    ///
    /// This method initializes a `SourceStream` using the file name and content
    /// of the `SourceFile`.
    ///
    /// # Returns
    ///
    /// Returns a `SourceStream` instance.
    #[must_use]
    pub fn get_source_stream(&'_ self) -> SourceStream<'_> {
        let source_stream = SourceStream::new(self.file_name.clone(), self.file_content.as_str());

        source_stream
    }
}
