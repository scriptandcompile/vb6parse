use std::borrow::Cow;
use std::fs;
use std::path::Path;

use crate::errors::{ErrorDetails, SourceFileErrorKind};
use crate::parsers::SourceStream;

use encoding_rs::{mem::utf8_latin1_up_to, CoderResult, WINDOWS_1252};

pub struct SourceFile {
    file_content: String,
    pub file_name: String,
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
                message: format!("Failed to read file: {}", io_err),
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

    pub fn decode(
        file_name: impl Into<String>,
        source_code: &[u8],
    ) -> Result<Self, ErrorDetails<'_, SourceFileErrorKind>> {
        Self::decode_internal(file_name, source_code, false)
    }

    #[must_use]
    pub fn get_source_stream(&'_ self) -> SourceStream<'_> {
        let source_stream = SourceStream::new(self.file_name.clone(), self.file_content.as_str());

        source_stream
    }
}
