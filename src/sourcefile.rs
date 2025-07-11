use std::borrow::Cow;

use crate::errors::{ErrorDetails, SourceFileErrorKind};
use crate::parsers::SourceStream;

use encoding_rs::{mem::*, CoderResult, WINDOWS_1252};

pub struct SourceFile {
    file_content: String,
    pub file_name: String,
}

impl SourceFile {
    pub fn decode_with_replacement(
        file_name: impl Into<String>,
        source_code: &[u8],
    ) -> Result<Self, ErrorDetails<SourceFileErrorKind>> {
        Self::decode_internal(file_name, source_code, true)
    }

    fn decode_internal(
        file_name: impl Into<String>,
        source_code: &[u8],
        allow_replacement: bool,
    ) -> Result<Self, ErrorDetails<SourceFileErrorKind>> {
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
            // looks like we actualy succeded even if the coder_result might be
            // confused at that.
            return Ok(source_file);
        }

        if (all_processed == false && allow_replacement == false)
            || coder_result == CoderResult::OutputFull
        {
            let mut decoded_len = utf8_latin1_up_to(source_code);
            let mut error_offset = decoded_len - 1;

            // looks like we actualy succeded even if the coder_result might be
            // confused at that.
            if attempted_decode_len == decoded_len {
                return Ok(source_file);
            }

            let text_upto_error = match str::from_utf8(&source_code[0..decoded_len]) {
                Ok(v) => v.to_owned(),
                Err(_) => {
                    // For some reason, even though this should never happen
                    // we ended up here. Oh well. Report that things failed at
                    // the start of the file since we can't pinpoint the exact
                    // location.
                    error_offset = 0;
                    decoded_len = 0;
                    "".to_owned()
                }
            };

            let details = ErrorDetails {
                kind: SourceFileErrorKind::MalformedSource {
                    message: format!(
                        r"Failed to decode the source file. '{file_name}' may not use latin-1 (Windows-1252) code page. 
Currently, only latin-1 source code is supported."
                    ),
                },
                source_content: Cow::Owned(text_upto_error),
                source_name: file_name,
                error_offset: error_offset,
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
    ) -> Result<Self, ErrorDetails<SourceFileErrorKind>> {
        Self::decode_internal(file_name, source_code, false)
    }

    pub fn get_source_stream(&self) -> SourceStream {
        let source_stream = SourceStream::new(self.file_name.clone(), self.file_content.as_str());

        source_stream
    }
}
