//! Provides a character stream abstraction for parsing source files.
//! It includes functionality for tracking the current position in the stream,
//! as well as methods for peeking and consuming characters based on various criteria.
//!
//! This module is essential for building tokenizers that need to read and analyze
//! source code character by character.
//!
//! # Example
//! ```rust
//! use vb6parse::io::SourceStream;
//!
//! let source_stream = SourceStream::new("example.bas", "Dim x As Integer");
//! assert_eq!(source_stream.peek(3), Some("Dim"));
//! ```
//!
//! # See Also
//! - [`SourceFile`](crate::io::SourceFile): for reading and decoding source files
//! - [`ErrorDetails`]: for error handling details

use std::fmt::Debug;

use crate::errors::ErrorDetails;

/// A structure representing a stream of characters from a source file.
/// It holds the file name, the contents of the file, and an offset
/// indicating the current position in the stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceStream<'a> {
    /// The name of the source file.
    pub file_name: String,
    /// The contents of the source file.
    pub contents: &'a str,
    /// The current offset in the stream.
    pub offset: usize,
}

/// An enum representing the type of comparison to be used when taking characters
/// from the `SourceStream`.
/// It can be either case-sensitive or case-insensitive.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum Comparator {
    /// A case-sensitive comparison.
    CaseSensitive = 0,
    /// A case-insensitive comparison.
    CaseInsensitive = 1,
}

impl<'a> SourceStream<'a> {
    /// Creates a new `SourceStream` with the given file name and contents.
    ///
    /// The `file_name` is a `String` representing the name of the file being parsed.
    /// The `contents` is a `str` that contains the contents of the stream.
    pub fn new<S: Into<String>>(file_name: S, contents: &'a str) -> Self {
        Self {
            file_name: file_name.into(),
            contents,
            offset: 0,
        }
    }

    /// Resets the offset to the start of the stream.
    pub fn reset_to_start(&mut self) {
        self.offset = 0;
    }

    /// Moves the offset forward by `count` characters in the stream.
    ///
    /// If the `count` exceeds the length of the contents, the offset
    /// is set to the end of the contents.
    ///
    /// Note:
    /// This method moves the offset by characters, not bytes. It respects
    /// UTF-8 character boundaries.
    pub fn forward(&mut self, count: usize) {
        let end_offset = self.offset + count;

        if end_offset > self.contents.len() {
            self.offset = self.contents.len();
        } else {
            self.offset = end_offset;
        }
    }

    /// Moves the offset forward to the next line in the stream.
    ///
    /// This method consumes characters until a newline character
    /// is encountered, and then consumes the newline character itself.
    pub fn forward_to_next_line(&mut self) {
        let _ = self.take_until_newline();
        self.take_newline();
    }

    /// Returns the file name of the stream.
    #[must_use]
    pub fn file_name(&self) -> &str {
        &self.file_name
    }

    /// Returns the current offset in the stream.
    #[must_use]
    pub fn offset(&self) -> usize {
        self.offset
    }

    /// Returns the start offset of the current line in the stream.
    #[must_use]
    pub fn start_of_line(&self) -> usize {
        self.start_of_line_from(self.offset)
    }

    /// Returns the start offset of the line containing the given `offset`.
    ///
    /// This method searches backwards from the given `offset` to find the
    /// last newline character and returns the position after it. If no newline
    /// is found, it returns `0`, indicating the start of the stream.
    #[must_use]
    pub fn start_of_line_from(&self, offset: usize) -> usize {
        // Find the last newline character before the current offset
        if let Some(pos) = self.contents[..offset].rfind('\n') {
            pos + 1 // Return the position after the newline character
        } else {
            0 // If no newline found, return the start of the stream
        }
    }

    /// Returns the end offset of the current line in the stream.
    #[must_use]
    pub fn end_of_line(&self) -> usize {
        self.end_of_line_from(self.offset)
    }

    /// Returns the end offset of the line containing the given `offset`.
    ///
    /// This method searches forwards from the given `offset` to find the
    /// next newline character and returns its position. If no newline
    /// is found, it returns the length of the contents, indicating the
    /// end of the stream.
    #[must_use]
    pub fn end_of_line_from(&self, offset: usize) -> usize {
        // Find the next newline character after the current offset
        if let Some(pos) = self.contents[offset..].find('\n') {
            self.offset + pos // Return the position of the newline character
        } else {
            self.contents.len() // If no newline found, return the end of the stream
        }
    }

    /// Checks if the stream is empty, meaning the offset is at or beyond the
    /// end of the contents.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.offset >= self.contents.len()
    }

    /// Peeks at the next `count` characters in the stream without consuming them.
    #[must_use]
    pub fn peek(&self, count: usize) -> Option<&'a str> {
        let end_offset = self.offset + count;

        if end_offset > self.contents.len() {
            None
        } else {
            Some(&self.contents[self.offset..end_offset])
        }
    }

    /// Peeks at the next characters in the stream to see if they match the `compare` value.
    ///
    /// If they match, it returns the characters that matched the `str`. If they
    /// do not match, it returns `None`. This is important when parsing for a case
    /// insensitive match but the case sensitivity has further implications.
    ///
    /// The `case_sensitive` parameter determines whether the comparison is case-sensitive
    /// or case-insensitive.
    pub fn peek_text<'b>(
        &self,
        compare: impl Into<&'b str>,
        case_sensitive: Comparator,
    ) -> Option<&'a str> {
        let compare = compare.into();
        let peek_len = compare.len();
        let peek_slice = self.peek(peek_len)?;

        let matches = match case_sensitive {
            Comparator::CaseSensitive => peek_slice.eq(compare),
            Comparator::CaseInsensitive => peek_slice.eq_ignore_ascii_case(compare),
        };

        if matches {
            Some(peek_slice)
        } else {
            None
        }
    }

    /// Peeks at the next character in the stream to see if it matches a linux
    /// newline character (`\n`). If a linux newline character is found, it returns it
    /// as a `str`. If no linux newline character is found, it returns `None`.
    #[must_use]
    pub fn peek_linux_newline(&self) -> Option<&'a str> {
        let preview_character = self.peek(1);
        if preview_character == Some("\n") {
            return preview_character;
        }

        None
    }

    /// Peeks at the next pair of characters in the stream to see if they match
    /// a Windows newline character pair (`\r\n`). If a Windows newline character pair
    /// is found, it returns it as a `str`. If no Windows newline character pair
    /// is found, it returns `None`.
    #[must_use]
    pub fn peek_windows_newline(&self) -> Option<&'a str> {
        let preview_pair = self.peek(2);
        if preview_pair == Some("\r\n") {
            return preview_pair;
        }

        None
    }

    /// Peeks at the next characters in the stream to see if they match a Linux
    /// newline character or Windows newline character pair.
    /// Windows "\r\n\" or Linux "\n". If a newline character or pair is found,
    /// it returns the `str`.
    /// otherwise it returns `None`.
    #[must_use]
    pub fn peek_newline(&self) -> Option<&'a str> {
        self.peek_windows_newline()
            .or_else(|| self.peek_linux_newline())
    }

    /// Takes characters from the stream if they match the `compare` str.
    #[must_use]
    pub fn take<'b>(
        &mut self,
        compare: impl Into<&'b str>,
        case_sensitive: Comparator,
    ) -> Option<&'a str> {
        let mut end_offset = self.offset;
        let compare = compare.into();
        let compare_len = compare.len();
        let compare_slice = self.peek(compare_len)?;

        let matches = match case_sensitive {
            Comparator::CaseSensitive => compare_slice.eq(compare),
            Comparator::CaseInsensitive => compare_slice.eq_ignore_ascii_case(compare),
        };

        if !matches {
            return None;
        }

        end_offset += compare_len;
        let result = &self.contents[self.offset..end_offset];
        self.offset = end_offset;
        Some(result)
    }

    /// Takes a specific number of characters from the stream.
    ///
    /// If the requested number of characters exceeds the remaining characters
    /// in the stream, it returns `None`.
    #[must_use]
    pub fn take_count(&mut self, count: usize) -> Option<&'a str> {
        let end_offset = self.offset + count;

        if end_offset > self.contents.len() {
            None
        } else {
            self.offset = end_offset;
            Some(&self.contents[self.offset..end_offset])
        }
    }

    /// Takes characters from the stream until a character that matches the
    /// compare `str` is encountered or the end of the stream is reached.
    ///
    /// If a match is found, it returns a tuple containing a `str` for the
    /// characters taken from the stream until the match was found,
    /// and the matched characters.
    ///
    /// If no match is found, it returns `None`.
    ///
    /// The `case_sensitive` parameter determines whether the comparison is
    /// case-sensitive or case-insensitive.
    ///
    /// Note:
    /// This does not consume the matched characters from the stream, it only
    /// consumes the characters that were taken until the match was found.
    ///
    /// if the matched characters are needed to be consumed, use `take` after using
    /// the matched characters to consume them.
    pub fn take_until<'b>(
        &mut self,
        compare: impl Into<&'b str>,
        case_sensitive: Comparator,
    ) -> Option<(&'a str, &'a str)> {
        let mut end_offset = self.offset;
        let compare = compare.into();
        let content_len = self.contents.len();
        let compare_len = compare.len();

        while end_offset < content_len {
            if end_offset + compare_len > content_len {
                return None;
            }

            let slice = &self.contents[end_offset..end_offset + compare_len];
            let matches = match case_sensitive {
                Comparator::CaseSensitive => slice.eq(compare),
                Comparator::CaseInsensitive => slice.eq_ignore_ascii_case(compare),
            };

            if matches {
                let result = &self.contents[self.offset..end_offset];
                self.offset = end_offset;
                return Some((result, slice));
            }
            end_offset += 1;
        }

        None
    }

    /// Takes characters from the stream until a character that does not match
    /// any of the `compare_set` strings is encountered or the end of the stream
    /// is reached.
    ///
    /// If a non-matching character is found, it returns a tuple containing a `str`
    /// for the characters taken from the stream until the non-matching character
    /// was found, and the non-matching character.
    pub fn take_until_not(
        &mut self,
        compare_set: &[&str],
        case_sensitive: Comparator,
    ) -> Option<(&'a str, &'a str)> {
        let mut end_offset = self.offset;
        let content_len = self.contents.len();

        while end_offset < content_len {
            if compare_set.iter().any(|&s| match case_sensitive {
                Comparator::CaseSensitive => s.eq(&self.contents[end_offset..end_offset + s.len()]),
                Comparator::CaseInsensitive => {
                    s.eq_ignore_ascii_case(&self.contents[end_offset..end_offset + s.len()])
                }
            }) {
                end_offset += 1;
            } else {
                let result = &self.contents[self.offset..end_offset];
                self.offset = end_offset;
                return Some((result, &self.contents[end_offset..=end_offset]));
            }
        }

        if end_offset > self.offset {
            let result = &self.contents[self.offset..end_offset];
            self.offset = end_offset;
            Some((result, &self.contents[end_offset..=end_offset]))
        } else {
            None
        }
    }

    /// Takes a newline character - Windows "\r\n" or Linux "\n" - from the stream
    /// if it exists. If a newline character is found, it consumes it and returns
    /// it as a `str`. If no newline character is found, it returns `None`.
    pub fn take_newline(&mut self) -> Option<&'a str> {
        self.take_windows_newline()
            .or_else(|| self.take_linux_newline())
    }

    /// Takes a Windows newline character pair "\r\n" from the stream if the pair
    /// exists at the current offset, or returns `None` if it does not.
    pub fn take_windows_newline(&mut self) -> Option<&'a str> {
        self.take("\r\n", Comparator::CaseSensitive)
    }

    /// Takes a Linux newline character "\n" from the stream if it exists or
    /// returns `None` if it does not.
    pub fn take_linux_newline(&mut self) -> Option<&'a str> {
        self.take("\n", Comparator::CaseSensitive)
    }

    /// Takes characters from the stream until a character that matches the predicate
    /// function is encountered or the end of the stream is reached. If `none_on_eos`
    /// is true than reaching the end of the stream will return `None` if the scanning
    /// is still going on.
    /// Otherwise, hitting End-of-Stream will just return the currently matched pattern.
    ///
    /// Setting `none_on_eos` to false is useful in predicates when the End-of-Stream
    /// would also indicate the end of the predicate pattern.
    ///
    /// This method is useful for parsing various types of content where you need
    /// to consume characters until a specific condition is met, such as whitespace,
    /// alphabetic characters, alphanumeric characters, digits, punctuation, etc.
    pub fn take_until_lambda(
        &mut self,
        mut predicate: impl FnMut(char) -> bool,
        none_on_eos: bool,
    ) -> Option<&'a str> {
        let mut end_offset = self.offset;
        let content_len = self.contents.len();

        while end_offset < content_len {
            let current_char = self.contents[end_offset..].chars().next()?;

            if predicate(current_char) {
                let result = &self.contents[self.offset..end_offset];
                self.offset = end_offset;

                if result.is_empty() {
                    return None;
                }
                return Some(result);
            }
            end_offset += current_char.len_utf8();
        }

        if none_on_eos {
            None
        } else {
            let result = &self.contents[self.offset..end_offset];
            self.offset = end_offset;

            if result.is_empty() {
                return None;
            }

            Some(result)
        }
    }

    /// Takes characters from the stream until a character that is not a whitespace
    /// (including carriage return and newline) is encountered or the end of the
    /// stream is reached.
    pub fn take_ascii_whitespaces(&mut self) -> Option<&'a str> {
        self.take_until_lambda(
            |character| !character.is_ascii_whitespace() || character == '\r' || character == '\n',
            false,
        )
    }

    /// Takes characters from the stream until a character that is not an ASCII
    /// alphabetic character (a-z, A-Z) is encountered or the end of the stream
    /// is reached.
    pub fn take_ascii_alphabetics(&mut self) -> Option<&'a str> {
        self.take_until_lambda(|character| !character.is_ascii_alphabetic(), false)
    }

    /// Takes a single character from the stream until a character that is not
    /// an ASCII alphabetic (a-z, A-Z) is encountered or the end of the stream is
    /// reached.
    pub fn take_ascii_alphabetic(&mut self) -> Option<&'a str> {
        match self.take_count(1usize) {
            None => None,
            Some(character) => {
                if character.chars().next()?.is_ascii_alphabetic() {
                    Some(character)
                } else {
                    None
                }
            }
        }
    }

    /// Takes characters from the stream until a character that is not an ASCII
    /// alphabetic character (a-z, A-Z), or "_" is encountered or the end of the
    /// stream is reached.
    pub fn take_ascii_underscore_alphabetics(&mut self) -> Option<&'a str> {
        self.take_until_lambda(
            |character| !character.is_ascii_alphabetic() && character != '_',
            false,
        )
    }

    /// Takes a single character from the stream until a character that is not
    /// an ASCII alphabetic (a-z, A-Z), "_", is encountered or the end of the
    /// stream is reached.
    pub fn take_ascii_underscore_alphabetic(&mut self) -> Option<&'a str> {
        match self.take_count(1usize) {
            None => None,
            Some(character) => {
                let single_character = character.chars().next()?;
                if single_character.is_ascii_alphabetic() || single_character == '_' {
                    Some(character)
                } else {
                    None
                }
            }
        }
    }

    /// Takes characters from the stream until a characters that is not an ASCII
    /// alphanumeric character (a-z, A-Z, 0-9) is encountered or the end of the
    /// stream is reached.
    ///
    /// This method is useful for parsing identifiers, variable names, etc.
    ///
    /// Note: This method does not include non-ASCII alphanumeric characters.
    /// If you want to include non-ASCII alphanumeric characters, you should use
    /// the `take_until_lambda` method with a custom predicate.
    pub fn take_ascii_alphanumerics(&mut self) -> Option<&'a str> {
        self.take_until_lambda(|character| !character.is_ascii_alphanumeric(), false)
    }

    /// Takes characters from the stream until a character that is not an ASCII
    /// alphanumeric character (a-z, A-Z, 0-9) or "_" is encountered or the end
    /// of the stream is reached.
    ///
    /// This method is useful for parsing identifiers, variable names, etc.
    ///
    /// Note: This method does not include non-ASCII alphanumeric characters.
    /// If you want to include non-ASCII alphanumeric characters, you should use
    /// the `take_until_lambda` method with a custom predicate.
    pub fn take_ascii_underscore_alphanumerics(&mut self) -> Option<&'a str> {
        self.take_until_lambda(
            |character| !character.is_ascii_alphanumeric() && character != '_',
            false,
        )
    }

    /// Takes characters from the stream until a character that is not a digit (0-9)
    /// is encountered or the end of the stream is reached.
    ///
    /// This method is useful for parsing numeric literals, such as integers or
    /// floating-point numbers.
    ///
    /// Note: This method only considers ASCII digits (0-9), so it does not include
    /// characters that some languages use to format numbers such as: "x", ".",
    /// ",", " ", "_", or non-ASCII digits such as "६" (Devanagari 6), or "೬" (Kannada 6).
    pub fn take_ascii_digits(&mut self) -> Option<&'a str> {
        self.take_until_lambda(|character| !character.is_ascii_digit(), false)
    }

    /// Takes a single character from the stream if it is an ASCII digit (0-9) or
    /// `None` if the end-of-stream is found or the next character is not an
    /// ASCII digit.
    ///
    /// This method is useful for parsing numeric literals, such as integers or floating-point
    /// numbers.
    ///
    /// Note: This method only considers ASCII digits (0-9), so it does not include
    /// characters that some languages use to format numbers such as: "x", ".",
    /// ",", " ", "_", or non-ASCII digits such as "६" (Devanagari 6), or "೬" (Kannada 6).
    pub fn take_ascii_digit(&mut self) -> Option<&'a str> {
        match self.take_count(1usize) {
            None => None,
            Some(character) => {
                if character.chars().next()?.is_ascii_digit() {
                    Some(character)
                } else {
                    None
                }
            }
        }
    }

    /// Takes a single character from the stream if it is an ASCII punctuation character.
    ///
    /// This method checks the next character in the stream and if it is an ASCII punctuation
    /// character. it consumes that character and returns it as a `str`, on the other
    /// hand if the next character is not an ASCII punctuation character, it returns
    /// `None`.
    ///
    /// This method is useful for parsing operands in a text stream, such as
    /// commas, periods, exclamation marks, etc.
    ///
    /// Note: This method does not consume the character if it is not a punctuation
    /// character nor does it consume multiple punctuation characters.
    pub fn take_ascii_punctuation(&mut self) -> Option<&'a str> {
        self.peek(1).and_then(|peek| {
            if peek.chars().next()?.is_ascii_punctuation() {
                let result = &self.contents[self.offset..=self.offset];
                self.offset += 1;
                Some(result)
            } else {
                None
            }
        })
    }

    /// Takes characters from the stream until a linux/windows newline or the end of
    /// the stream is encountered.
    /// Returns a tuple containing the line content and the newline character(s)
    /// found or `None` if no newline was found.
    ///
    /// If the end of the stream is reached without finding a newline, it
    /// returns the remaining content and `None`.
    /// If the stream is empty, it returns `None`.
    ///
    pub fn take_until_newline(&mut self) -> Option<(&'a str, Option<&'a str>)> {
        let mut end_offset = self.offset;
        let content_len = self.contents.len();

        while end_offset < content_len {
            if self.contents[end_offset..].starts_with("\r\n") {
                let result = &self.contents[self.offset..end_offset];
                let newline = &self.contents[end_offset..end_offset + 2];
                self.offset = end_offset + newline.len();

                return Some((result, Some(newline)));
            } else if self.contents[end_offset..].starts_with('\n') {
                let result = &self.contents[self.offset..end_offset];
                let newline = &self.contents[end_offset..=end_offset];
                self.offset = end_offset + newline.len();

                return Some((result, Some(newline)));
            }

            // Advance by one UTF-8 character to respect character boundaries
            if let Some(ch) = self.contents[end_offset..].chars().next() {
                end_offset += ch.len_utf8();
            } else {
                break;
            }
        }

        if end_offset > self.offset {
            let result = &self.contents[self.offset..end_offset];
            self.offset = end_offset;
            return Some((result, None));
        }

        None
    }

    /// Creates a `Span` at the current position in the stream.
    ///
    /// The span will have length 1 and cover the current line.
    ///
    /// # Returns
    ///
    /// A `Span` representing the current position.
    #[must_use]
    pub fn span_here(&self) -> crate::errors::Span {
        crate::errors::Span {
            offset: u32::try_from(self.offset()).unwrap_or(0),
            line_start: u32::try_from(self.start_of_line()).unwrap_or(0),
            line_end: u32::try_from(self.end_of_line()).unwrap_or(0),
            length: 1,
        }
    }

    /// Creates a `Span` at the specified offset in the stream.
    ///
    /// The span will have length 1 and cover the line containing the offset.
    ///
    /// # Arguments
    ///
    /// * `offset` - The byte offset into the stream.
    ///
    /// # Returns
    ///
    /// A `Span` representing the specified position.
    #[must_use]
    pub fn span_at(&self, offset: usize) -> crate::errors::Span {
        crate::errors::Span {
            offset: u32::try_from(offset).unwrap_or(0),
            line_start: u32::try_from(self.start_of_line_from(offset)).unwrap_or(0),
            line_end: u32::try_from(self.end_of_line_from(offset)).unwrap_or(0),
            length: 1,
        }
    }

    /// Creates a `Span` covering a range of bytes.
    ///
    /// # Arguments
    ///
    /// * `start` - The starting byte offset.
    /// * `end` - The ending byte offset (exclusive).
    ///
    /// # Returns
    ///
    /// A `Span` representing the specified range.
    #[must_use]
    pub fn span_range(&self, start: usize, end: usize) -> crate::errors::Span {
        let length = end.saturating_sub(start);
        crate::errors::Span {
            offset: u32::try_from(start).unwrap_or(0),
            line_start: u32::try_from(self.start_of_line_from(start)).unwrap_or(0),
            line_end: u32::try_from(self.end_of_line_from(end.saturating_sub(1))).unwrap_or(0),
            length: u32::try_from(length).unwrap_or(0),
        }
    }

    /// Generates an `ErrorDetails` struct for the current offset in the stream
    /// with the provided `error_kind`.
    #[must_use]
    pub fn generate_error(&self, error_kind: crate::errors::ErrorKind) -> ErrorDetails<'a> {
        ErrorDetails {
            // Normally we would use usize for offsets, but VB6 was limited to 32-bit addressing.
            // Therefore, we safely cast to u32 here.
            error_offset: u32::try_from(self.offset()).unwrap_or(0),
            source_name: self.file_name.clone().into_boxed_str(),
            source_content: self.contents,
            line_end: u32::try_from(self.end_of_line()).unwrap_or(0),
            line_start: u32::try_from(self.start_of_line()).unwrap_or(0),
            kind: Box::new(error_kind),
            severity: crate::errors::Severity::Error,
            labels: vec![],
            notes: vec![],
        }
    }

    /// Generates an `ErrorDetails` struct for the specified `offset` in the stream
    /// with the provided `error_kind`.
    #[must_use]
    pub fn generate_error_at(
        &self,
        offset: usize,
        error_kind: crate::errors::ErrorKind,
    ) -> ErrorDetails<'a> {
        ErrorDetails {
            // Normally we would use usize for offsets, but VB6 was limited to 32-bit addressing.
            // Therefore, we safely cast to u32 here.
            error_offset: u32::try_from(offset).unwrap_or(0),
            source_name: self.file_name.clone().into_boxed_str(),
            source_content: self.contents,
            line_end: u32::try_from(self.end_of_line_from(offset)).unwrap_or(0),
            line_start: u32::try_from(self.start_of_line_from(offset)).unwrap_or(0),
            kind: Box::new(error_kind),
            severity: crate::errors::Severity::Error,
            labels: vec![],
            notes: vec![],
        }
    }

    /// Generates an `ErrorDetails` struct for the specified line start, offset,
    /// and line end in the stream with the provided `error_kind`.
    ///
    /// The method ensures that the provided offsets are in the correct order
    /// and adjusts them if necessary. If the `line_end` exceeds the length of
    /// the contents, it is set to the length of the contents.
    #[must_use]
    pub fn generate_bounded_error_at(
        &self,
        line_start: usize,
        offset: usize,
        line_end: usize,
        error_kind: crate::errors::ErrorKind,
    ) -> ErrorDetails<'a> {
        let mut offsets = [line_start, offset, line_end];
        // Used unstable sort for performance since order of usize primitives is identical to stable sort.
        offsets.sort_unstable();

        if offsets[2] > self.contents.len() {
            offsets[2] = self.contents.len();
        }

        ErrorDetails {
            source_name: self.file_name.clone().into_boxed_str(),
            source_content: self.contents,
            // Normally we would use usize for offsets, but VB6 was limited to 32-bit addressing.
            // Therefore, we safely cast to u32 here.
            line_start: u32::try_from(offsets[0]).unwrap_or(0),
            error_offset: u32::try_from(offsets[1]).unwrap_or(0),
            line_end: u32::try_from(offsets[2]).unwrap_or(0),
            kind: Box::new(error_kind),
            severity: crate::errors::Severity::Error,
            labels: vec![],
            notes: vec![],
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn take_case_sensitive() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take("Hello", Comparator::CaseSensitive),
            Some("Hello")
        );
        assert_eq!(stream.peek(1), Some(","));
    }

    #[test]
    fn take_case_insensitive() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take("hello", Comparator::CaseInsensitive),
            Some("Hello")
        );
        assert_eq!(stream.peek(1), Some(","));
    }

    #[test]
    fn take_no_match() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take("Goodbye", Comparator::CaseSensitive), None);
        assert_eq!(stream.peek(1), Some("H"));
    }

    #[test]
    fn take_no_match_case_insensitive() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take("goodbye", Comparator::CaseInsensitive), None);
        assert_eq!(stream.peek(1), Some("H"));
    }

    #[test]
    fn forward() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        stream.forward(7);
        assert_eq!(stream.peek(1), Some("W"));
        stream.forward(6);
        assert_eq!(stream.peek(1), None);
    }

    #[test]
    fn forward_out_of_bounds() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        stream.forward(20); // Forwarding beyond the length of the contents
        assert_eq!(stream.peek(1), None);
        assert_eq!(stream.offset(), contents.len());
    }

    #[test]
    fn peek() {
        let contents = "Hello, World!";
        let stream = SourceStream::new("test.txt", contents);
        assert_eq!(stream.peek(5), Some("Hello"));
        assert_eq!(stream.peek(20), None); // Peek beyond the length of the contents
    }

    #[test]
    fn take_until() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until(", ", Comparator::CaseSensitive),
            Some(("Hello", ", "))
        );
        assert_eq!(stream.peek(1), Some(","));
    }

    #[test]
    fn take_until_case_insensitive() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until("world!", Comparator::CaseInsensitive),
            Some(("Hello, ", "World!"))
        );
        assert_eq!(stream.peek(1), Some("W"));
    }

    #[test]
    fn take_until_no_match() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until("Goodbye", Comparator::CaseSensitive),
            None
        );
        assert_eq!(stream.peek(1), Some("H"));
    }

    #[test]
    fn peek_windows_newline() {
        let contents = "Hello, World!\r\nThis is a test.\n";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.peek_newline(), None);
        stream.forward(13); // Move past "Hello, World!"
        assert_eq!(stream.peek_newline(), Some("\r\n"));
        assert_eq!(stream.peek_newline(), Some("\r\n"));
        stream.forward(2); // Move past the "\r\n"
        assert_eq!(stream.peek_newline(), None);
        let _ = stream.take("This is a test.", Comparator::CaseSensitive); // Move past the "This is a test."
        assert_eq!(stream.peek_newline(), Some("\n"));
    }

    #[test]
    fn peek_newline_no_newline() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.peek_newline(), None);
        let _ = stream.take("Hello, World! This is a test.", Comparator::CaseSensitive);
        assert_eq!(stream.peek_newline(), None);
        assert_eq!(stream.peek(1), None); // No more content to peek
    }

    #[test]
    fn peek_linux_newline() {
        let contents_with_crlf = "Hello, World!\r\nThis is a test.";
        let mut stream_with_crlf = SourceStream::new("test.txt", contents_with_crlf);

        assert_eq!(stream_with_crlf.peek_linux_newline(), None);
        stream_with_crlf.forward(13); // Move past "Hello, World!"
        assert_eq!(stream_with_crlf.peek_linux_newline(), None);
        stream_with_crlf.forward(1); // Move past the carriage return
        assert_eq!(stream_with_crlf.peek_linux_newline(), Some("\n"));
        assert_eq!(stream_with_crlf.peek_linux_newline(), Some("\n"));
        stream_with_crlf.forward(1); // Move past the newline
        assert_eq!(stream_with_crlf.peek_linux_newline(), None);
        assert_eq!(stream_with_crlf.peek_linux_newline(), None);
        let _ = stream_with_crlf.take("This is a test.", Comparator::CaseSensitive); // Move past the "This is a test."
        assert_eq!(stream_with_crlf.peek_linux_newline(), None);
    }

    #[test]
    fn take_line_or_eos() {
        let contents = "Hello, World!\nThis is a test.\n";
        let mut stream = SourceStream::new("test.txt", contents);

        assert!(!stream.is_empty());
        assert_eq!(
            stream.take_until_newline(),
            Some(("Hello, World!", Some("\n")))
        );
        assert!(!stream.is_empty());
        assert_eq!(
            stream.take_until_newline(),
            Some(("This is a test.", Some("\n")))
        );
        assert!(stream.is_empty());
        assert_eq!(stream.take_until_newline(), None);
        assert!(stream.is_empty());
    }

    #[test]
    fn take_line_is_empty() {
        let contents = "Words\r\nAre hard.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert!(!stream.is_empty());
        assert_eq!(stream.take_until_newline(), Some(("Words", Some("\r\n"))));
        assert!(!stream.is_empty());
        assert_eq!(stream.take_until_newline(), Some(("Are hard.", None)));
        assert!(stream.is_empty());
        assert_eq!(stream.take_until_newline(), None);
        assert!(stream.is_empty());
    }

    #[test]
    fn take_line_no_newline() {
        let contents = "Words Are hard.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert!(!stream.is_empty());
        assert_eq!(stream.take_until_newline(), Some(("Words Are hard.", None)));
        assert!(stream.is_empty());
        assert_eq!(stream.take_until_newline(), None);
        assert!(stream.is_empty());
    }

    #[test]
    fn take_whitespaces() {
        let contents = "   Hello, World!   \r\n";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_ascii_whitespaces(), Some("   "));
        assert_eq!(stream.peek(1), Some("H"));
        assert_eq!(stream.take_ascii_whitespaces(), None);
        assert_eq!(
            stream.take("Hello, World!", Comparator::CaseSensitive),
            Some("Hello, World!")
        );
        assert!(!stream.is_empty());
        assert_eq!(stream.take_ascii_whitespaces(), Some("   "));
        assert!(!stream.is_empty());
        assert_eq!(stream.take_ascii_whitespaces(), None);
        assert_eq!(stream.peek(1), Some("\r"));
        assert_eq!(stream.take("\r\n", Comparator::CaseSensitive), Some("\r\n"));
        assert!(stream.is_empty());
        assert_eq!(stream.take_ascii_whitespaces(), None);
    }

    #[test]
    fn take_alphabetics() {
        let contents = "Hello123 World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_ascii_alphabetics(), Some("Hello"));
        assert_eq!(stream.peek(1), Some("1"));
        assert_eq!(stream.take_ascii_alphabetics(), None);
        assert_eq!(
            stream.take("123 World!", Comparator::CaseSensitive),
            Some("123 World!")
        );
        assert!(stream.is_empty());
    }

    #[test]
    fn take_alphanumerics() {
        let contents = "Hello123 World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_ascii_alphanumerics(), Some("Hello123"));
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take_ascii_alphanumerics(), None);
        assert_eq!(
            stream.take(" World!", Comparator::CaseSensitive),
            Some(" World!")
        );
        assert!(stream.is_empty());
    }

    #[test]
    fn take_digits() {
        let contents = "12345abcde";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_ascii_digits(), Some("12345"));
        assert_eq!(stream.peek(1), Some("a"));
        assert_eq!(stream.take_ascii_digits(), None);
        assert_eq!(
            stream.take("abcde", Comparator::CaseSensitive),
            Some("abcde")
        );
        assert!(stream.is_empty());
    }

    #[test]
    fn take_punctuation() {
        let contents = "Hello, World!! How are you?";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_ascii_punctuation(), None);
        assert_eq!(
            stream.take("Hello", Comparator::CaseSensitive),
            Some("Hello")
        );
        assert_eq!(stream.take_ascii_punctuation(), Some(","));
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take_ascii_punctuation(), None);
        assert_eq!(
            stream.take(" World", Comparator::CaseSensitive),
            Some(" World")
        );
        assert_eq!(stream.peek(1), Some("!"));
        assert_eq!(stream.take_ascii_punctuation(), Some("!"));
        assert!(!stream.is_empty());
        assert_eq!(stream.peek(1), Some("!"));
        assert_eq!(stream.take_ascii_punctuation(), Some("!"));
        assert!(!stream.is_empty());
        assert_eq!(
            stream.take(" How are you", Comparator::CaseSensitive),
            Some(" How are you")
        );
        assert_eq!(stream.peek(1), Some("?"));
        assert_eq!(stream.take_ascii_punctuation(), Some("?"));
        assert!(stream.is_empty());
        assert_eq!(stream.take_ascii_punctuation(), None);
        assert_eq!(stream.peek(1), None);
    }

    #[test]
    fn take_until_lambda() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until_lambda(|character| character == ' ', false),
            Some("Hello,")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == ' ', false),
            Some("World!")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == ' ', false),
            Some("This")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == ' ', false),
            Some("is")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == ' ', false),
            Some("a")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == '.', false),
            Some("test")
        );
        assert!(!stream.is_empty());
        assert_eq!(stream.peek(1), Some("."));
        assert_eq!(stream.take(".", Comparator::CaseSensitive), Some("."));
        assert!(stream.is_empty());
        assert_eq!(
            stream.take_until_lambda(|character| character == ' ', false),
            None
        );
        assert_eq!(stream.peek(1), None);
        assert!(stream.is_empty());
    }
}
