/// A structure representing a stream of bytes from a source file.
/// It holds the file name, the contents of the file as a `BStr`, and an offset
/// indicating the current position in the stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceStream<'a> {
    pub file_name: String,
    pub contents: &'a str,
    pub offset: usize,
}

/// An enum representing the type of comparison to be used when taking bytes
/// from the `SourceStream`.
/// It can be either case-sensitive or case-insensitive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Comparator {
    CaseSensitive,
    CaseInsensitive,
}

pub enum SourceStreamError {
    EmptyContents,
    MalformedContents,
}

impl<'a> SourceStream<'a> {
    /// Creates a new `SourceStream` with the given file name and contents.
    ///
    /// The `file_name` is a `String` representing the name of the file being parsed.
    /// The `contents` is a `str` that contains the contents of the stream.
    pub fn new<S: Into<String>>(file_name: S, contents: &'a str) -> Self {
        Self {
            file_name: file_name.into(),
            contents: contents,
            offset: 0,
        }
    }

    /// Moves the offset forward by `count` characters in the stream.
    pub fn forward(&mut self, count: usize) {
        let end_offset = self.offset + count;

        if end_offset > self.contents.len() {
            self.offset = self.contents.len();
        } else {
            self.offset = end_offset;
        }
    }

    pub fn forward_to_next_line(&mut self) {
        let _ = self.take_until_newline();
        self.take_newline();
    }

    /// Returns the file name of the stream.
    pub fn file_name(&self) -> &str {
        &self.file_name
    }

    /// Returns the current offset in the stream.
    pub fn offset(&self) -> usize {
        self.offset
    }

    pub fn start_of_line(&self) -> usize {
        // Find the last newline character before the current offset
        if let Some(pos) = self.contents[..self.offset].rfind('\n') {
            pos + 1 // Return the position after the newline character
        } else {
            0 // If no newline found, return the start of the stream
        }
    }

    pub fn end_of_line(&self) -> usize {
        // Find the next newline character after the current offset
        if let Some(pos) = self.contents[self.offset..].find('\n') {
            self.offset + pos // Return the position of the newline character
        } else {
            self.contents.len() // If no newline found, return the end of the stream
        }
    }

    /// Checks if the stream is empty, meaning the offset is at or beyond the
    /// end of the contents.
    pub fn is_empty(&self) -> bool {
        self.offset >= self.contents.len()
    }

    /// Peeks at the next `count` bytes in the stream without consuming them.
    pub fn peek(&self, count: usize) -> Option<&'a str> {
        let end_offset = self.offset + count;

        if end_offset > self.contents.len() {
            None
        } else {
            Some(&self.contents[self.offset..end_offset])
        }
    }

    /// Peeks at the next bytes in the stream to see if they match the `compare` value.
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
            Comparator::CaseInsensitive => peek_slice.eq_ignore_ascii_case(&compare),
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
    pub fn peek_newline(&self) -> Option<&'a str> {
        self.peek_windows_newline()
            .or_else(|| self.peek_linux_newline())
    }

    /// Takes characters from the stream if they match the `compare` str.
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
            Comparator::CaseInsensitive => compare_slice.eq_ignore_ascii_case(&compare),
        };

        if !matches {
            return None;
        }

        end_offset += compare_len;
        let result = &self.contents[self.offset..end_offset];
        self.offset = end_offset;
        Some(result)
    }

    /// Takes characters from the stream until a character that matches the
    /// compare `str` is encountered or the end of the stream is reached.
    ///
    /// If a match is found, it returns a tuple containing an `str` for the
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
                Comparator::CaseInsensitive => slice.eq_ignore_ascii_case(&compare),
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
                return Some((result, &self.contents[end_offset..end_offset + 1]));
            }
        }

        if end_offset > self.offset {
            let result = &self.contents[self.offset..end_offset];
            self.offset = end_offset;
            Some((result, &self.contents[end_offset..end_offset + 1]))
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
    /// function is encountered or the end of the stream is reached.
    ///
    /// This method is useful for parsing various types of content where you need
    /// to consume characters until a specific condition is met, such as whitespace,
    /// alphabetic characters, alphanumeric characters, digits, punctuation, etc.
    pub fn take_until_lambda(
        &mut self,
        mut predicate: impl FnMut(char) -> bool,
    ) -> Option<&'a str> {
        let mut end_offset = self.offset;
        let content_len = self.contents.len();

        while end_offset < content_len {
            let current_char = self.contents[end_offset..].chars().next()?;
            if predicate(current_char) {
                let result = &self.contents[self.offset..end_offset];
                self.offset = end_offset;

                if result.len() == 0 {
                    return None;
                } else {
                    return Some(result);
                }
            }
            end_offset += current_char.len_utf8();
        }

        None
    }

    /// Takes characters from the stream until a character that is not a whitespace
    /// (including carriage return and newline) is encountered or the end of the
    /// stream is reached.
    pub fn take_whitespaces(&mut self) -> Option<&'a str> {
        self.take_until_lambda(|character| {
            !character.is_ascii_whitespace() || character == '\r' || character == '\n'
        })
    }

    /// Takes characters from the stream until a character that is not an alphabetic
    /// character (a-z, A-Z) is encountered or the end of the stream is reached.
    pub fn take_alphabetics(&mut self) -> Option<&'a str> {
        self.take_until_lambda(|byte| !byte.is_ascii_alphabetic())
    }

    /// Takes characters from the stream until a characters that is not an
    /// alphanumeric character (a-z, A-Z, 0-9) is encountered or the end of the
    /// stream is reached.
    ///
    /// This method is useful for parsing identifiers, variable names, etc.
    ///
    /// Note: This method does not include non-ASCII alphanumeric characters.
    /// If you want to include non-ASCII alphanumeric characters, you should use
    /// the `take_until_lambda` method with a custom predicate.
    pub fn take_alphanumerics(&mut self) -> Option<&'a str> {
        self.take_until_lambda(|byte| !byte.is_ascii_alphanumeric())
    }

    /// Takes characters from the stream until a character that is not a digit (0-9) is encountered
    /// or the end of the stream is reached.
    ///
    /// This method is useful for parsing numeric literals, such as integers or floating-point
    /// numbers.
    ///
    /// Note: This method only considers ASCII digits (0-9), so it does not include
    /// non-ASCII digits, ".", ",", " ", or "_" characters which may be needed for
    /// parsing numeric literals in some languages.
    pub fn take_digits(&mut self) -> Option<&'a str> {
        self.take_until_lambda(|character| !character.is_ascii_digit())
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
    pub fn take_punctuation(&mut self) -> Option<&'a str> {
        self.peek(1).and_then(|peek| {
            if peek.chars().next()?.is_ascii_punctuation() {
                let result = &self.contents[self.offset..self.offset + 1];
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
            } else if self.contents[end_offset..].starts_with("\n") {
                let result = &self.contents[self.offset..end_offset];
                let newline = &self.contents[end_offset..end_offset + 1];
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

        assert_eq!(stream.is_empty(), false);
        assert_eq!(
            stream.take_until_newline(),
            Some(("Hello, World!", Some("\n")))
        );
        assert_eq!(stream.is_empty(), false);
        assert_eq!(
            stream.take_until_newline(),
            Some(("This is a test.", Some("\n")))
        );
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_until_newline(), None);
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_line_is_empty() {
        let contents = "Words\r\nAre hard.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.take_until_newline(), Some(("Words", Some("\r\n"))));
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.take_until_newline(), Some(("Are hard.", None)));
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_until_newline(), None);
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_line_no_newline() {
        let contents = "Words Are hard.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.take_until_newline(), Some(("Words Are hard.", None)));
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_until_newline(), None);
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_whitespaces() {
        let contents = "   Hello, World!   \r\n";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_whitespaces(), Some("   "));
        assert_eq!(stream.peek(1), Some("H"));
        assert_eq!(stream.take_whitespaces(), None);
        assert_eq!(
            stream.take("Hello, World!", Comparator::CaseSensitive),
            Some("Hello, World!")
        );
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.take_whitespaces(), Some("   "));
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.take_whitespaces(), None);
        assert_eq!(stream.peek(1), Some("\r"));
        assert_eq!(stream.take("\r\n", Comparator::CaseSensitive), Some("\r\n"));
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_whitespaces(), None);
    }

    #[test]
    fn take_alphabetics() {
        let contents = "Hello123 World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_alphabetics(), Some("Hello"));
        assert_eq!(stream.peek(1), Some("1"));
        assert_eq!(stream.take_alphabetics(), None);
        assert_eq!(
            stream.take("123 World!", Comparator::CaseSensitive),
            Some("123 World!")
        );
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_alphanumerics() {
        let contents = "Hello123 World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_alphanumerics(), Some("Hello123"));
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take_alphanumerics(), None);
        assert_eq!(
            stream.take(" World!", Comparator::CaseSensitive),
            Some(" World!")
        );
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_digits() {
        let contents = "12345abcde";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_digits(), Some("12345"));
        assert_eq!(stream.peek(1), Some("a"));
        assert_eq!(stream.take_digits(), None);
        assert_eq!(
            stream.take("abcde", Comparator::CaseSensitive),
            Some("abcde")
        );
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_punctuation() {
        let contents = "Hello, World!! How are you?";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_punctuation(), None);
        assert_eq!(
            stream.take("Hello", Comparator::CaseSensitive),
            Some("Hello")
        );
        assert_eq!(stream.take_punctuation(), Some(","));
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take_punctuation(), None);
        assert_eq!(
            stream.take(" World", Comparator::CaseSensitive),
            Some(" World")
        );
        assert_eq!(stream.peek(1), Some("!"));
        assert_eq!(stream.take_punctuation(), Some("!"));
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.peek(1), Some("!"));
        assert_eq!(stream.take_punctuation(), Some("!"));
        assert_eq!(stream.is_empty(), false);
        assert_eq!(
            stream.take(" How are you", Comparator::CaseSensitive),
            Some(" How are you")
        );
        assert_eq!(stream.peek(1), Some("?"));
        assert_eq!(stream.take_punctuation(), Some("?"));
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_punctuation(), None);
        assert_eq!(stream.peek(1), None);
    }

    #[test]
    fn take_until_lambda() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until_lambda(|character| character == ' '),
            Some("Hello,")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == ' '),
            Some("World!")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == ' '),
            Some("This")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == ' '),
            Some("is")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == ' '),
            Some("a")
        );
        assert_eq!(stream.peek(1), Some(" "));
        assert_eq!(stream.take(" ", Comparator::CaseSensitive), Some(" "));
        assert_eq!(
            stream.take_until_lambda(|character| character == '.'),
            Some("test")
        );
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.peek(1), Some("."));
        assert_eq!(stream.take(".", Comparator::CaseSensitive), Some("."));
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_until_lambda(|character| character == ' '), None);
        assert_eq!(stream.peek(1), None);
        assert_eq!(stream.is_empty(), true);
    }
}
