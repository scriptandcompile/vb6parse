use bstr::{BStr, ByteSlice};

/// A structure representing a stream of bytes from a source file.
/// It holds the file name, the contents of the file as a `BStr`, and an offset
/// indicating the current position in the stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SourceStream<'a> {
    pub file_name: String,
    pub contents: &'a BStr,
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

impl<'a> SourceStream<'a> {
    /// Creates a new `SourceStream` with the given file name and contents.
    ///
    /// The `file_name` is a string representing the name of the file being parsed.
    /// The `contents` is a reference to a `BStr` containing the bytes of the file.
    pub fn new(file_name: impl Into<String>, contents: impl Into<&'a BStr>) -> Self {
        Self {
            file_name: file_name.into(),
            contents: contents.into(),
            offset: 0,
        }
    }

    /// Moves the offset forward by `count` bytes in the stream.
    pub fn forward(&mut self, count: usize) {
        let end_offset = self.offset + count;

        if end_offset > self.contents.len() {
            self.offset = self.contents.len();
        } else {
            self.offset = end_offset;
        }
    }

    /// Returns the file name of the stream.
    pub fn file_name(&self) -> &str {
        &self.file_name
    }

    /// Returns the current offset in the stream.
    pub fn offset(&self) -> usize {
        self.offset
    }

    /// Checks if the stream is empty, meaning the offset is at or beyond the
    /// end of the contents.
    pub fn is_empty(&self) -> bool {
        self.offset >= self.contents.len()
    }

    /// Peeks at the next `count` bytes in the stream without consuming them.
    pub fn peek(&self, count: usize) -> Option<&'a BStr> {
        let end_offset = self.offset + count;

        if end_offset > self.contents.len() {
            None
        } else {
            Some(&self.contents[self.offset..end_offset])
        }
    }

    /// Peeks at the next bytes in the stream to see if they match the `compare` BStr.
    ///
    /// If they match, it returns the bytes as a `BStr`. If they do not match, it returns `None`.
    ///
    /// The `case_sensitive` parameter determines whether the comparison is case-sensitive
    /// or case-insensitive.
    pub fn peek_text<'b>(
        &self,
        compare: impl Into<&'b BStr>,
        case_sensitive: Comparator,
    ) -> Option<&'a BStr> {
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

    /// Peeks at the next bytes in the stream to see if they match a linux newline
    /// character (`\n`). If a linux newline character is found, it returns it
    /// as a `BStr`. If no linux newline character is found, it returns `None`.
    pub fn peek_linux_newline(&self) -> Option<&'a BStr> {
        let peek_len_1 = self.peek(1);
        if peek_len_1 == Some(b"\n".as_bstr()) {
            return peek_len_1;
        }

        None
    }

    /// Peeks at the next bytes in the stream to see if they match a windows newline
    /// character (`\r\n`). If a windows newline character is found, it returns it
    /// as a `BStr`. If no windows newline character is found, it returns `None`.
    pub fn peek_windows_newline(&self) -> Option<&'a BStr> {
        let peek_len_2 = self.peek(2);
        if peek_len_2 == Some(b"\r\n".as_bstr()) {
            return peek_len_2;
        }

        None
    }

    /// Peeks at the next bytes in the stream to see if they match a newline character
    /// (either windows `\r\n` or linux `\n`). If a newline character is found,
    /// it returns it as a `BStr`. If no newline character is found, it returns `None`.
    pub fn peek_newline(&self) -> Option<&'a BStr> {
        self.peek_windows_newline()
            .or_else(|| self.peek_linux_newline())
    }

    /// Takes bytes from the stream if they match the `compare` BStr.
    pub fn take<'b>(
        &mut self,
        compare: impl Into<&'b BStr>,
        case_sensitive: Comparator,
    ) -> Option<&'a BStr> {
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

    /// Takes bytes from the stream until a byte that matches the `compare` BStr
    /// is encountered or the end of the stream is reached.
    pub fn take_until<'b>(
        &mut self,
        compare: impl Into<&'b BStr>,
        case_sensitive: Comparator,
    ) -> Option<(&'a BStr, &'a BStr)> {
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

    /// Takes a newline character (either windows `\r\n` or linux `\n`) from the stream
    /// if it exists. If a newline character is found, it consumes it and returns
    /// it as a `BStr`. If no newline character is found, it returns `None`.
    pub fn take_newline(&mut self) -> Option<&'a BStr> {
        self.take_windows_newline()
            .or_else(|| self.take_linux_newline())
    }

    /// Takes a windows newline character (`\r\n`) from the stream if it exists.
    pub fn take_windows_newline(&mut self) -> Option<&'a BStr> {
        self.take(b"\r\n".as_bstr(), Comparator::CaseSensitive)
    }

    /// Takes a linux newline character (`\n`) from the stream if it exists.
    pub fn take_linux_newline(&mut self) -> Option<&'a BStr> {
        self.take(b"\n".as_bstr(), Comparator::CaseSensitive)
    }

    /// Takes characters from the stream until a byte that matches the predicate
    /// function is encountered or the end of the stream is reached.
    ///
    /// This method is useful for parsing various types of content where you need
    /// to consume bytes until a specific condition is met, such as whitespace,
    /// alphabetic characters, alphanumeric characters, digits, punctuation, etc.
    pub fn take_until_lambda(&mut self, mut predicate: impl FnMut(u8) -> bool) -> Option<&'a BStr> {
        let mut end_offset = self.offset;
        let content_len = self.contents.len();

        while end_offset < content_len {
            if predicate(self.contents[end_offset]) {
                break;
            }
            end_offset += 1;
        }

        if end_offset > self.offset {
            let result = &self.contents[self.offset..end_offset];
            self.offset = end_offset;
            Some(result)
        } else {
            None
        }
    }

    /// Takes bytes from the stream until a byte that is not a whitespace character
    /// (including carriage return and newline) is encountered or the end of the
    /// stream is reached.
    pub fn take_whitespaces(&mut self) -> Option<&'a BStr> {
        self.take_until_lambda(|byte| !byte.is_ascii_whitespace() || byte == b'\r' || byte == b'\n')
    }

    /// Takes bytes from the stream until a byte that is not an alphabetic character
    /// (a-z, A-Z) is encountered or the end of the stream is reached.
    pub fn take_alphabetics(&mut self) -> Option<&'a BStr> {
        self.take_until_lambda(|byte| !byte.is_ascii_alphabetic())
    }

    /// Takes bytes from the stream until a byte that is not an alphanumeric character
    /// (a-z, A-Z, 0-9) is encountered or the end of the stream is reached.
    ///
    /// This method is useful for parsing identifiers, variable names, etc.
    ///
    /// Note: This method does not include non-ASCII alphanumeric characters.
    /// If you want to include non-ASCII alphanumeric characters, you can modify the
    /// predicate function to include them.
    pub fn take_alphanumerics(&mut self) -> Option<&'a BStr> {
        self.take_until_lambda(|byte| !byte.is_ascii_alphanumeric())
    }

    /// Takes bytes from the stream until a byte that is not a digit (0-9) is encountered
    /// or the end of the stream is reached.
    ///
    /// This method is useful for parsing numeric literals, such as integers or floating-point
    /// numbers.
    ///
    /// Note: This method only considers ASCII digits (0-9), so it does not include
    /// non-ASCII digits, ".", ",", " ", or "_" characters which may be needed for
    /// parsing numeric literals in some languages.
    pub fn take_digits(&mut self) -> Option<&'a BStr> {
        self.take_until_lambda(|byte| !byte.is_ascii_digit())
    }

    /// Takes a single byte from the stream if it is an ASCII punctuation character.
    ///
    /// This method checks the next byte in the stream and if it is an ASCII punctuation
    /// character. it consumes that byte and returns it as a `BStr`, on the other
    /// hand if the next byte is not an ASCII punctuation character, it returns
    /// `None`.
    ///
    /// This method is useful for parsing operands in a text stream, such as
    /// commas, periods, exclamation marks, etc.
    ///
    /// Note: This method does not consume the byte if it is not a punctuation
    /// character nor does it consume multiple punctuation characters.
    pub fn take_punctuation(&mut self) -> Option<&'a BStr> {
        self.peek(1).and_then(|peek| {
            if peek[0].is_ascii_punctuation() {
                let result = peek;
                self.forward(1);
                Some(result)
            } else {
                None
            }
        })
    }

    /// Takes bytes from the stream until a linux/windows newline or the end of
    /// the stream is encountered.
    /// Returns a tuple containing the line content and the newline character(s)
    /// found or `None` if no newline was found.
    ///
    /// If the end of the stream is reached without finding a newline, it
    /// returns the remaining content and `None`.
    /// If the stream is empty, it returns `None`.
    ///
    pub fn take_until_newline(&mut self) -> Option<(&'a BStr, Option<&'a BStr>)> {
        let mut end_offset = self.offset;
        let content_len = self.contents.len();

        while end_offset < content_len {
            if self.contents[end_offset..].starts_with(b"\r\n") {
                let result = &self.contents[self.offset..end_offset];
                let newline = self.contents[end_offset..end_offset + 2].as_bstr();
                self.offset = end_offset + newline.len();

                return Some((result, Some(newline)));
            } else if self.contents[end_offset..].starts_with(b"\n") {
                let result = &self.contents[self.offset..end_offset];
                let newline = self.contents[end_offset..end_offset + 1].as_bstr();
                self.offset = end_offset + newline.len();

                return Some((result, Some(newline)));
            }

            end_offset += 1;
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
    //use bstr::BStr;
    use bstr::ByteSlice;

    #[test]
    fn take_case_sensitive() {
        let contents = b"Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take("Hello", Comparator::CaseSensitive),
            Some(b"Hello".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b",".as_bstr()));
    }

    #[test]
    fn take_case_insensitive() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take("hello", Comparator::CaseInsensitive),
            Some(b"Hello".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b",".as_bstr()));
    }

    #[test]
    fn take_no_match() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take("Goodbye", Comparator::CaseSensitive), None);
        assert_eq!(stream.peek(1), Some(b"H".as_bstr()));
    }

    #[test]
    fn take_no_match_case_insensitive() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take("goodbye", Comparator::CaseInsensitive), None);
        assert_eq!(stream.peek(1), Some(b"H".as_bstr()));
    }

    #[test]
    fn forward() {
        let contents = "Hello, World!";
        let mut stream = SourceStream::new("test.txt", contents);

        stream.forward(7);
        assert_eq!(stream.peek(1), Some(b"W".as_bstr()));
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
        assert_eq!(stream.peek(5), Some(b"Hello".as_bstr()));
        assert_eq!(stream.peek(20), None); // Peek beyond the length of the contents
    }

    #[test]
    fn take_until() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until(", ", Comparator::CaseSensitive),
            Some((b"Hello".as_bstr(), b", ".as_bstr()))
        );
        assert_eq!(stream.peek(1), Some(b",".as_bstr()));
    }

    #[test]
    fn take_until_case_insensitive() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until("world!", Comparator::CaseInsensitive),
            Some((b"Hello, ".as_bstr(), b"World!".as_bstr()))
        );
        assert_eq!(stream.peek(1), Some(b"W".as_bstr()));
    }

    #[test]
    fn take_until_no_match() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until("Goodbye", Comparator::CaseSensitive),
            None
        );
        assert_eq!(stream.peek(1), Some(b"H".as_bstr()));
    }

    #[test]
    fn peek_windows_newline() {
        let contents = "Hello, World!\r\nThis is a test.\n";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.peek_newline(), None);
        stream.forward(13); // Move past "Hello, World!"
        assert_eq!(stream.peek_newline(), Some(b"\r\n".as_bstr()));
        assert_eq!(stream.peek_newline(), Some(b"\r\n".as_bstr()));
        stream.forward(2); // Move past the "\r\n"
        assert_eq!(stream.peek_newline(), None);
        let _ = stream.take(b"This is a test.".as_bstr(), Comparator::CaseSensitive); // Move past the "This is a test."
        assert_eq!(stream.peek_newline(), Some(b"\n".as_bstr()));
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
        assert_eq!(stream_with_crlf.peek_linux_newline(), Some(b"\n".as_bstr()));
        assert_eq!(stream_with_crlf.peek_linux_newline(), Some(b"\n".as_bstr()));
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
            Some((b"Hello, World!".as_bstr(), Some(b"\n".as_bstr())))
        );
        assert_eq!(stream.is_empty(), false);
        assert_eq!(
            stream.take_until_newline(),
            Some((b"This is a test.".as_bstr(), Some(b"\n".as_bstr())))
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
        assert_eq!(
            stream.take_until_newline(),
            Some((b"Words".as_bstr(), Some(b"\r\n".as_bstr())))
        );
        assert_eq!(stream.is_empty(), false);
        assert_eq!(
            stream.take_until_newline(),
            Some((b"Are hard.".as_bstr(), None))
        );
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_until_newline(), None);
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_line_no_newline() {
        let contents = "Words Are hard.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.is_empty(), false);
        assert_eq!(
            stream.take_until_newline(),
            Some((b"Words Are hard.".as_bstr(), None))
        );
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_until_newline(), None);
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_whitespaces() {
        let contents = "   Hello, World!   \r\n";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_whitespaces(), Some(b"   ".as_bstr()));
        assert_eq!(stream.peek(1), Some(b"H".as_bstr()));
        assert_eq!(stream.take_whitespaces(), None);
        assert_eq!(
            stream.take("Hello, World!", Comparator::CaseSensitive),
            Some(b"Hello, World!".as_bstr())
        );
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.take_whitespaces(), Some(b"   ".as_bstr()));
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.take_whitespaces(), None);
        assert_eq!(stream.peek(1), Some(b"\r".as_bstr()));
        assert_eq!(
            stream.take("\r\n", Comparator::CaseSensitive),
            Some(b"\r\n".as_bstr())
        );
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_whitespaces(), None);
    }

    #[test]
    fn take_alphabetics() {
        let contents = "Hello123 World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_alphabetics(), Some(b"Hello".as_bstr()));
        assert_eq!(stream.peek(1), Some(b"1".as_bstr()));
        assert_eq!(stream.take_alphabetics(), None);
        assert_eq!(
            stream.take("123 World!", Comparator::CaseSensitive),
            Some(b"123 World!".as_bstr())
        );
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_alphanumerics() {
        let contents = "Hello123 World!";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_alphanumerics(), Some(b"Hello123".as_bstr()));
        assert_eq!(stream.peek(1), Some(b" ".as_bstr()));
        assert_eq!(stream.take_alphanumerics(), None);
        assert_eq!(
            stream.take(" World!", Comparator::CaseSensitive),
            Some(b" World!".as_bstr())
        );
        assert_eq!(stream.is_empty(), true);
    }

    #[test]
    fn take_digits() {
        let contents = "12345abcde";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(stream.take_digits(), Some(b"12345".as_bstr()));
        assert_eq!(stream.peek(1), Some(b"a".as_bstr()));
        assert_eq!(stream.take_digits(), None);
        assert_eq!(
            stream.take("abcde", Comparator::CaseSensitive),
            Some(b"abcde".as_bstr())
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
            Some(b"Hello".as_bstr())
        );
        assert_eq!(stream.take_punctuation(), Some(b",".as_bstr()));
        assert_eq!(stream.peek(1), Some(b" ".as_bstr()));
        assert_eq!(stream.take_punctuation(), None);
        assert_eq!(
            stream.take(" World", Comparator::CaseSensitive),
            Some(b" World".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b"!".as_bstr()));
        assert_eq!(stream.take_punctuation(), Some(b"!".as_bstr()));
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.peek(1), Some(b"!".as_bstr()));
        assert_eq!(stream.take_punctuation(), Some(b"!".as_bstr()));
        assert_eq!(stream.is_empty(), false);
        assert_eq!(
            stream.take(b" How are you".as_bstr(), Comparator::CaseSensitive),
            Some(b" How are you".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b"?".as_bstr()));
        assert_eq!(stream.take_punctuation(), Some(b"?".as_bstr()));
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_punctuation(), None);
        assert_eq!(stream.peek(1), None);
    }

    #[test]
    fn take_until_lambda() {
        let contents = "Hello, World! This is a test.";
        let mut stream = SourceStream::new("test.txt", contents);

        assert_eq!(
            stream.take_until_lambda(|byte| byte == b' '),
            Some(b"Hello,".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b" ".as_bstr()));
        assert_eq!(
            stream.take(" ", Comparator::CaseSensitive),
            Some(b" ".as_bstr())
        );
        assert_eq!(
            stream.take_until_lambda(|byte| byte == b' '),
            Some(b"World!".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b" ".as_bstr()));
        assert_eq!(
            stream.take(" ", Comparator::CaseSensitive),
            Some(b" ".as_bstr())
        );
        assert_eq!(
            stream.take_until_lambda(|byte| byte == b' '),
            Some(b"This".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b" ".as_bstr()));
        assert_eq!(
            stream.take(" ", Comparator::CaseSensitive),
            Some(b" ".as_bstr())
        );
        assert_eq!(
            stream.take_until_lambda(|byte| byte == b' '),
            Some(b"is".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b" ".as_bstr()));
        assert_eq!(
            stream.take(" ", Comparator::CaseSensitive),
            Some(b" ".as_bstr())
        );
        assert_eq!(
            stream.take_until_lambda(|byte| byte == b' '),
            Some(b"a".as_bstr())
        );
        assert_eq!(stream.peek(1), Some(b" ".as_bstr()));
        assert_eq!(
            stream.take(b" ", Comparator::CaseSensitive),
            Some(b" ".as_bstr())
        );
        assert_eq!(
            stream.take_until_lambda(|byte| byte == b'.'),
            Some(b"test".as_bstr())
        );
        assert_eq!(stream.is_empty(), false);
        assert_eq!(stream.peek(1), Some(b".".as_bstr()));
        assert_eq!(
            stream.take(".", Comparator::CaseSensitive),
            Some(b".".as_bstr())
        );
        assert_eq!(stream.is_empty(), true);
        assert_eq!(stream.take_until_lambda(|byte| byte == b' '), None);
        assert_eq!(stream.peek(1), None);
        assert_eq!(stream.is_empty(), true);
    }
}
