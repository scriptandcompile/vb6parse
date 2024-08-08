use bstr::ByteSlice;

use winnow::{
    ascii::Caseless,
    error::Needed,
    stream::{Compare, CompareResult, FindSlice, Offset, Stream, StreamIsPartial},
};

use core::{
    fmt::Debug,
    iter::{Cloned, Enumerate, Iterator},
    num::NonZeroUsize,
    slice::Iter,
};

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct VB6Stream<'a> {
    pub file_name: String,
    pub stream: &'a bstr::BStr,
    pub index: usize,
    pub line_number: usize,
    pub column: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct VB6StreamCheckpoint {
    pub index: usize,
    pub line_number: usize,
    pub column: usize,
}

impl Offset for VB6StreamCheckpoint {
    fn offset_from(&self, start: &Self) -> usize {
        self.index - start.index
    }
}
impl Offset<VB6StreamCheckpoint> for VB6Stream<'_> {
    fn offset_from(&self, start: &VB6StreamCheckpoint) -> usize {
        self.index - start.index
    }
}

impl<'a, 'b> VB6Stream<'a> {
    pub fn new(file_name: impl Into<String>, stream: &'a [u8]) -> Self {
        Self {
            file_name: file_name.into(),
            stream: stream.as_bstr(),
            index: 0,
            line_number: 1,
            column: 1,
        }
    }

    pub fn is_empty(&self) -> bool {
        self.stream.len() == self.index
    }
}

impl<'a> FindSlice<&str> for VB6Stream<'a> {
    fn find_slice(&self, needle: &str) -> Option<std::ops::Range<usize>> {
        self.stream[self.index..]
            .find(needle)
            .map(|start| start..start + needle.len())
    }
}

impl<'a> FindSlice<(&str, &str)> for VB6Stream<'a> {
    fn find_slice(&self, needle: (&str, &str)) -> Option<std::ops::Range<usize>> {
        for needle in &[needle.0, needle.1] {
            if let Some(range) = self.stream[self.index..]
                .find(needle)
                .map(|start| start..start + needle.len())
            {
                return Some(range);
            }
        }

        None
    }
}

impl<'a> FindSlice<(&str, &str, &str)> for VB6Stream<'a> {
    fn find_slice(&self, needle: (&str, &str, &str)) -> Option<std::ops::Range<usize>> {
        for needle in &[needle.0, needle.1, needle.2] {
            if let Some(range) = self.stream[self.index..]
                .find(needle)
                .map(|start| start..start + needle.len())
            {
                return Some(range);
            }
        }

        None
    }
}

impl<'a> FindSlice<(&str, &str, &str, &str)> for VB6Stream<'a> {
    fn find_slice(&self, needle: (&str, &str, &str, &str)) -> Option<std::ops::Range<usize>> {
        for needle in &[needle.0, needle.1, needle.2, needle.3] {
            if let Some(range) = self.stream[self.index..]
                .find(needle)
                .map(|start| start..start + needle.len())
            {
                return Some(range);
            }
        }

        None
    }
}

impl<'a> FindSlice<char> for VB6Stream<'a> {
    fn find_slice(&self, needle: char) -> Option<std::ops::Range<usize>> {
        if let Some(range) = self.stream[self.index..]
            .find(needle.to_string())
            .map(|start| start..start + 1)
        {
            return Some(range);
        }

        None
    }
}

impl<'a> FindSlice<u8> for VB6Stream<'a> {
    fn find_slice(&self, needle: u8) -> Option<std::ops::Range<usize>> {
        if let Some(range) = self.stream[self.index..]
            .find(needle.to_string())
            .map(|start| start..start + 1)
        {
            return Some(range);
        }

        None
    }
}

impl<'a> FindSlice<(char, char)> for VB6Stream<'a> {
    fn find_slice(&self, needle: (char, char)) -> Option<std::ops::Range<usize>> {
        for needle in &[needle.0.to_string(), needle.1.to_string()] {
            if let Some(range) = self.stream[self.index..]
                .find(needle)
                .map(|start| start..start + needle.len())
            {
                return Some(range);
            }
        }

        None
    }
}

impl<'a> FindSlice<(u8, u8)> for VB6Stream<'a> {
    fn find_slice(&self, needle: (u8, u8)) -> Option<std::ops::Range<usize>> {
        for needle in &[needle.0.to_string(), needle.1.to_string()] {
            if let Some(range) = self.stream[self.index..]
                .find(needle)
                .map(|start| start..start + needle.len())
            {
                return Some(range);
            }
        }

        None
    }
}

impl<'a> FindSlice<(u8, u8, u8)> for VB6Stream<'a> {
    fn find_slice(&self, needle: (u8, u8, u8)) -> Option<std::ops::Range<usize>> {
        for needle in &[
            needle.0.to_string(),
            needle.1.to_string(),
            needle.2.to_string(),
        ] {
            if let Some(range) = self.stream[self.index..]
                .find(needle)
                .map(|start| start..start + needle.len())
            {
                return Some(range);
            }
        }

        None
    }
}

impl<'a> FindSlice<(u8, u8, u8, u8)> for VB6Stream<'a> {
    fn find_slice(&self, needle: (u8, u8, u8, u8)) -> Option<std::ops::Range<usize>> {
        for needle in &[
            needle.0.to_string(),
            needle.1.to_string(),
            needle.2.to_string(),
            needle.3.to_string(),
        ] {
            if let Some(range) = self.stream[self.index..]
                .find(needle)
                .map(|start| start..start + needle.len())
            {
                return Some(range);
            }
        }

        None
    }
}

impl<'a> Compare<char> for VB6Stream<'a> {
    fn compare(&self, other: char) -> CompareResult {
        if self.stream[self.index..].len() < 1 {
            CompareResult::Incomplete
        } else if self.stream[self.index..].starts_with(other.to_string().as_bytes()) {
            CompareResult::Ok(1)
        } else {
            CompareResult::Error
        }
    }
}

impl<'a> Compare<&str> for VB6Stream<'a> {
    fn compare(&self, other: &str) -> CompareResult {
        let other = other.as_bytes();
        let len = other.len();

        if self.stream[self.index..].len() < len {
            CompareResult::Incomplete
        } else if self.stream[self.index..].starts_with(other) {
            CompareResult::Ok(len)
        } else {
            CompareResult::Error
        }
    }
}

impl<'a> Compare<Caseless<&str>> for VB6Stream<'a> {
    fn compare(&self, other: Caseless<&str>) -> CompareResult {
        let other = other.as_bytes();
        let len = other.0.len();

        if self.stream[self.index..].len() < len {
            CompareResult::Incomplete
        } else if self.stream[self.index..(self.index + len)].eq_ignore_ascii_case(other.0) {
            CompareResult::Ok(len)
        } else {
            CompareResult::Error
        }
    }
}

impl<'a> StreamIsPartial for VB6Stream<'a> {
    type PartialState = usize;

    fn complete(&mut self) -> usize {
        self.stream[self.index..].len()
    }

    fn is_partial(&self) -> bool {
        self.index < self.stream.len()
    }

    fn restore_partial(&mut self, state: Self::PartialState) {
        self.index = state;
    }

    fn is_partial_supported() -> bool {
        false
    }
}

impl<'a> Stream for VB6Stream<'a> {
    type Token = u8;
    type Slice = &'a bstr::BStr;
    type IterOffsets = Enumerate<Cloned<Iter<'a, u8>>>;
    type Checkpoint = VB6StreamCheckpoint;

    fn iter_offsets(&self) -> Self::IterOffsets {
        self.stream[self.index..].iter().cloned().enumerate()
    }

    fn eof_offset(&self) -> usize {
        self.stream[self.index..].len()
    }

    fn next_token(&mut self) -> Option<Self::Token> {
        let (token, _) = self.stream[self.index..].split_first()?;
        self.index += 1;

        if *token == b'\n' {
            // if we have a newline then we need to reset the
            // column and line number.
            self.line_number += 1;
            self.column = 1;
        } else {
            self.column += 1;
        }

        Some(*token)
    }

    fn offset_for<P>(&self, predicate: P) -> Option<usize>
    where
        P: Fn(Self::Token) -> bool,
    {
        self.stream[self.index..].iter().position(|b| predicate(*b))
    }

    fn offset_at(&self, tokens: usize) -> Result<usize, Needed> {
        if let Some(needed) = tokens
            .checked_sub(self.stream[self.index..].len())
            .and_then(NonZeroUsize::new)
        {
            Err(Needed::Size(needed))
        } else {
            Ok(tokens)
        }
    }

    fn next_slice(&mut self, offset: usize) -> Self::Slice {
        let slice = self.stream[self.index..(self.index + offset)].as_bstr();

        self.index += offset;

        for token in slice.iter() {
            if *token == b'\n' {
                // on newline we need to reset the column and increment
                // the line number
                self.line_number += 1;
                self.column = 1;
            } else {
                self.column += 1;
            }
        }

        slice
    }

    fn checkpoint(&self) -> Self::Checkpoint {
        VB6StreamCheckpoint {
            index: self.index,
            line_number: self.line_number,
            column: self.column,
        }
    }

    fn reset(&mut self, checkpoint: &Self::Checkpoint) {
        self.index = checkpoint.index;
        self.line_number = checkpoint.line_number;
        self.column = checkpoint.column;
    }

    fn raw(&self) -> &dyn Debug {
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // Because the documentation of how stream is supposed to operate is a bit
    // unclear, these unit tests are based on the behavior of the winnow crate
    // and confirm that our behavior is consistent with their behavior.
    //
    // We first test with a normal stream, then with a winnow string stream.
    //
    // The two streams should behave the same way.
    #[test]
    fn next_slice() {
        let mut wstream = b"Hello, World!".as_slice();
        let mut stream = VB6Stream::new("", b"Hello, World!");

        assert_eq!(wstream.next_slice(5), "Hello".as_bytes().as_bstr());
        assert_eq!(stream.next_slice(5), "Hello".as_bytes().as_bstr());

        assert_eq!(wstream.next_slice(2), ", ".as_bytes().as_bstr());
        assert_eq!(stream.next_slice(2), ", ".as_bytes().as_bstr());

        assert_eq!(wstream.next_slice(6), "World!".as_bytes().as_bstr());
        assert_eq!(stream.next_slice(6), "World!".as_bytes().as_bstr());
    }

    #[test]
    fn compare() {
        let mut wstream = b"Hello, World!".as_slice();
        let mut stream = VB6Stream::new("", b"Hello, World!");

        let wcheckpoint = wstream.checkpoint();
        let checkpoint = stream.checkpoint();

        assert_eq!(wstream.compare("Hello"), CompareResult::Ok(5));
        assert_eq!(stream.compare("Hello"), CompareResult::Ok(5));

        assert_eq!(wstream.compare(", "), CompareResult::Error);
        assert_eq!(stream.compare(", "), CompareResult::Error);

        assert_eq!(wstream.next_slice(5), "Hello".as_bytes().as_bstr());
        assert_eq!(stream.next_slice(5), "Hello".as_bytes().as_bstr());

        assert_eq!(wstream.compare(", "), CompareResult::Ok(2));
        assert_eq!(stream.compare(", "), CompareResult::Ok(2));

        wstream.reset(&wcheckpoint);
        stream.reset(&checkpoint);

        assert_eq!(wstream.compare("World!"), CompareResult::Error);
        assert_eq!(stream.compare("World!"), CompareResult::Error);

        assert_eq!(wstream.compare("Hello, World!"), CompareResult::Ok(13));
        assert_eq!(stream.compare("Hello, World!"), CompareResult::Ok(13));

        assert_eq!(wstream.compare("Hello, World! "), CompareResult::Incomplete);
        assert_eq!(stream.compare("Hello, World! "), CompareResult::Incomplete);

        assert_eq!(wstream.compare("Hello, World"), CompareResult::Ok(12));
        assert_eq!(stream.compare("Hello, World"), CompareResult::Ok(12));

        assert_eq!(wstream.compare("Hello, World!!"), CompareResult::Incomplete);
        assert_eq!(stream.compare("Hello, World!!"), CompareResult::Incomplete);

        assert_eq!(wstream.compare("Hello, World! "), CompareResult::Incomplete);
        assert_eq!(stream.compare("Hello, World! "), CompareResult::Incomplete);
    }

    #[test]
    fn eof_offset() {
        let wstream = b"Hello, World!".as_slice();
        let stream = VB6Stream::new("", b"Hello, World!");

        assert_eq!(wstream.eof_offset(), stream.eof_offset());
    }

    #[test]
    fn iter_offsets() {
        let mut wstream = b"Hello, World!".as_slice();
        let mut stream = VB6Stream::new("", b"Hello, World!");

        assert_eq!(
            wstream.iter_offsets().collect::<Vec<_>>(),
            stream.iter_offsets().collect::<Vec<_>>()
        );

        assert_eq!(wstream.next_token(), Some(b'H'));
        assert_eq!(stream.next_token(), Some(b'H'));

        assert_eq!(wstream.next_token(), Some(b'e'));
        assert_eq!(stream.next_token(), Some(b'e'));

        assert_eq!(
            wstream.iter_offsets().collect::<Vec<_>>(),
            stream.iter_offsets().collect::<Vec<_>>()
        );
    }

    #[test]
    fn offset_at() {
        let wstream = b"Hello, World!".as_slice();
        let stream = VB6Stream::new("", b"Hello, World!");

        // Test offset_at with a valid offset
        assert_eq!(wstream.offset_at(5), Ok(5));
        assert_eq!(stream.offset_at(5), Ok(5));

        // Test offset_at with an offset that is on the last element
        assert_eq!(wstream.offset_at(13), Ok(13));
        assert_eq!(stream.offset_at(13), Ok(13));

        // Test offset_at with an offset that is too large
        assert_eq!(
            wstream.offset_at(14),
            Err(winnow::error::Needed::Size(NonZeroUsize::new(1).unwrap()))
        );
        assert_eq!(
            stream.offset_at(14),
            Err(winnow::error::Needed::Size(NonZeroUsize::new(1).unwrap()))
        );
    }

    #[test]
    fn offset_for() {
        let wstream = b"Hello, World!".as_slice();
        let stream = VB6Stream::new("", b"Hello, World!");

        // Test offset_for with a predicate that matches 'e'
        assert_eq!(wstream.offset_for(|b| b == b'e'), Some(1));
        assert_eq!(stream.offset_for(|b| b == b'e'), Some(1));

        // Test offset_for with a predicate that matches 'H'
        assert_eq!(wstream.offset_for(|b| b == b'H'), Some(0));
        assert_eq!(stream.offset_for(|b| b == b'H'), Some(0));

        // Test offset_for with a predicate that matches 'l'
        assert_eq!(wstream.offset_for(|b| b == b'l'), Some(2));
        assert_eq!(stream.offset_for(|b| b == b'l'), Some(2));

        // Test offset_for with a predicate that matches 'o'
        assert_eq!(wstream.offset_for(|b| b == b'o'), Some(4));
        assert_eq!(stream.offset_for(|b| b == b'o'), Some(4));

        // Test offset_for with a predicate that matches ','
        assert_eq!(wstream.offset_for(|b| b == b','), Some(5));
        assert_eq!(stream.offset_for(|b| b == b','), Some(5));

        // Test offset_for with a predicate that matches ' '
        assert_eq!(wstream.offset_for(|b| b == b' '), Some(6));
        assert_eq!(stream.offset_for(|b| b == b' '), Some(6));

        // Test offset_for with a predicate that matches 'W'
        assert_eq!(wstream.offset_for(|b| b == b'W'), Some(7));
        assert_eq!(stream.offset_for(|b| b == b'W'), Some(7));

        // Test offset_for with a predicate that matches 'r'
        assert_eq!(wstream.offset_for(|b| b == b'r'), Some(9));
        assert_eq!(stream.offset_for(|b| b == b'r'), Some(9));

        // Test offset_for with a predicate that matches 'd'
        assert_eq!(wstream.offset_for(|b| b == b'd'), Some(11));
        assert_eq!(stream.offset_for(|b| b == b'd'), Some(11));

        // Test offset_for with a predicate that matches '!'
        assert_eq!(wstream.offset_for(|b| b == b'!'), Some(12));
        assert_eq!(stream.offset_for(|b| b == b'!'), Some(12));

        // Test offset_for with a predicate that doesn't match any character
        assert_eq!(wstream.offset_for(|b| b == b'z'), None);
        assert_eq!(stream.offset_for(|b| b == b'z'), None);
    }

    #[test]
    fn stream() {
        let mut wstream = b"Hello, World!".as_slice();
        let mut stream = VB6Stream::new("", b"Hello, World!");

        assert_eq!(wstream.next_token(), Some(b'H'));
        assert_eq!(stream.next_token(), Some(b'H'));

        assert_eq!(wstream.next_token(), Some(b'e'));
        assert_eq!(stream.next_token(), Some(b'e'));

        assert_eq!(wstream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));

        assert_eq!(wstream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));

        assert_eq!(wstream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b'o'));

        assert_eq!(wstream.next_token(), Some(b','));
        assert_eq!(stream.next_token(), Some(b','));

        assert_eq!(wstream.next_token(), Some(b' '));
        assert_eq!(stream.next_token(), Some(b' '));

        assert_eq!(wstream.next_token(), Some(b'W'));
        assert_eq!(stream.next_token(), Some(b'W'));

        assert_eq!(wstream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b'o'));

        assert_eq!(wstream.next_token(), Some(b'r'));
        assert_eq!(stream.next_token(), Some(b'r'));

        assert_eq!(wstream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));

        assert_eq!(wstream.next_token(), Some(b'd'));
        assert_eq!(stream.next_token(), Some(b'd'));

        assert_eq!(wstream.next_token(), Some(b'!'));
        assert_eq!(stream.next_token(), Some(b'!'));

        assert_eq!(wstream.next_token(), None);
        assert_eq!(stream.next_token(), None);
    }

    #[test]
    fn line_and_column() {
        let mut stream = VB6Stream::new("", b"Hello,\r\n World!");

        let checkpoint = stream.checkpoint();

        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 1);
        assert_eq!(stream.next_token(), Some(b'H'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 2);
        assert_eq!(stream.next_token(), Some(b'e'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 3);
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 4);

        stream.reset(&checkpoint);

        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 1);
        assert_eq!(stream.next_token(), Some(b'H'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 2);
        assert_eq!(stream.next_token(), Some(b'e'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 3);
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 4);

        let checkpoint = stream.checkpoint();

        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 4);
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 5);
        assert_eq!(stream.next_token(), Some(b'o'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 6);
        assert_eq!(stream.next_token(), Some(b','));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 7);
        assert_eq!(stream.next_token(), Some(b'\r'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 8);
        assert_eq!(stream.next_token(), Some(b'\n'));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 1);
        assert_eq!(stream.next_token(), Some(b' '));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 2);

        stream.reset(&checkpoint);

        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 4);
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 5);
        assert_eq!(stream.next_token(), Some(b'o'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 6);
        assert_eq!(stream.next_token(), Some(b','));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 7);
        assert_eq!(stream.next_token(), Some(b'\r'));
        assert_eq!(stream.line_number, 1);
        assert_eq!(stream.column, 8);
        assert_eq!(stream.next_token(), Some(b'\n'));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 1);
        assert_eq!(stream.next_token(), Some(b' '));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 2);
        assert_eq!(stream.next_token(), Some(b'W'));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 3);
        assert_eq!(stream.next_token(), Some(b'o'));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 4);
        assert_eq!(stream.next_token(), Some(b'r'));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 5);
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 6);
        assert_eq!(stream.next_token(), Some(b'd'));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 7);
        assert_eq!(stream.next_token(), Some(b'!'));
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 8);
        assert_eq!(stream.next_token(), None);
        assert_eq!(stream.line_number, 2);
        assert_eq!(stream.column, 8);
    }

    #[test]
    fn checkpoint() {
        let mut wstream = b"Hello, World!".as_slice();
        let mut stream = VB6Stream::new("", b"Hello, World!");

        assert_eq!(wstream.next_token(), Some(b'H'));
        assert_eq!(stream.next_token(), Some(b'H'));

        assert_eq!(wstream.next_token(), Some(b'e'));
        assert_eq!(stream.next_token(), Some(b'e'));

        assert_eq!(wstream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));

        assert_eq!(wstream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));

        assert_eq!(wstream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b'o'));

        assert_eq!(wstream.next_token(), Some(b','));
        assert_eq!(stream.next_token(), Some(b','));

        let wcheckpoint = wstream.checkpoint();
        let checkpoint = stream.checkpoint();

        assert_eq!(wstream.next_token(), Some(b' '));
        assert_eq!(stream.next_token(), Some(b' '));

        assert_eq!(wstream.next_token(), Some(b'W'));
        assert_eq!(stream.next_token(), Some(b'W'));

        assert_eq!(wstream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b'o'));

        assert_eq!(wstream.next_token(), Some(b'r'));
        assert_eq!(stream.next_token(), Some(b'r'));

        assert_eq!(wstream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));

        assert_eq!(wstream.next_token(), Some(b'd'));
        assert_eq!(stream.next_token(), Some(b'd'));

        assert_eq!(wstream.next_token(), Some(b'!'));
        assert_eq!(stream.next_token(), Some(b'!'));

        assert_eq!(wstream.next_token(), None);
        assert_eq!(stream.next_token(), None);

        wstream.reset(&wcheckpoint);
        stream.reset(&checkpoint);

        assert_eq!(wstream.next_token(), Some(b' '));
        assert_eq!(stream.next_token(), Some(b' '));

        assert_eq!(wstream.next_token(), Some(b'W'));
        assert_eq!(stream.next_token(), Some(b'W'));

        assert_eq!(wstream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b'o'));

        assert_eq!(wstream.next_token(), Some(b'r'));
        assert_eq!(stream.next_token(), Some(b'r'));

        assert_eq!(wstream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));

        assert_eq!(wstream.next_token(), Some(b'd'));
        assert_eq!(stream.next_token(), Some(b'd'));

        assert_eq!(wstream.next_token(), Some(b'!'));
        assert_eq!(stream.next_token(), Some(b'!'));

        assert_eq!(wstream.next_token(), None);
        assert_eq!(stream.next_token(), None);
    }
}
