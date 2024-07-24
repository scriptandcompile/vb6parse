use bstr::ByteSlice;
use winnow::ascii::Caseless;
use winnow::error::Needed;
use winnow::stream::{Compare, CompareResult, FindSlice, Offset, Stream, StreamIsPartial};

use core::fmt::Debug;
use core::iter::{Cloned, Enumerate, Iterator};
use core::slice::Iter;

use core::num::NonZeroUsize;
use std::num::NonZero;

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct VB6Stream<'a> {
    pub stream: &'a bstr::BStr,
    pub index: usize,
}

impl<'a> VB6Stream<'a> {
    pub fn new(stream: &'a [u8]) -> Self {
        Self {
            stream: stream.as_bstr(),
            index: 0,
        }
    }

    pub fn is_empty(&self) -> bool {
        self.stream.len() == self.index
    }
}

impl<'a> Offset for VB6Stream<'a> {
    fn offset_from(&self, start: &Self) -> usize {
        start.stream.len() - self.index
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

impl<'a> Compare<char> for VB6Stream<'a> {
    fn compare(&self, other: char) -> CompareResult {
        if self.stream[self.index..].len() < 1 {
            CompareResult::Incomplete
        } else if self.stream[self.index..].starts_with(other.to_string().as_bytes()) {
            CompareResult::Ok(0)
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
        } else if self.stream[self.index..].eq_ignore_ascii_case(other.0) {
            CompareResult::Ok(len)
        } else {
            CompareResult::Error
        }
    }
}

impl<'a> StreamIsPartial for VB6Stream<'a> {
    type PartialState = usize;

    fn complete(&mut self) -> usize {
        self.stream[self.index..].len() - self.index
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
    type Checkpoint = VB6Stream<'a>;

    fn iter_offsets(&self) -> Self::IterOffsets {
        self.stream.iter().cloned().enumerate()
    }

    fn eof_offset(&self) -> usize {
        self.stream.len()
    }

    fn next_token(&mut self) -> Option<Self::Token> {
        let (token, _) = self.stream[self.index..].split_first()?;
        self.index = self.index + 1;
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
        let (slice, _) = self.stream[self.index..].split_at(offset);
        self.index = self.index + offset;
        bstr::BStr::new(slice)
    }

    fn checkpoint(&self) -> Self::Checkpoint {
        *self
    }

    fn reset(&mut self, checkpoint: &Self::Checkpoint) {
        self.index = checkpoint.index;
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
        let mut stream = VB6Stream::new(b"Hello, World!");

        assert_eq!(wstream.next_slice(5), "Hello".as_bytes().as_bstr());
        assert_eq!(stream.next_slice(5), "Hello".as_bytes().as_bstr());

        assert_eq!(wstream.next_slice(2), ", ".as_bytes().as_bstr());
        assert_eq!(stream.next_slice(2), ", ".as_bytes().as_bstr());

        assert_eq!(wstream.next_slice(6), "World!".as_bytes().as_bstr());
        assert_eq!(stream.next_slice(6), "World!".as_bytes().as_bstr());
    }

    #[test]
    fn offset_at() {
        let wstream = b"Hello, World!".as_slice();
        let stream = VB6Stream::new(b"Hello, World!");

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
        let stream = VB6Stream::new(b"Hello, World!");

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
        let mut stream = VB6Stream::new(b"Hello, World!");

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
    fn checkpoint() {
        let mut wstream = b"Hello, World!".as_slice();
        let mut stream = VB6Stream::new(b"Hello, World!");

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
