use bstr::ByteSlice;
use winnow::error::Needed;
use winnow::stream::{Offset, Stream};

use core::fmt::Debug;
use core::iter::{Cloned, Enumerate, Iterator};
use core::slice::Iter;
use std::num::NonZero;

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub struct VB6Stream<'a> {
    stream: &'a bstr::BStr,
}

impl<'a> VB6Stream<'a> {
    pub fn new(stream: &'a [u8]) -> Self {
        Self {
            stream: stream.as_bstr(),
        }
    }
}

impl<'a> Offset for VB6Stream<'a> {
    fn offset_from(&self, start: &Self) -> usize {
        start.stream.len() - self.stream.len()
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
        let (token, next) = self.stream.split_first()?;
        *self = VB6Stream::new(next);
        Some(*token)
    }

    fn offset_for<P>(&self, predicate: P) -> Option<usize>
    where
        P: Fn(Self::Token) -> bool,
    {
        self.stream.iter().position(|b| predicate(*b))
    }

    fn offset_at(&self, tokens: usize) -> Result<usize, Needed> {
        if let Some(needed) = tokens.checked_sub(self.stream.len()).and_then(NonZero::new) {
            Err(Needed::Size(needed))
        } else {
            Ok(tokens)
        }
    }

    fn next_slice(&mut self, offset: usize) -> Self::Slice {
        let (slice, rest) = self.stream.split_at(offset);
        self.stream = bstr::BStr::new(rest);
        bstr::BStr::new(slice)
    }

    fn checkpoint(&self) -> Self::Checkpoint {
        *self
    }

    fn reset(&mut self, checkpoint: &Self::Checkpoint) {
        self.stream = checkpoint.stream;
    }

    fn raw(&self) -> &dyn Debug {
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_vb6_stream() {
        let mut stream = VB6Stream::new(b"Hello, World!");
        assert_eq!(stream.next_token(), Some(b'H'));
        assert_eq!(stream.next_token(), Some(b'e'));
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b','));
        assert_eq!(stream.next_token(), Some(b' '));
        assert_eq!(stream.next_token(), Some(b'W'));
        assert_eq!(stream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b'r'));
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'd'));
        assert_eq!(stream.next_token(), Some(b'!'));
        assert_eq!(stream.next_token(), None);
    }

    #[test]
    fn test_vb6_stream_checkpoint() {
        let mut stream = VB6Stream::new(b"Hello, World!");

        assert_eq!(stream.next_token(), Some(b'H'));
        assert_eq!(stream.next_token(), Some(b'e'));
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b','));

        let checkpoint = stream.checkpoint();

        assert_eq!(stream.next_token(), Some(b' '));
        assert_eq!(stream.next_token(), Some(b'W'));
        assert_eq!(stream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b'r'));
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'd'));
        assert_eq!(stream.next_token(), Some(b'!'));
        assert_eq!(stream.next_token(), None);

        stream.reset(&checkpoint);

        assert_eq!(stream.next_token(), Some(b' '));
        assert_eq!(stream.next_token(), Some(b'W'));
        assert_eq!(stream.next_token(), Some(b'o'));
        assert_eq!(stream.next_token(), Some(b'r'));
        assert_eq!(stream.next_token(), Some(b'l'));
        assert_eq!(stream.next_token(), Some(b'd'));
        assert_eq!(stream.next_token(), Some(b'!'));
        assert_eq!(stream.next_token(), None);
    }

    #[test]
    fn test_vb6_stream_offset() {
        let stream = VB6Stream::new(b"Hello, World!");

        assert_eq!(stream.offset_for(|b| b == b'H'), Some(0));
        assert_eq!(stream.offset_for(|b| b == b'e'), Some(1));
        assert_eq!(stream.offset_for(|b| b == b'l'), Some(2));
        assert_eq!(stream.offset_for(|b| b == b'o'), Some(4));
        assert_eq!(stream.offset_for(|b| b == b','), Some(5));
        assert_eq!(stream.offset_for(|b| b == b' '), Some(6));
        assert_eq!(stream.offset_for(|b| b == b'W'), Some(7));
        assert_eq!(stream.offset_for(|b| b == b'r'), Some(9));
        assert_eq!(stream.offset_for(|b| b == b'd'), Some(11));
        assert_eq!(stream.offset_for(|b| b == b'!'), Some(12));
        assert_eq!(stream.offset_for(|b| b == b'z'), None);
    }
}
