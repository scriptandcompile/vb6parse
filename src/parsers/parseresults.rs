use core::convert::From;
use std::iter::IntoIterator;

use crate::errors::ErrorDetails;

pub struct ParseResult<'a, T, E> {
    pub result: Option<T>,
    pub failures: Vec<ErrorDetails<'a, E>>,
}

impl<'a, T, E> ParseResult<'a, T, E> {
    #[inline]
    pub const fn has_result(&self) -> bool {
        matches!(self.result, Some(_))
    }

    #[inline]
    pub const fn has_failures(&self) -> bool {
        !matches!(self.failures.len(), 0)
    }

    pub fn push_failure(&mut self, failure: ErrorDetails<'a, E>) {
        self.failures.push(failure);
    }

    pub fn append_failures(&mut self, failures: &mut Vec<ErrorDetails<'a, E>>) {
        self.failures.append(failures);
    }

    pub fn unwrap(self) -> T {
        self.result.unwrap()
    }
}

impl<'a, T, E> From<(T, ErrorDetails<'a, E>)> for ParseResult<'a, T, E> {
    fn from(parse_pair: (T, ErrorDetails<'a, E>)) -> ParseResult<'a, T, E> {
        ParseResult {
            result: Some(parse_pair.0),
            failures: vec![parse_pair.1],
        }
    }
}

impl<'a, I, T, E> From<(I, Vec<ErrorDetails<'a, E>>)> for ParseResult<'a, Vec<T>, E>
where
    I: IntoIterator<Item = T>,
{
    fn from(parse_pair: (I, Vec<ErrorDetails<'a, E>>)) -> ParseResult<'a, Vec<T>, E> {
        let collection: Vec<T> = parse_pair.0.into_iter().collect();
        if collection.len() == 0 {
            return ParseResult {
                result: None,
                failures: parse_pair.1,
            };
        }

        ParseResult {
            result: Some(collection),
            failures: parse_pair.1,
        }
    }
}

// impl<'a, T, E> From<(T, Vec<ErrorDetails<'a, E>>)> for ParseResult<'a, T, E> {
//     fn from(parse_pair: (T, Vec<ErrorDetails<'a, E>>)) -> ParseResult<'a, T, E> {
//         ParseResult {
//             result: Some(parse_pair.0),
//             failures: parse_pair.1,
//         }
//     }
// }

// impl<'a, T, E> From<(Option<T>, Vec<ErrorDetails<'a, E>>)> for ParseResult<'a, T, E> {
//     fn from(parse_pair: (Option<T>, Vec<ErrorDetails<'a, E>>)) -> ParseResult<'a, T, E> {
//         ParseResult {
//             result: parse_pair.0,
//             failures: parse_pair.1,
//         }
//     }
// }
