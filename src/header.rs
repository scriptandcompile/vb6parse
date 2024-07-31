#![warn(clippy::pedantic)]

use bstr::BStr;
use miette::Result;

use crate::{vb6::line_comment_parse, vb6stream::VB6Stream};

use winnow::{
    ascii::{line_ending, space0},
    combinator::{alt, delimited, eof, opt, separated_pair},
    error::{ContextError, ErrMode},
    token::{literal, take_while},
    PResult, Parser,
};

pub fn key_value_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> PResult<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> Result<(&'a BStr, &'a BStr), ErrMode<ContextError>> {
        let (key, value) = separated_pair(
            delimited(
                space0,
                take_while(1.., ('_', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9')),
                space0,
            ),
            literal(divider),
            alt((
                delimited(
                    (space0, opt("\"")),
                    take_while(1.., ('_', '.', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9')),
                    (opt("\""), space0),
                ),
                delimited(
                    space0,
                    take_while(1.., ('_', '.', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9')),
                    space0,
                ),
            )),
        )
        .parse_next(input)?;

        Ok((key, value))
    }
}

pub fn key_value_line_parse<'a>(
    divider: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> PResult<(&'a BStr, &'a BStr)> {
    move |input: &mut VB6Stream<'a>| -> PResult<(&'a BStr, &'a BStr)> {
        let (key, value) = key_value_parse(divider).parse_next(input)?;

        // we have to check for eof here because it's perfectly possible to have a
        // header file that is empty of actual code. This means the last line of the file
        // should be an empty line, but it might be that the filed ends at the end of the
        // header attribute section.
        (space0, opt(line_comment_parse), alt((line_ending, eof))).parse_next(input)?;

        Ok((key, value))
    }
}
