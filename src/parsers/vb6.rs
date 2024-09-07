use bstr::BStr;

use winnow::{
    ascii::{digit1, line_ending, space1, Caseless},
    combinator::{alt, delimited},
    error::{ContextError, ErrMode},
    stream::Stream,
    token::{one_of, take_till, take_until, take_while},
    Parser,
};

use crate::{errors::VB6ErrorKind, language::VB6Token, parsers::VB6Stream};

pub type VB6Result<T> = Result<T, ErrMode<VB6ErrorKind>>;

/// Parses a VB6 end-of-line comment.
///
/// The comment starts with a single quote and continues until the end of the
/// line. It includes the single quote, but excludes the carriage return
/// character, the newline character, and it does not consume the carriage
/// return or newline character.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Errors
///
/// Will return an error if it is not able to parse a comment. This can happen
/// if the comment is not terminated by a newline character, or if the comment
/// lacks a single quote.
///
/// # Returns
///
/// The comment with the single quote, but without carriage return, and
/// newline characters.
///
/// # Example
///
/// ```rust
/// use vb6parse::parsers::{vb6::line_comment_parse, VB6Stream};
///
/// let mut input = VB6Stream::new("line_comment.bas".to_owned(), "' This is a comment\r\n".as_bytes());
/// let comment = line_comment_parse(&mut input).unwrap();
///
/// assert_eq!(comment, "' This is a comment");
/// ```
pub fn line_comment_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    let comment = ('\'', take_till(0.., (b"\r\n", b"\n", b"\r")))
        .take()
        .parse_next(input)?;

    Ok(comment)
}

/// Parses a VB6 variable name.
///
/// The variable name starts with a letter and can contain letters, numbers, and underscores.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Errors
///
/// If the variable name is too long, it will return an error.
///
/// # Returns
///
/// The VB6 variable name.
///
/// # Example
///
/// ```rust
/// use vb6parse::parsers::{vb6::variable_name_parse, VB6Stream};
///
/// let mut input = VB6Stream::new("variable_name_test.bas".to_owned(), "variable_name".as_bytes());
/// let variable_name = variable_name_parse(&mut input).unwrap();
///
/// assert_eq!(variable_name, "variable_name");
/// ```
pub fn variable_name_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    let variable_name = (
        one_of(('a'..='z', 'A'..='Z')),
        take_while(0.., ('_', 'a'..='z', 'A'..='Z', '0'..='9')),
    )
        .take()
        .parse_next(input)?;

    if variable_name.len() >= 255 {
        return Err(ErrMode::Cut(VB6ErrorKind::VariableNameTooLong));
    }

    Ok(variable_name)
}

pub fn take_until_line_ending<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    alt((take_until(1.., "\r\n"), take_until(1.., "\n"))).parse_next(input)
}

/// Parses a VB6 keyword.
///
/// The keyword is case-insensitive.
///
/// # Arguments
///
/// * `keyword` - The keyword to parse.
///
/// # Errors
///
/// If the keyword is not found, it will return an error.
///
/// # Returns
///
/// The keyword.
///
/// # Example
///
/// ```rust
/// use vb6parse::{
///     parsers::{vb6::keyword_parse, VB6Stream},
///     errors::{VB6ErrorKind, VB6Error},
/// };
///
/// use bstr::{BStr, ByteSlice};
///
/// let mut input1 = VB6Stream::new("test1.bas", "Option".as_bytes());
/// let mut input2 = VB6Stream::new("test2.bas","op do".as_bytes());
///
/// let mut op_parse = keyword_parse("Op");
///
/// let keyword = op_parse(&mut input1);
/// let keyword2 = op_parse(&mut input2);
///
/// assert!(keyword.is_err());
/// assert_eq!(keyword2.unwrap(), b"op".as_bstr());
/// ```
pub fn keyword_parse<'a>(
    keyword: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<&'a BStr> {
        let checkpoint = input.checkpoint();

        let word = Caseless(keyword).parse_next(input)?;

        if one_of::<VB6Stream, _, ContextError>(('_', 'a'..='z', 'A'..='Z', '0'..='9'))
            .parse_next(input)
            .is_ok()
        {
            input.reset(&checkpoint);

            return Err(ErrMode::Backtrack(VB6ErrorKind::KeywordNotFound));
        }

        Ok(word)
    }
}

/// Parses VB6 code into a token stream.
///
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// A vector of VB6 tokens.
///
/// # Errors
///
/// If the parser encounters an unknown token, it will return an error.
///
/// # Example
///
/// ```rust
/// use vb6parse::language::VB6Token;
/// use vb6parse::parsers::{VB6Stream, vb6_parse};
///
/// use bstr::{BStr, ByteSlice};
///
/// let mut input = VB6Stream::new("test.bas", b"Dim x As Integer");
/// let tokens = vb6_parse(&mut input).unwrap();
///
/// assert_eq!(tokens.len(), 7);
/// assert_eq!(tokens[0], VB6Token::DimKeyword("Dim".into()));
/// assert_eq!(tokens[1], VB6Token::Whitespace(" ".into()));
/// assert_eq!(tokens[2], VB6Token::VariableName("x".into()));
/// assert_eq!(tokens[3], VB6Token::Whitespace(" ".into()));
/// assert_eq!(tokens[4], VB6Token::AsKeyword("As".into()));
/// assert_eq!(tokens[5], VB6Token::Whitespace(" ".into()));
/// assert_eq!(tokens[6], VB6Token::IntegerKeyword("Integer".into()));
/// ```
pub fn vb6_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<Vec<VB6Token<'a>>> {
    let mut tokens = Vec::new();

    while !input.is_empty() {
        if let Ok(token) = line_ending::<VB6Stream<'a>, VB6ErrorKind>.parse_next(input) {
            tokens.push(VB6Token::Newline(token));
            continue;
        }

        if let Ok(token) = line_comment_parse.parse_next(input) {
            tokens.push(VB6Token::Comment(token));
            continue;
        }

        if let Ok(token) = delimited::<VB6Stream<'a>, _, &BStr, _, VB6ErrorKind, _, _, _>(
            '\"',
            take_till(0.., '\"'),
            '\"',
        )
        .take()
        .parse_next(input)
        {
            tokens.push(VB6Token::StringLiteral(token));
            continue;
        }

        let token = vb6_token_parse.parse_next(input);

        if let Ok(token) = token {
            tokens.push(token);
            continue;
        }

        return Err(ErrMode::Cut(VB6ErrorKind::UnknownToken));
    }

    Ok(tokens)
}

fn vb6_keyword_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6Token<'a>> {
    // 'alt' only allows for a limited number of parsers to be passed in.
    // so we need to chain the 'alt' parsers together.
    alt((
        alt((
            keyword_parse("Type").map(|token: &BStr| VB6Token::TypeKeyword(token)),
            keyword_parse("Optional").map(|token: &BStr| VB6Token::OptionalKeyword(token)),
            keyword_parse("Option").map(|token: &BStr| VB6Token::OptionKeyword(token)),
            keyword_parse("Explicit").map(|token: &BStr| VB6Token::ExplicitKeyword(token)),
            keyword_parse("Private").map(|token: &BStr| VB6Token::PrivateKeyword(token)),
            keyword_parse("Public").map(|token: &BStr| VB6Token::PublicKeyword(token)),
            keyword_parse("Dim").map(|token: &BStr| VB6Token::DimKeyword(token)),
            keyword_parse("With").map(|token: &BStr| VB6Token::WithKeyword(token)),
            keyword_parse("Declare").map(|token: &BStr| VB6Token::DeclareKeyword(token)),
            keyword_parse("Lib").map(|token: &BStr| VB6Token::LibKeyword(token)),
            keyword_parse("Const").map(|token: &BStr| VB6Token::ConstKeyword(token)),
            keyword_parse("As").map(|token: &BStr| VB6Token::AsKeyword(token)),
            keyword_parse("Enum").map(|token: &BStr| VB6Token::EnumKeyword(token)),
            keyword_parse("Long").map(|token: &BStr| VB6Token::LongKeyword(token)),
            keyword_parse("Integer").map(|token: &BStr| VB6Token::IntegerKeyword(token)),
            keyword_parse("Boolean").map(|token: &BStr| VB6Token::BooleanKeyword(token)),
            keyword_parse("Byte").map(|token: &BStr| VB6Token::ByteKeyword(token)),
            keyword_parse("Single").map(|token: &BStr| VB6Token::SingleKeyword(token)),
            keyword_parse("String").map(|token: &BStr| VB6Token::StringKeyword(token)),
        )),
        alt((
            keyword_parse("True").map(|token: &BStr| VB6Token::TrueKeyword(token)),
            keyword_parse("False").map(|token: &BStr| VB6Token::FalseKeyword(token)),
            keyword_parse("Function").map(|token: &BStr| VB6Token::FunctionKeyword(token)),
            keyword_parse("Sub").map(|token: &BStr| VB6Token::SubKeyword(token)),
            keyword_parse("End").map(|token: &BStr| VB6Token::EndKeyword(token)),
            keyword_parse("If").map(|token: &BStr| VB6Token::IfKeyword(token)),
            keyword_parse("Else").map(|token: &BStr| VB6Token::ElseKeyword(token)),
            keyword_parse("And").map(|token: &BStr| VB6Token::AndKeyword(token)),
            keyword_parse("Or").map(|token: &BStr| VB6Token::OrKeyword(token)),
            keyword_parse("Not").map(|token: &BStr| VB6Token::NotKeyword(token)),
            keyword_parse("Then").map(|token: &BStr| VB6Token::ThenKeyword(token)),
            keyword_parse("For").map(|token: &BStr| VB6Token::ForKeyword(token)),
            keyword_parse("To").map(|token: &BStr| VB6Token::ToKeyword(token)),
            keyword_parse("Step").map(|token: &BStr| VB6Token::StepKeyword(token)),
            keyword_parse("Next").map(|token: &BStr| VB6Token::NextKeyword(token)),
            keyword_parse("ReDim").map(|token: &BStr| VB6Token::ReDimKeyword(token)),
            keyword_parse("ByVal").map(|token: &BStr| VB6Token::ByValKeyword(token)),
            keyword_parse("ByRef").map(|token: &BStr| VB6Token::ByRefKeyword(token)),
            keyword_parse("Goto").map(|token: &BStr| VB6Token::GotoKeyword(token)),
            keyword_parse("Exit").map(|token: &BStr| VB6Token::ExitKeyword(token)),
        )),
    ))
    .parse_next(input)
}

fn vb6_symbol_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6Token<'a>> {
    // 'alt' only allows for a limited number of parsers to be passed in.
    // so we need to chain the 'alt' parsers together.
    alt((
        "=".map(|token: &BStr| VB6Token::EqualityOperator(token)),
        "$".map(|token: &BStr| VB6Token::DollarSign(token)),
        "_".map(|token: &BStr| VB6Token::Underscore(token)),
        "&".map(|token: &BStr| VB6Token::Ampersand(token)),
        "%".map(|token: &BStr| VB6Token::Percent(token)),
        "#".map(|token: &BStr| VB6Token::Octothorpe(token)),
        "<".map(|token: &BStr| VB6Token::LessThanOperator(token)),
        ">".map(|token: &BStr| VB6Token::GreaterThanOperator(token)),
        "(".map(|token: &BStr| VB6Token::LeftParanthesis(token)),
        ")".map(|token: &BStr| VB6Token::RightParanthesis(token)),
        ",".map(|token: &BStr| VB6Token::Comma(token)),
        "+".map(|token: &BStr| VB6Token::AdditionOperator(token)),
        "-".map(|token: &BStr| VB6Token::SubtractionOperator(token)),
        "*".map(|token: &BStr| VB6Token::MultiplicationOperator(token)),
        "\\".map(|token: &BStr| VB6Token::ForwardSlashOperator(token)),
        "/".map(|token: &BStr| VB6Token::DivisionOperator(token)),
        ".".map(|token: &BStr| VB6Token::PeriodOperator(token)),
        ":".map(|token: &BStr| VB6Token::ColonOperator(token)),
        "^".map(|token: &BStr| VB6Token::ExponentiationOperator(token)),
        "!".map(|token: &BStr| VB6Token::ExclamationMark(token)),
    ))
    .parse_next(input)
}

fn vb6_token_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6Token<'a>> {
    // 'alt' only allows for a limited number of parsers to be passed in.
    // so we need to chain the 'alt' parsers together.
    alt((
        vb6_keyword_parse,
        vb6_symbol_parse,
        alt((
            digit1.map(|token: &BStr| VB6Token::Number(token)),
            variable_name_parse.map(|token: &BStr| VB6Token::VariableName(token)),
            space1.map(|token: &BStr| VB6Token::Whitespace(token)),
        )),
    ))
    .parse_next(input)
}

#[cfg(test)]
mod test {
    use super::*;
    use bstr::ByteSlice;

    #[test]
    fn keyword() {
        let mut input1 = VB6Stream::new("", "option".as_bytes());
        let mut input2 = VB6Stream::new("", "op do".as_bytes());

        let mut op_parse = keyword_parse("op");

        let keyword = op_parse(&mut input1);
        let keyword2 = op_parse(&mut input2);

        assert!(keyword.is_err());
        assert!(keyword2.is_ok());
        assert_eq!(keyword2.unwrap(), b"op".as_bstr());
    }

    #[test]
    fn eol_comment_carriage_return_newline() {
        use crate::parsers::VB6Stream;
        use crate::vb6::line_comment_parse;

        let mut input = VB6Stream::new("", "' This is a comment\r\n".as_bytes());
        let comment = line_comment_parse(&mut input).unwrap();

        assert_eq!(comment, "' This is a comment");
    }

    #[test]
    fn eol_comment_newline() {
        use crate::parsers::VB6Stream;
        use crate::vb6::line_comment_parse;

        let mut input = VB6Stream::new("", "' This is a comment\n".as_bytes());
        let comment = line_comment_parse(&mut input).unwrap();

        assert_eq!(comment, "' This is a comment");
    }

    #[test]
    fn eol_comment_carriage_return() {
        use crate::parsers::VB6Stream;
        use crate::vb6::line_comment_parse;

        let mut input = VB6Stream::new("", "' This is a comment\r".as_bytes());
        let comment = line_comment_parse(&mut input).unwrap();

        assert_eq!(comment, "' This is a comment");
    }

    #[test]
    fn eol_comment_eof() {
        use crate::parsers::VB6Stream;
        use crate::vb6::line_comment_parse;

        let mut input = VB6Stream::new("", "' This is a comment".as_bytes());
        let comment = line_comment_parse(&mut input).unwrap();

        assert_eq!(comment, "' This is a comment");
    }

    #[test]
    fn variable_name() {
        use crate::parsers::VB6Stream;
        use crate::vb6::variable_name_parse;

        let mut input = VB6Stream::new("", "variable_name".as_bytes());

        let variable_name = variable_name_parse(&mut input).unwrap();

        assert_eq!(variable_name, "variable_name");
    }

    #[test]
    fn vb6_parse() {
        use crate::parsers::VB6Stream;
        use crate::vb6::{vb6_parse, VB6Token};

        let mut input = VB6Stream::new("", "Dim x As Integer".as_bytes());
        let tokens = vb6_parse(&mut input).unwrap();

        assert_eq!(tokens.len(), 7);
        assert_eq!(tokens[0], VB6Token::DimKeyword("Dim".into()));
        assert_eq!(tokens[1], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[2], VB6Token::VariableName("x".into()));
        assert_eq!(tokens[3], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[4], VB6Token::AsKeyword("As".into()));
        assert_eq!(tokens[5], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[6], VB6Token::IntegerKeyword("Integer".into()));
    }

    #[test]
    fn multi_keyword() {
        use crate::vb6::keyword_parse;

        let mut input = VB6Stream::new("", "Option As Integer".as_bytes());

        let key1 = keyword_parse("Option").parse_next(&mut input).unwrap();

        let _ = space1::<_, VB6ErrorKind>.parse_next(&mut input);

        let key2 = keyword_parse("As").parse_next(&mut input).unwrap();

        let _ = space1::<_, VB6ErrorKind>.parse_next(&mut input);

        let key3 = keyword_parse("Integer").parse_next(&mut input).unwrap();

        assert_eq!(key1, "Option");
        assert_eq!(key2, "As");
        assert_eq!(key3, "Integer");
    }
}
