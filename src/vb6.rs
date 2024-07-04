use bstr::{BStr, ByteSlice};

use winnow::{
    ascii::{digit1, line_ending, Caseless},
    combinator::not,
    error::{ContextError, ErrMode, ParserError},
    stream::Stream,
    token::{literal, one_of, take_till, take_while},
    PResult, Parser,
};

/// Parses a VB6 end-of-line comment.
///
/// The comment starts with a single quote and continues until the end of the line.
/// But it does not include the single quote, the carriage return character, the newline character,
/// and it does not consume the carriage return or newline character.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// The comment without the single quote, carriage return, and newline characters.
///
/// # Example
///
/// ```rust
/// use vb6parse::vb6::eol_comment_parse;
///
/// let mut input = "' This is a comment\r\n".as_bytes();
/// let comment = eol_comment_parse(&mut input).unwrap();
///
/// assert_eq!(comment, b"' This is a comment");
/// ```
pub fn eol_comment_parse<'a>(input: &mut &'a [u8]) -> PResult<&'a [u8]> {
    let comment = ('\'', take_till(0.., ('\r', '\n')))
        .recognize()
        .parse_next(input)?;

    Ok(comment)
}

pub fn whitespace_parse<'a>(input: &mut &'a [u8]) -> PResult<&'a [u8]> {
    let whitespace = take_while(1.., (' ', '\t')).parse_next(input)?;

    Ok(whitespace)
}

/// Parses a VB6 variable name.
///
/// The variable name starts with a letter and can contain letters, numbers, and underscores.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// The VB6 variable name.
///
/// # Example
///
/// ```rust
/// use vb6parse::vb6::variable_name_parse;
///
/// let mut input = "variable_name".as_bytes();
/// let variable_name = variable_name_parse(&mut input).unwrap();
///
/// assert_eq!(variable_name, b"variable_name");
/// ```
pub fn variable_name_parse<'a>(input: &mut &'a [u8]) -> PResult<&'a [u8]> {
    let variable_name = (
        one_of(('a'..='z', 'A'..='Z')),
        take_while(0.., ('_', 'a'..='z', 'A'..='Z', '0'..='9')),
    )
        .recognize()
        .parse_next(input)?;

    Ok(variable_name)
}

/// Parses a VB6 keyword.
///
/// The keyword is case-insensitive.
///
/// # Arguments
///
/// * `keyword` - The keyword to parse.
///
/// # Returns
///
/// The keyword.
///
/// # Example
///
/// ```rust
/// use vb6parse::vb6::keyword_parse;
///
/// use winnow::Parser;
/// use winnow::error::{ParserError, ErrMode, ContextError};
///
/// use bstr::ByteSlice;
///
/// let mut op_parse = keyword_parse("Op");
///
/// let keyword = op_parse.parse_next(&mut "Option".as_bytes());
/// let keyword2 = op_parse.parse_next(&mut "op do".as_bytes());
///
///
/// assert_eq!(keyword, Err(ErrMode::Backtrack(ContextError::new())));
/// assert_eq!(keyword2, Ok("op".as_bytes()));
/// ```
pub fn keyword_parse<'a>(keyword: &'a str) -> impl FnMut(&mut &'a [u8]) -> PResult<&'a [u8]> {
    move |input: &mut &'a [u8]| {
        let checkpoint = input.checkpoint();

        let keyword: Result<&[u8], ErrMode<ContextError>> = Caseless(keyword).parse_next(input);

        let continuation = not::<&[u8], u8, ContextError, _>(one_of::<&[u8], _, _>((
            b'_',
            b'a'..=b'z',
            b'A'..=b'Z',
            b'0'..=b'9',
        )))
        .parse_next(input);

        match keyword {
            Ok(keyword) => {
                // the not indicates the keyword is not followed by a letter, number, or underscore.
                // and the 'not' function will give a success when it doesn't match.
                if continuation.is_ok() {
                    return Ok(keyword);
                }

                input.reset(&checkpoint);
                Err(ErrMode::Backtrack(ContextError::new()))
            }
            _ => {
                input.reset(&checkpoint);
                Err(ErrMode::Backtrack(ContextError::new()))
            }
        }
    }
}

/// Represents a VB6 token.
///
/// This is a simple enum that represents the different types of tokens that can be parsed from VB6 code.
///
/// # Example
///
/// ```rust
/// use vb6parse::vb6::VB6Token;
/// use bstr::ByteSlice;
///
/// let token =
///    VB6Token::VariableName(b"variable_name".as_bstr());
///
/// assert_eq!(token, VB6Token::VariableName(b"variable_name".as_bstr()));
/// ```
#[derive(Debug, PartialEq, Clone, Eq)]
pub enum VB6Token<'a> {
    Whitespace(&'a BStr),
    Newline(&'a BStr),

    Comment(&'a BStr),

    ReDimKeyword(&'a BStr),
    DimKeyword(&'a BStr),
    DeclareKeyword(&'a BStr),
    LibKeyword(&'a BStr),
    WithKeyword(&'a BStr),

    OptionKeyword(&'a BStr),
    ExplicitKeyword(&'a BStr),

    PrivateKeyword(&'a BStr),
    PublicKeyword(&'a BStr),

    ConstKeyword(&'a BStr),
    AsKeyword(&'a BStr),
    ByValKeyword(&'a BStr),
    ByRefKeyword(&'a BStr),

    FunctionKeyword(&'a BStr),
    SubKeyword(&'a BStr),
    EndKeyword(&'a BStr),

    TrueKeyword(&'a BStr),
    FalseKeyword(&'a BStr),

    EnumKeyword(&'a BStr),

    BooleanKeyword(&'a BStr),
    ByteKeyword(&'a BStr),
    LongKeyword(&'a BStr),
    SingleKeyword(&'a BStr),
    StringKeyword(&'a BStr),

    IfKeyword(&'a BStr),
    ElseKeyword(&'a BStr),
    AndKeyword(&'a BStr),
    OrKeyword(&'a BStr),
    NotKeyword(&'a BStr),
    ThenKeyword(&'a BStr),

    GotoKeyword(&'a BStr),
    ExitKeyword(&'a BStr),

    ForKeyword(&'a BStr),
    ToKeyword(&'a BStr),
    StepKeyword(&'a BStr),
    NextKeyword(&'a BStr),

    Ampersand(&'a BStr),
    Percent(&'a BStr),
    Octothorpe(&'a BStr),
    LeftParanthesis(&'a BStr),
    RightParanthesis(&'a BStr),
    Comma(&'a BStr),
    EqualityOperator(&'a BStr),
    LessThanOperator(&'a BStr),
    GreaterThanOperator(&'a BStr),
    MultiplicationOperator(&'a BStr),
    SubtractionOperator(&'a BStr),
    AdditionOperator(&'a BStr),
    DivisionOperator(&'a BStr),
    PeriodOperator(&'a BStr),
    ColonOperator(&'a BStr),

    VariableName(&'a BStr),
    Number(&'a BStr),
}

/// Parses a VB6 code lines.
pub fn vb6_parse<'a>(input: &mut &'a [u8]) -> PResult<Vec<VB6Token<'a>>> {
    let mut tokens = Vec::new();

    while !input.is_empty() {
        if tokens.len() == 671 {
            println!("{:?}", input[0..20].as_bstr());
        }

        if let Ok(token) = line_ending::<&'a [u8], ContextError>.parse_next(input) {
            tokens.push(VB6Token::Newline(token.as_bstr()));
            continue;
        }

        if let Ok(token) = eol_comment_parse.parse_next(input) {
            tokens.push(VB6Token::Comment(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Option").parse_next(input) {
            tokens.push(VB6Token::OptionKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Explicit").parse_next(input) {
            tokens.push(VB6Token::ExplicitKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Private").parse_next(input) {
            tokens.push(VB6Token::PrivateKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Public").parse_next(input) {
            tokens.push(VB6Token::PublicKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Dim").parse_next(input) {
            tokens.push(VB6Token::DimKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("With").parse_next(input) {
            tokens.push(VB6Token::WithKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Declare").parse_next(input) {
            tokens.push(VB6Token::DeclareKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Lib").parse_next(input) {
            tokens.push(VB6Token::LibKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Const").parse_next(input) {
            tokens.push(VB6Token::ConstKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("As").parse_next(input) {
            tokens.push(VB6Token::AsKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Enum").parse_next(input) {
            tokens.push(VB6Token::EnumKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Long").parse_next(input) {
            tokens.push(VB6Token::LongKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Boolean").parse_next(input) {
            tokens.push(VB6Token::BooleanKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Byte").parse_next(input) {
            tokens.push(VB6Token::ByteKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Single").parse_next(input) {
            tokens.push(VB6Token::SingleKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("String").parse_next(input) {
            tokens.push(VB6Token::StringKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("True").parse_next(input) {
            tokens.push(VB6Token::TrueKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("False").parse_next(input) {
            tokens.push(VB6Token::FalseKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Function").parse_next(input) {
            tokens.push(VB6Token::FunctionKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Sub").parse_next(input) {
            tokens.push(VB6Token::SubKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("End").parse_next(input) {
            tokens.push(VB6Token::EndKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("If").parse_next(input) {
            tokens.push(VB6Token::IfKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Else").parse_next(input) {
            tokens.push(VB6Token::ElseKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("And").parse_next(input) {
            tokens.push(VB6Token::AndKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Or").parse_next(input) {
            tokens.push(VB6Token::OrKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Not").parse_next(input) {
            tokens.push(VB6Token::NotKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Then").parse_next(input) {
            tokens.push(VB6Token::ThenKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("For").parse_next(input) {
            tokens.push(VB6Token::ForKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("To").parse_next(input) {
            tokens.push(VB6Token::ToKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Step").parse_next(input) {
            tokens.push(VB6Token::StepKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Next").parse_next(input) {
            tokens.push(VB6Token::NextKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("ReDim").parse_next(input) {
            tokens.push(VB6Token::ReDimKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("ByVal").parse_next(input) {
            tokens.push(VB6Token::ByValKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("ByRef").parse_next(input) {
            tokens.push(VB6Token::ByRefKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Goto").parse_next(input) {
            tokens.push(VB6Token::GotoKeyword(token.as_bstr()));
            continue;
        }

        if let Ok(token) = keyword_parse("Exit").parse_next(input) {
            tokens.push(VB6Token::ExitKeyword(token.as_bstr()));
            continue;
        }

        // Technically, this could be an equality operator or a assignment operator.
        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"=").parse_next(input) {
            tokens.push(VB6Token::EqualityOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"&").parse_next(input) {
            tokens.push(VB6Token::Ampersand(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"%").parse_next(input) {
            tokens.push(VB6Token::Percent(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"#").parse_next(input) {
            tokens.push(VB6Token::Octothorpe(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"<").parse_next(input) {
            tokens.push(VB6Token::LessThanOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b">").parse_next(input) {
            tokens.push(VB6Token::GreaterThanOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"(").parse_next(input) {
            tokens.push(VB6Token::LeftParanthesis(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b")").parse_next(input) {
            tokens.push(VB6Token::RightParanthesis(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b",").parse_next(input) {
            tokens.push(VB6Token::Comma(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"+").parse_next(input) {
            tokens.push(VB6Token::AdditionOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"-").parse_next(input) {
            tokens.push(VB6Token::SubtractionOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"*").parse_next(input) {
            tokens.push(VB6Token::MultiplicationOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b"/").parse_next(input) {
            tokens.push(VB6Token::DivisionOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b".").parse_next(input) {
            tokens.push(VB6Token::PeriodOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = literal::<&[u8], &'a [u8], ContextError>(b":").parse_next(input) {
            tokens.push(VB6Token::ColonOperator(token.as_bstr()));
            continue;
        }

        if let Ok(token) = digit1::<&'a [u8], ContextError>.parse_next(input) {
            tokens.push(VB6Token::Number(token.as_bstr()));
            continue;
        }

        if let Ok(token) = variable_name_parse.parse_next(input) {
            tokens.push(VB6Token::VariableName(token.as_bstr()));
            continue;
        }

        if let Ok(token) = whitespace_parse.parse_next(input) {
            tokens.push(VB6Token::Whitespace(token.as_bstr()));
            continue;
        }

        return Err(ErrMode::Cut(ParserError::assert(
            input,
            "Unable to match VB6 token.",
        )));
    }

    Ok(tokens)
}
