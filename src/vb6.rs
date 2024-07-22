use bstr::BStr;

use winnow::{
    ascii::{digit1, line_ending, Caseless},
    combinator::{alt, delimited},
    error::{ContextError, ErrMode, ParserError},
    stream::Stream,
    token::{one_of, take_till, take_while},
    PResult, Parser,
};

use crate::vb6stream::VB6Stream;

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
/// use vb6parse::vb6stream::VB6Stream;
///
/// let mut input = VB6Stream::new("' This is a comment\r\n".as_bytes());
/// let comment = eol_comment_parse(&mut input).unwrap();
///
/// assert_eq!(comment, "' This is a comment");
/// ```
pub fn eol_comment_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<&'a BStr> {
    let comment = ('\'', take_till(0.., ('\r', '\n')))
        .recognize()
        .parse_next(input)?;

    Ok(comment)
}

/// Parses whitespace.
///
/// Whitespace is defined as one or more spaces or tabs.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// The whitespace.
///
/// # Example
///
/// ```rust
/// use vb6parse::vb6::whitespace_parse;
/// use vb6parse::vb6stream::VB6Stream;
///
/// let mut input = VB6Stream::new("    t".as_bytes());
/// let whitespace = whitespace_parse(&mut input).unwrap();
///
/// assert_eq!(whitespace, "    ");
/// ```
pub fn whitespace_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<&'a BStr> {
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
/// use vb6parse::vb6stream::VB6Stream;
///
/// let mut input = VB6Stream::new("variable_name".as_bytes());
/// let variable_name = variable_name_parse(&mut input).unwrap();
///
/// assert_eq!(variable_name, "variable_name");
/// ```
pub fn variable_name_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<&'a BStr> {
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
/// use vb6parse::vb6stream::VB6Stream;
///
/// use bstr::{BStr, ByteSlice};
/// use winnow::error::{ContextError, ErrMode};
///
/// let mut input1 = VB6Stream::new("Option".as_bytes());
/// let mut input2 = VB6Stream::new("op do".as_bytes());
///
/// let mut op_parse = keyword_parse("Op");
///
/// let keyword = op_parse(&mut input1);
/// let keyword2 = op_parse(&mut input2);
///
/// assert_eq!(keyword, Err(ErrMode::Backtrack(ContextError::new())));
/// assert_eq!(keyword2, Ok(b"op".as_bstr()));
/// ```
pub fn keyword_parse<'a, 'b>(
    keyword: &'static str,
) -> impl FnMut(&mut VB6Stream<'a>) -> PResult<&'a BStr> {
    move |input: &mut VB6Stream<'a>| -> PResult<&'a BStr> {
        let checkpoint = input.checkpoint();

        let keyword = Caseless(keyword).parse_next(input)?;

        if one_of::<VB6Stream, _, ContextError>(('_', 'a'..='z', 'A'..='Z', '0'..='9'))
            .parse_next(input)
            .is_ok()
        {
            input.reset(&checkpoint);

            let context_error = ParserError::assert(input, "{keyword} followed by one of _, a-z, A-Z, or 0-9, this indicates {keyword} was likely found as part of another variable name. ie, 'Dim' found within 'Dimming'");
            return Err(ErrMode::Backtrack(context_error));
        }

        Ok(keyword)
    }
}

/// Represents a VB6 token.
///
/// This is a simple enum that represents the different types of tokens that can be parsed from VB6 code.
///
#[derive(Debug, PartialEq, Clone, Eq)]
pub enum VB6Token<'a> {
    /// Represents whitespace.
    Whitespace(&'a BStr),
    /// Represents a newline.
    /// This can be a carriage return, a newline, or a carriage return followed by a newline.
    Newline(&'a BStr),

    /// Represents a comment.
    /// Includes the single quote character.
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
    OptionalKeyword(&'a BStr),

    FunctionKeyword(&'a BStr),
    SubKeyword(&'a BStr),
    EndKeyword(&'a BStr),

    /// Represents the boolean literal `True`.
    TrueKeyword(&'a BStr),
    /// Represents the boolean literal `False`.
    FalseKeyword(&'a BStr),

    EnumKeyword(&'a BStr),
    TypeKeyword(&'a BStr),

    BooleanKeyword(&'a BStr),
    ByteKeyword(&'a BStr),
    LongKeyword(&'a BStr),
    SingleKeyword(&'a BStr),
    StringKeyword(&'a BStr),
    IntegerKeyword(&'a BStr),

    /// Represents a string literal.
    /// The string literal is enclosed in double quotes.
    StringLiteral(&'a BStr),

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

    /// Represents a dollar sign '$'.
    DollarSign(&'a BStr),
    /// Represents an underscore '_'.
    Underscore(&'a BStr),
    /// Represents an ampersand '&'.
    Ampersand(&'a BStr),
    /// Represents a percent sign '%'.
    Percent(&'a BStr),
    /// Represents an octothorpe '#'.
    Octothorpe(&'a BStr),
    /// Represents a left paranthesis '('.
    LeftParanthesis(&'a BStr),
    /// Represents a right paranthesis ')'.
    RightParanthesis(&'a BStr),
    /// Represents a comma ','.
    Comma(&'a BStr),

    /// Represents an equality operator '=' can also be the assignment operator.
    EqualityOperator(&'a BStr),
    /// Represents a less than operator '<'.
    LessThanOperator(&'a BStr),
    /// Represents a greater than operator '>'.
    GreaterThanOperator(&'a BStr),
    /// Represents a multiplication operator '*'.
    MultiplicationOperator(&'a BStr),
    /// Represents a subtraction operator '-'.
    SubtractionOperator(&'a BStr),
    /// Represents an addition operator '+'.
    AdditionOperator(&'a BStr),
    /// Represents a division operator '/'.
    DivisionOperator(&'a BStr),
    /// Represents a forward slash operator '\\'.
    ForwardSlashOperator(&'a BStr),
    /// Represents a period operator '.'.
    PeriodOperator(&'a BStr),
    /// Represents a colon operator ':'.
    ColonOperator(&'a BStr),
    /// Represents an exponentiation operator '^'.
    ExponentiationOperator(&'a BStr),

    /// Represents a variable name.
    /// This is a name that starts with a letter and can contain letters, numbers, and underscores.
    VariableName(&'a BStr),
    /// Represents a number.
    /// This is just a collection of digits and hasn't been parsed into a
    /// specific kind of number yet.
    Number(&'a BStr),
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
/// # Example
///
/// ```rust
/// use vb6parse::vb6::{vb6_parse, VB6Token};
/// use vb6parse::vb6stream::VB6Stream;
///
/// use bstr::{BStr, ByteSlice};
///
/// let mut input = VB6Stream::new("Dim x As Integer".as_bytes());
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
pub fn vb6_parse<'a>(input: &mut VB6Stream<'a>) -> PResult<Vec<VB6Token<'a>>> {
    let mut tokens = Vec::new();

    while !input.is_empty() {
        if let Ok(token) = line_ending::<VB6Stream<'a>, ContextError>.parse_next(input) {
            tokens.push(VB6Token::Newline(token));
            continue;
        }

        if let Ok(token) = eol_comment_parse.parse_next(input) {
            tokens.push(VB6Token::Comment(token));
            continue;
        }

        if let Ok(token) = delimited::<VB6Stream<'a>, _, &BStr, _, ContextError, _, _, _>(
            '\"',
            take_till(0.., '\"'),
            '\"',
        )
        .recognize()
        .parse_next(input)
        {
            tokens.push(VB6Token::StringLiteral(token));
            continue;
        }

        // 'alt' only allows for a limited number of parsers to be passed in.
        // so we need to chain the 'alt' parsers together.
        let token = alt((
            alt((
                "Type".map(|token: &BStr| VB6Token::TypeKeyword(token)),
                "Optional".map(|token: &BStr| VB6Token::OptionalKeyword(token)),
                "Option".map(|token: &BStr| VB6Token::OptionKeyword(token)),
                "Explicit".map(|token: &BStr| VB6Token::ExplicitKeyword(token)),
                "Private".map(|token: &BStr| VB6Token::PrivateKeyword(token)),
                "Public".map(|token: &BStr| VB6Token::PublicKeyword(token)),
                "Dim".map(|token: &BStr| VB6Token::DimKeyword(token)),
                "With".map(|token: &BStr| VB6Token::WithKeyword(token)),
                "Declare".map(|token: &BStr| VB6Token::DeclareKeyword(token)),
                "Lib".map(|token: &BStr| VB6Token::LibKeyword(token)),
                "Const".map(|token: &BStr| VB6Token::ConstKeyword(token)),
                "As".map(|token: &BStr| VB6Token::AsKeyword(token)),
                "Enum".map(|token: &BStr| VB6Token::EnumKeyword(token)),
                "Long".map(|token: &BStr| VB6Token::LongKeyword(token)),
                "Integer".map(|token: &BStr| VB6Token::IntegerKeyword(token)),
                "Boolean".map(|token: &BStr| VB6Token::BooleanKeyword(token)),
                "Byte".map(|token: &BStr| VB6Token::ByteKeyword(token)),
                "Single".map(|token: &BStr| VB6Token::SingleKeyword(token)),
                "String".map(|token: &BStr| VB6Token::StringKeyword(token)),
            )),
            alt((
                "True".map(|token: &BStr| VB6Token::TrueKeyword(token)),
                "False".map(|token: &BStr| VB6Token::FalseKeyword(token)),
                "Function".map(|token: &BStr| VB6Token::FunctionKeyword(token)),
                "Sub".map(|token: &BStr| VB6Token::SubKeyword(token)),
                "End".map(|token: &BStr| VB6Token::EndKeyword(token)),
                "If".map(|token: &BStr| VB6Token::IfKeyword(token)),
                "Else".map(|token: &BStr| VB6Token::ElseKeyword(token)),
                "And".map(|token: &BStr| VB6Token::AndKeyword(token)),
                "Or".map(|token: &BStr| VB6Token::OrKeyword(token)),
                "Not".map(|token: &BStr| VB6Token::NotKeyword(token)),
                "Then".map(|token: &BStr| VB6Token::ThenKeyword(token)),
                "For".map(|token: &BStr| VB6Token::ForKeyword(token)),
                "To".map(|token: &BStr| VB6Token::ToKeyword(token)),
                "Step".map(|token: &BStr| VB6Token::StepKeyword(token)),
                "Next".map(|token: &BStr| VB6Token::NextKeyword(token)),
                "ReDim".map(|token: &BStr| VB6Token::ReDimKeyword(token)),
                "ByVal".map(|token: &BStr| VB6Token::ByValKeyword(token)),
                "ByRef".map(|token: &BStr| VB6Token::ByRefKeyword(token)),
                "Goto".map(|token: &BStr| VB6Token::GotoKeyword(token)),
                "Exit".map(|token: &BStr| VB6Token::ExitKeyword(token)),
                "=".map(|token: &BStr| VB6Token::EqualityOperator(token)),
            )),
            alt((
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
                digit1.map(|token: &BStr| VB6Token::Number(token)),
                variable_name_parse.map(|token: &BStr| VB6Token::VariableName(token)),
                whitespace_parse.map(|token: &BStr| VB6Token::Whitespace(token)),
            )),
        ))
        .parse_next(input);

        if let Ok(token) = token {
            tokens.push(token);
            continue;
        }

        return Err(ErrMode::Cut(ParserError::assert(
            input,
            "Unable to match VB6 token.",
        )));
    }

    Ok(tokens)
}
