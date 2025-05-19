use bstr::{BStr, ByteSlice};

use winnow::{
    ascii::{digit1, line_ending, space1, Caseless},
    combinator::{alt, delimited, repeat},
    error::ErrMode,
    stream::Stream,
    token::{literal, one_of, take_till, take_until, take_while},
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
/// use winnow::Parser;
/// use vb6parse::parsers::{vb6::line_comment_parse, VB6Stream};
///
/// let mut input = VB6Stream::new("line_comment.bas".to_owned(), "' This is a comment\r\n".as_bytes());
/// let comment = line_comment_parse.parse_next(&mut input).unwrap();
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
        one_of(('a'..='z', 'A'..='Z', 128..=255)),
        take_while(0.., ('_', 'a'..='z', 'A'..='Z', '0'..='9', 128..=255)),
    )
        .take()
        .parse_next(input)?;

    if variable_name.len() >= 255 {
        return Err(ErrMode::Cut(VB6ErrorKind::VariableNameTooLong));
    }

    Ok(variable_name)
}

/// Grabs until the end of line. This is usually used to grab the rest of the
/// line for a line comment.
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Errors
///
/// If the parser encounters an error, it will return a 'Cut' error and backup the stream.
///
/// # Returns
///
/// The rest of the line.
///
/// # Example
///
/// ```rust
/// use vb6parse::parsers::{vb6::take_until_line_ending, VB6Stream};
/// use bstr::{BStr, ByteSlice};
///
/// let mut input = VB6Stream::new("test.bas", b"Dim x As Integer\r\nDim y as String\r\n");
/// let line = take_until_line_ending(&mut input).unwrap();
///
/// assert_eq!(line, b"Dim x As Integer");
/// ```
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

        if one_of::<VB6Stream, _, VB6ErrorKind>(('_', 'a'..='z', 'A'..='Z', '0'..='9'))
            .parse_next(input)
            .is_ok()
        {
            input.reset(&checkpoint);

            return Err(ErrMode::Backtrack(VB6ErrorKind::KeywordNotFound));
        }

        Ok(word)
    }
}

/// Parses a VB6 string that may or may not contain double quotes (escaped by using two in a row).
/// This parser will return the string without the double quotes.
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Errors
///
/// If the string is not properly formatted, it will return a 'Cut' error and backup the stream.
///
/// # Example
///
/// ```
/// use crate::*;
/// use vb6parse::parsers::VB6Stream;
/// use vb6parse::parsers::vb6::string_parse;
///
/// let input_line2 = b"\"This is a string\"";
/// let input_line1 = b"\"This is also \"\"a\"\" string\"";
///
/// let mut stream1 = VB6Stream::new("", input_line1);
/// let mut stream2 = VB6Stream::new("", input_line2);
///
/// let string1 = string_parse(&mut stream1).unwrap();
/// let string2 = string_parse(&mut stream2).unwrap();
///
/// assert_eq!(string1, "This is also \"\"a\"\" string");
/// assert_eq!(string2, "This is a string");
/// ```
pub fn string_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    // We parse and build the string even though we won't actually return it
    // since we will just directly build a BStr from the input stream
    // THIS IS A HORRIBLE HACK! but at least it works.
    // TODO: Figure out how to actually get this right. Perhaps when we
    // change over all the BStrs to be owned types.
    let mut build_string =
        repeat(0.., string_fragment_parse).fold(Vec::new, |mut string, fragment| {
            match fragment {
                StringFragment::Literal(literal) => {
                    string.extend_from_slice(literal.as_bytes());
                }
                StringFragment::EscapedDoubleQuote(double_qoutes) => {
                    string.extend_from_slice(double_qoutes.as_bytes());
                }
            }
            string
        });

    "\"".parse_next(input)?;
    let start_index = input.index;

    build_string.parse_next(input)?;

    let end_index = input.index;
    "\"".parse_next(input)?;

    Ok(&input.stream[start_index..end_index])
}

enum StringFragment<'a> {
    Literal(&'a BStr),
    EscapedDoubleQuote(&'a BStr),
}

fn string_fragment_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<StringFragment<'a>> {
    let fragment = alt((
        "\"\"".take().map(StringFragment::EscapedDoubleQuote),
        take_until(1.., "\"").map(StringFragment::Literal),
    ))
    .parse_next(input)?;

    Ok(fragment)
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

    if !is_english_code(input.stream) {
        return Err(ErrMode::Cut(VB6ErrorKind::LikelyNonEnglishCharacterSet));
    }

    while !input.is_empty() {
        // The file should end if there is a null byte.
        if literal::<_, _, VB6ErrorKind>('\0')
            .parse_next(input)
            .is_ok()
        {
            break;
        }

        if let Ok(token) = line_ending::<VB6Stream<'a>, VB6ErrorKind>.parse_next(input) {
            let token = VB6Token::Newline(token);
            tokens.push(token);
            continue;
        }

        if let Ok(token) = line_comment_parse.parse_next(input) {
            let token = VB6Token::Comment(token);
            tokens.push(token);
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
            let token = VB6Token::StringLiteral(token);
            tokens.push(token);
            continue;
        }

        if let Ok(token) = vb6_token_parse.parse_next(input) {
            tokens.push(token);
            continue;
        }

        return Err(ErrMode::Cut(VB6ErrorKind::UnknownToken));
    }

    Ok(tokens)
}

#[must_use]
pub fn is_english_code(content: &BStr) -> bool {
    // We are looking to see if we have a large-ish number of higher half ANSI characters.
    let character_count = content.len();
    let higher_half_character_count = content.iter().filter(|&c| *c >= 128).count();

    higher_half_character_count == 0 || (100 * higher_half_character_count / character_count) < 1
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
            keyword_parse("WithEvents").map(|token: &BStr| VB6Token::WithEventsKeyword(token)),
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
            keyword_parse("Alias").map(|token: &BStr| VB6Token::AliasKeyword(token)),
        )),
        alt((
            keyword_parse("True").map(|token: &BStr| VB6Token::TrueKeyword(token)),
            keyword_parse("False").map(|token: &BStr| VB6Token::FalseKeyword(token)),
            keyword_parse("Function").map(|token: &BStr| VB6Token::FunctionKeyword(token)),
            keyword_parse("Sub").map(|token: &BStr| VB6Token::SubKeyword(token)),
            keyword_parse("End").map(|token: &BStr| VB6Token::EndKeyword(token)),
            keyword_parse("If").map(|token: &BStr| VB6Token::IfKeyword(token)),
            keyword_parse("ElseIf").map(|token: &BStr| VB6Token::ElseIfKeyword(token)),
            keyword_parse("And").map(|token: &BStr| VB6Token::AndKeyword(token)),
            keyword_parse("Or").map(|token: &BStr| VB6Token::OrKeyword(token)),
            keyword_parse("Not").map(|token: &BStr| VB6Token::NotKeyword(token)),
            keyword_parse("Then").map(|token: &BStr| VB6Token::ThenKeyword(token)),
            keyword_parse("For").map(|token: &BStr| VB6Token::ForKeyword(token)),
            keyword_parse("To").map(|token: &BStr| VB6Token::ToKeyword(token)),
            keyword_parse("Step").map(|token: &BStr| VB6Token::StepKeyword(token)),
            keyword_parse("Next").map(|token: &BStr| VB6Token::NextKeyword(token)),
            keyword_parse("ReDim").map(|token: &BStr| VB6Token::ReDimKeyword(token)),
            keyword_parse("Preserve").map(|token: &BStr| VB6Token::PreserveKeyword(token)),
            keyword_parse("ByVal").map(|token: &BStr| VB6Token::ByValKeyword(token)),
            keyword_parse("ByRef").map(|token: &BStr| VB6Token::ByRefKeyword(token)),
            keyword_parse("Goto").map(|token: &BStr| VB6Token::GotoKeyword(token)),
            keyword_parse("Wend").map(|token: &BStr| VB6Token::WendKeyword(token)),
        )),
        alt((
            keyword_parse("Exit").map(|token: &BStr| VB6Token::ExitKeyword(token)),
            keyword_parse("Compare").map(|token: &BStr| VB6Token::CompareKeyword(token)),
            keyword_parse("Static").map(|token: &BStr| VB6Token::StaticKeyword(token)),
            keyword_parse("Double").map(|token: &BStr| VB6Token::DoubleKeyword(token)),
            keyword_parse("Decimal").map(|token: &BStr| VB6Token::DecimalKeyword(token)),
            keyword_parse("Date").map(|token: &BStr| VB6Token::DateKeyword(token)),
            keyword_parse("Variant").map(|token: &BStr| VB6Token::VariantKeyword(token)),
            keyword_parse("Object").map(|token: &BStr| VB6Token::ObjectKeyword(token)),
            keyword_parse("Currency").map(|token: &BStr| VB6Token::CurrencyKeyword(token)),
            keyword_parse("Base").map(|token: &BStr| VB6Token::BaseKeyword(token)),
            keyword_parse("Else").map(|token: &BStr| VB6Token::ElseKeyword(token)),
            keyword_parse("Xor").map(|token: &BStr| VB6Token::XorKeyword(token)),
            keyword_parse("Mod").map(|token: &BStr| VB6Token::ModKeyword(token)),
            keyword_parse("Eqv").map(|token: &BStr| VB6Token::EqvKeyword(token)),
            keyword_parse("Imp").map(|token: &BStr| VB6Token::ImpKeyword(token)),
            keyword_parse("Is").map(|token: &BStr| VB6Token::IsKeyword(token)),
            keyword_parse("Lock").map(|token: &BStr| VB6Token::LockKeyword(token)),
            keyword_parse("Unlock").map(|token: &BStr| VB6Token::UnlockKeyword(token)),
            keyword_parse("Stop").map(|token: &BStr| VB6Token::StopKeyword(token)),
            keyword_parse("While").map(|token: &BStr| VB6Token::WhileKeyword(token)),
            keyword_parse("AddressOf").map(|token: &BStr| VB6Token::AddressOfKeyword(token)),
        )),
        alt((
            keyword_parse("Width").map(|token: &BStr| VB6Token::WidthKeyword(token)),
            keyword_parse("Write").map(|token: &BStr| VB6Token::WriteKeyword(token)),
            keyword_parse("Time").map(|token: &BStr| VB6Token::TimeKeyword(token)),
            keyword_parse("SetAttr").map(|token: &BStr| VB6Token::SetAttrKeyword(token)),
            keyword_parse("Set").map(|token: &BStr| VB6Token::SetKeyword(token)),
            keyword_parse("SendKeys").map(|token: &BStr| VB6Token::SendKeysKeyword(token)),
            keyword_parse("Select").map(|token: &BStr| VB6Token::SelectKeyword(token)),
            keyword_parse("Case").map(|token: &BStr| VB6Token::CaseKeyword(token)),
            keyword_parse("Seek").map(|token: &BStr| VB6Token::SeekKeyword(token)),
            keyword_parse("SaveSetting").map(|token: &BStr| VB6Token::SaveSettingKeyword(token)),
            keyword_parse("SavePicture").map(|token: &BStr| VB6Token::SavePictureKeyword(token)),
            keyword_parse("RSet").map(|token: &BStr| VB6Token::RSetKeyword(token)),
            keyword_parse("RmDir").map(|token: &BStr| VB6Token::RmDirKeyword(token)),
            keyword_parse("Resume").map(|token: &BStr| VB6Token::ResumeKeyword(token)),
            keyword_parse("Reset").map(|token: &BStr| VB6Token::ResetKeyword(token)),
            keyword_parse("Rem").map(|token: &BStr| VB6Token::RemKeyword(token)),
            keyword_parse("Randomize").map(|token: &BStr| VB6Token::RandomizeKeyword(token)),
            keyword_parse("RaiseEvent").map(|token: &BStr| VB6Token::RaiseEventKeyword(token)),
            keyword_parse("Put").map(|token: &BStr| VB6Token::PutKeyword(token)),
            keyword_parse("Property").map(|token: &BStr| VB6Token::PropertyKeyword(token)),
            keyword_parse("Print").map(|token: &BStr| VB6Token::PrintKeyword(token)),
        )),
        alt((
            keyword_parse("Open").map(|token: &BStr| VB6Token::OpenKeyword(token)),
            keyword_parse("On").map(|token: &BStr| VB6Token::OnKeyword(token)),
            keyword_parse("Name").map(|token: &BStr| VB6Token::NameKeyword(token)),
            keyword_parse("MkDir").map(|token: &BStr| VB6Token::MkDirKeyword(token)),
            keyword_parse("MidB").map(|token: &BStr| VB6Token::MidBKeyword(token)),
            keyword_parse("Mid").map(|token: &BStr| VB6Token::MidKeyword(token)),
            keyword_parse("LSet").map(|token: &BStr| VB6Token::LSetKeyword(token)),
            keyword_parse("Load").map(|token: &BStr| VB6Token::LoadKeyword(token)),
            keyword_parse("Line").map(|token: &BStr| VB6Token::LineKeyword(token)),
            keyword_parse("Input").map(|token: &BStr| VB6Token::InputKeyword(token)),
            keyword_parse("Let").map(|token: &BStr| VB6Token::LetKeyword(token)),
            keyword_parse("Kill").map(|token: &BStr| VB6Token::KillKeyword(token)),
            keyword_parse("Implements").map(|token: &BStr| VB6Token::ImplementsKeyword(token)),
            keyword_parse("Get").map(|token: &BStr| VB6Token::GetKeyword(token)),
            keyword_parse("FileCopy").map(|token: &BStr| VB6Token::FileCopyKeyword(token)),
            keyword_parse("Event").map(|token: &BStr| VB6Token::EventKeyword(token)),
            keyword_parse("Error").map(|token: &BStr| VB6Token::ErrorKeyword(token)),
            keyword_parse("Erase").map(|token: &BStr| VB6Token::EraseKeyword(token)),
            keyword_parse("Do").map(|token: &BStr| VB6Token::DoKeyword(token)),
            keyword_parse("Until").map(|token: &BStr| VB6Token::UntilKeyword(token)),
            keyword_parse("DeleteSetting")
                .map(|token: &BStr| VB6Token::DeleteSettingKeyword(token)),
        )),
        alt((
            keyword_parse("DefBool").map(|token: &BStr| VB6Token::DefBoolKeyword(token)),
            keyword_parse("DefByte").map(|token: &BStr| VB6Token::DefByteKeyword(token)),
            keyword_parse("DefInt").map(|token: &BStr| VB6Token::DefIntKeyword(token)),
            keyword_parse("DefLng").map(|token: &BStr| VB6Token::DefLngKeyword(token)),
            keyword_parse("DefCur").map(|token: &BStr| VB6Token::DefCurKeyword(token)),
            keyword_parse("DefSng").map(|token: &BStr| VB6Token::DefSngKeyword(token)),
            keyword_parse("DefDbl").map(|token: &BStr| VB6Token::DefDblKeyword(token)),
            keyword_parse("DefDec").map(|token: &BStr| VB6Token::DefDecKeyword(token)),
            keyword_parse("DefDate").map(|token: &BStr| VB6Token::DefDateKeyword(token)),
            keyword_parse("DefStr").map(|token: &BStr| VB6Token::DefStrKeyword(token)),
        )),
    ))
    .parse_next(input)
}

fn vb6_symbol_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6Token<'a>> {
    // 'alt' only allows for a limited number of parsers to be passed in.
    // so we need to chain the 'alt' parsers together.
    alt((
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
            "\\".map(|token: &BStr| VB6Token::BackwardSlashOperator(token)),
            "/".map(|token: &BStr| VB6Token::DivisionOperator(token)),
            ".".map(|token: &BStr| VB6Token::PeriodOperator(token)),
            ":".map(|token: &BStr| VB6Token::ColonOperator(token)),
            "^".map(|token: &BStr| VB6Token::ExponentiationOperator(token)),
        )),
        alt((
            "!".map(|token: &BStr| VB6Token::ExclamationMark(token)),
            "[".map(|token: &BStr| VB6Token::LeftSquareBracket(token)),
            "]".map(|token: &BStr| VB6Token::RightSquareBracket(token)),
            ";".map(|token: &BStr| VB6Token::Semicolon(token)),
            "@".map(|token: &BStr| VB6Token::AtSign(token)),
        )),
    ))
    .parse_next(input)
}

fn vb6_token_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6Token<'a>> {
    // 'alt' only allows for a limited number of parsers to be passed in.
    // so we need to chain the 'alt' parsers together.
    alt((
        (line_comment_parse).map(|token: &BStr| VB6Token::Comment(token)),
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
    fn no_escaped_double_quote_string_parse() {
        let input_line = b"\"This is a string\"\r\n";
        let mut stream = VB6Stream::new("", input_line);
        let string = string_parse(&mut stream).unwrap();

        assert_eq!(string, "This is a string");
    }

    #[test]
    fn contains_escaped_double_quote_string_parse() {
        let input_line = b"\"This is also \"\"a\"\" string\"\r\n";
        let mut stream = VB6Stream::new("", input_line);
        let string = string_parse(&mut stream).unwrap();

        assert_eq!(string, "This is also \"\"a\"\" string");
    }

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
    fn non_english_parse() {
        use crate::vb6::vb6_parse;
        use crate::vb6::VB6Stream;

        let code = "Option Explicit\r
Public app_path As String  '���|�]�w�X\r
Public ����H����ԤH��(1 To 2, 1 To 2) As Integer    '�����Ԩ���H�Ƭ�����(1.�ϥΪ�/2.�q��,1.�`�@�H��/2.�ثe�ĴX��)\r
Public ����ݾ��H��������(1 To 2, 1 To 3) As Integer    '����ݾ�����H���s��������(1.�ϥΪ�/2.�q��,1.���W����/2~3.�ݾ������n��s��)\r
Public �Ĥ@���Ұ�Ū�J�{�ǼаO As Boolean    '�Ĥ@���Ұʵ{��Ū�J�{�ǼаO��\r
Attribute �Ĥ@���Ұ�Ū�J�{�ǼаO.VB_VarUserMemId = 1073741834\r
Public �����ˬd����ؼм� As Integer    '�����ˬd����p�ƾ��ؼм�\r
Attribute �����ˬd����ؼм�.VB_VarUserMemId = 1073741836\r
Public �q������O�_�w�X�{ As Boolean    '���ҳq������O�_�w�g�X�{�Ȯ��ܼ�\r
Attribute �q������O�_�w�X�{.VB_VarUserMemId = 1073741837\r
Public ProgramIsOnWine As Boolean    '�{���O�_�B��Wine���ҤU����\r
Attribute ProgramIsOnWine.VB_VarUserMemId = 1073741838";

        let mut input = VB6Stream::new("", code.as_bytes());

        let result = vb6_parse(&mut input);

        assert!(result.is_err());
        assert!(matches!(
            result.unwrap_err(),
            ErrMode::Cut(VB6ErrorKind::LikelyNonEnglishCharacterSet)
        ));
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
