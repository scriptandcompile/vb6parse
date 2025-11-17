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

/// The result type for winnow based VB6 parsers.
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

/// Parses a VB6 REM full line comment.
///
/// The comment starts with 'REM' and continues until the end of the
/// line. It includes the 'REM' characters, but excludes the carriage return
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
/// does not start with 'REM'.
///
/// Note:
/// 'REM' must be followed by a whitespace character, a newline character, or a
/// non-alphanumeric character. This is to prevent 'REM' from being used as a
/// variable name, as well as to prevent parsing a variable such as 'reminder'
/// as a comment.
///
/// # Returns
///
/// The comment with the keyword 'REM', but without carriage return, and
/// newline characters.
///
/// # Example
///
/// ```rust
/// use winnow::Parser;
/// use vb6parse::parsers::{vb6::rem_comment_parse, VB6Stream};
///
/// let mut input = VB6Stream::new("rem_comment.bas".to_owned(), "REM This is a comment\r\n".as_bytes());
/// let comment = rem_comment_parse.parse_next(&mut input).unwrap();
///
/// assert_eq!(comment, "REM This is a comment");
/// ```
pub fn rem_comment_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    let comment = (
        keyword_parse("REM"),
        take_till(0.., (b"\r\n", b"\n", b"\r")),
    )
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
    // change over all the BStr's to be owned types.
    let mut build_string =
        repeat(0.., string_fragment_parse).fold(Vec::new, |mut string, fragment| {
            match fragment {
                StringFragment::Literal(literal) => {
                    string.extend_from_slice(literal.as_bytes());
                }
                StringFragment::EscapedDoubleQuote(double_quotes) => {
                    string.extend_from_slice(double_quotes.as_bytes());
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
/// assert_eq!(tokens[0], ("Dim", VB6Token::DimKeyword));
/// assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
/// assert_eq!(tokens[2], ("x", VB6Token::Identifier));
/// assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
/// assert_eq!(tokens[4], ("As", VB6Token::AsKeyword));
/// assert_eq!(tokens[5], (" ", VB6Token::Whitespace));
/// assert_eq!(tokens[6], ("Integer", VB6Token::IntegerKeyword));
/// ```
pub fn vb6_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<Vec<(&'a str, VB6Token)>> {
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
            let token_text = token.to_str().unwrap();
            tokens.push((token_text, VB6Token::Newline));
            continue;
        }

        if let Ok(token) = line_comment_parse.parse_next(input) {
            let token_text = token.to_str().unwrap();
            tokens.push((token_text, VB6Token::EndOfLineComment));
            continue;
        }

        if let Ok(token) = rem_comment_parse.parse_next(input) {
            let token_text = token.to_str().unwrap();
            tokens.push((token_text, VB6Token::RemComment));
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
            let token_text = token.to_str().unwrap();
            tokens.push((token_text, VB6Token::StringLiteral));
            continue;
        }

        if let Ok((token_text, token_type)) = vb6_token_parse.parse_next(input) {
            tokens.push((token_text, token_type));
            continue;
        }

        return Err(ErrMode::Cut(VB6ErrorKind::UnknownToken));
    }

    Ok(tokens)
}

/// Checks if the content is likely to be in English.
///
/// This function checks if the content contains a large number of higher half ANSI characters.
/// If the content contains a large number of higher half ANSI characters, it is likely not in English.
///
/// # Arguments
///
/// * `content` - The content to check.
///
/// # Returns
///
/// `true` if the content is likely in English, `false` otherwise.
///
/// # Example
///
/// ```rust
/// use vb6parse::parsers::vb6::is_english_code;
///
/// let content = b"Hello, World!";
/// let is_english = is_english_code(content.into());
///
/// assert_eq!(is_english, true);
///
/// let non_english_content = b"\xEF\xBF\xBD\xEF\xBF\xBD\xEF\xBF\xBD";
/// let is_english = is_english_code(non_english_content.into());
///
/// assert_eq!(is_english, false);
/// ```
#[must_use]
pub fn is_english_code(content: &BStr) -> bool {
    // We are looking to see if we have a large-ish number of higher half ANSI characters.
    let character_count = content.len();
    let higher_half_character_count = content.iter().filter(|&c| *c >= 128).count();

    higher_half_character_count == 0 || (100 * higher_half_character_count / character_count) < 1
}

fn vb6_keyword_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<(&'a str, VB6Token)> {
    // 'alt' only allows for a limited number of parsers to be passed in.
    // so we need to chain the 'alt' parsers together.

    alt((
        alt((
            keyword_parse("AddressOf")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::AddressOfKeyword)),
            keyword_parse("Alias")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::AliasKeyword)),
            keyword_parse("And")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::AndKeyword)),
            keyword_parse("AppActivate")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::AppActivateKeyword)),
            keyword_parse("As").map(|token: &BStr| (token.to_str().unwrap(), VB6Token::AsKeyword)),
            keyword_parse("Base")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::BaseKeyword)),
            keyword_parse("Beep")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::BeepKeyword)),
            keyword_parse("Binary")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::BinaryKeyword)),
            keyword_parse("Boolean")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::BooleanKeyword)),
            keyword_parse("ByRef")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ByRefKeyword)),
            keyword_parse("Byte")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ByteKeyword)),
            keyword_parse("ByVal")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ByValKeyword)),
            keyword_parse("Call")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::CallKeyword)),
            keyword_parse("Case")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::CaseKeyword)),
            keyword_parse("ChDir")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ChDirKeyword)),
            keyword_parse("ChDrive")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ChDriveKeyword)),
            keyword_parse("Close")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::CloseKeyword)),
            keyword_parse("Compare")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::CompareKeyword)),
            keyword_parse("Const")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ConstKeyword)),
            keyword_parse("Currency")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::CurrencyKeyword)),
            keyword_parse("Date")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DateKeyword)),
        )),
        alt((
            keyword_parse("Decimal")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DecimalKeyword)),
            keyword_parse("Declare")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DeclareKeyword)),
            keyword_parse("DefBool")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefBoolKeyword)),
            keyword_parse("DefByte")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefByteKeyword)),
            keyword_parse("DefCur")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefCurKeyword)),
            keyword_parse("DefDate")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefDateKeyword)),
            keyword_parse("DefDbl")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefDblKeyword)),
            keyword_parse("DefDec")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefDecKeyword)),
            keyword_parse("DefInt")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefIntKeyword)),
            keyword_parse("DefLng")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefLngKeyword)),
            keyword_parse("DefObj")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefObjKeyword)),
            keyword_parse("DefSng")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefSngKeyword)),
            keyword_parse("DefStr")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefStrKeyword)),
            keyword_parse("DefVar")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DefVarKeyword)),
            keyword_parse("DeleteSetting")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DeleteSettingKeyword)),
            keyword_parse("Dim")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DimKeyword)),
            keyword_parse("Do").map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DoKeyword)),
            keyword_parse("Double")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DoubleKeyword)),
            // switched so that `Else` isn't selected before `ElseIf`.
            keyword_parse("ElseIf")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ElseIfKeyword)),
        )),
        alt((
            keyword_parse("Else")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ElseKeyword)),
            keyword_parse("Empty")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::EmptyKeyword)),
            keyword_parse("End")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::EndKeyword)),
            keyword_parse("Enum")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::EnumKeyword)),
            keyword_parse("Eqv")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::EqvKeyword)),
            keyword_parse("Erase")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::EraseKeyword)),
            keyword_parse("Error")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ErrorKeyword)),
            keyword_parse("Event")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::EventKeyword)),
            keyword_parse("Exit")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ExitKeyword)),
            keyword_parse("Explicit")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ExplicitKeyword)),
            keyword_parse("False")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::FalseKeyword)),
            keyword_parse("FileCopy")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::FileCopyKeyword)),
            keyword_parse("For")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ForKeyword)),
            keyword_parse("Friend")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::FriendKeyword)),
            keyword_parse("Function")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::FunctionKeyword)),
            keyword_parse("Get")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::GetKeyword)),
            keyword_parse("Goto")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::GotoKeyword)),
            keyword_parse("If").map(|token: &BStr| (token.to_str().unwrap(), VB6Token::IfKeyword)),
            // switched so that `Imp` isn't selected before `Implements`.
            keyword_parse("Implements")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ImplementsKeyword)),
            keyword_parse("Imp")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ImpKeyword)),
            keyword_parse("Input")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::InputKeyword)),
        )),
        alt((
            keyword_parse("Integer")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::IntegerKeyword)),
            keyword_parse("Is").map(|token: &BStr| (token.to_str().unwrap(), VB6Token::IsKeyword)),
            keyword_parse("Kill")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::KillKeyword)),
            keyword_parse("Len")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LenKeyword)),
            keyword_parse("Let")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LetKeyword)),
            keyword_parse("Lib")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LibKeyword)),
            keyword_parse("Line")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LineKeyword)),
            keyword_parse("Lock")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LockKeyword)),
            keyword_parse("Load")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LoadKeyword)),
            keyword_parse("Long")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LongKeyword)),
            keyword_parse("LSet")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LSetKeyword)),
            keyword_parse("Me").map(|token: &BStr| (token.to_str().unwrap(), VB6Token::MeKeyword)),
            keyword_parse("Mid")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::MidKeyword)),
            keyword_parse("MkDir")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::MkDirKeyword)),
            keyword_parse("Mod")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ModKeyword)),
            keyword_parse("Name")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::NameKeyword)),
            keyword_parse("New")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::NewKeyword)),
            keyword_parse("Next")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::NextKeyword)),
            keyword_parse("Not")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::NotKeyword)),
            keyword_parse("Null")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::NullKeyword)),
            keyword_parse("Object")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ObjectKeyword)),
        )),
        alt((
            keyword_parse("On").map(|token: &BStr| (token.to_str().unwrap(), VB6Token::OnKeyword)),
            keyword_parse("Open")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::OpenKeyword)),
            // Switched so that `Option` isn't selected before `Optional`.
            keyword_parse("Optional")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::OptionalKeyword)),
            keyword_parse("Option")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::OptionKeyword)),
            keyword_parse("Or").map(|token: &BStr| (token.to_str().unwrap(), VB6Token::OrKeyword)),
            keyword_parse("ParamArray")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ParamArrayKeyword)),
            keyword_parse("Preserve")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::PreserveKeyword)),
            keyword_parse("Print")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::PrintKeyword)),
            keyword_parse("Private")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::PrivateKeyword)),
            keyword_parse("Property")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::PropertyKeyword)),
            keyword_parse("Public")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::PublicKeyword)),
            keyword_parse("Put")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::PutKeyword)),
            keyword_parse("RaiseEvent")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::RaiseEventKeyword)),
            keyword_parse("Randomize")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::RandomizeKeyword)),
            keyword_parse("ReDim")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ReDimKeyword)),
            keyword_parse("Reset")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ResetKeyword)),
            keyword_parse("Resume")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ResumeKeyword)),
            keyword_parse("RmDir")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::RmDirKeyword)),
            keyword_parse("RSet")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::RSetKeyword)),
            keyword_parse("SavePicture")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SavePictureKeyword)),
        )),
        alt((
            keyword_parse("SaveSetting")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SaveSettingKeyword)),
            keyword_parse("Seek")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SeekKeyword)),
            keyword_parse("Select")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SelectKeyword)),
            keyword_parse("SendKeys")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SendKeysKeyword)),
            // Switched so that `Set` isn't selected before `SetAttr`.
            keyword_parse("SetAttr")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SetAttrKeyword)),
            keyword_parse("Set")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SetKeyword)),
            keyword_parse("Single")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SingleKeyword)),
            keyword_parse("Static")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::StaticKeyword)),
            keyword_parse("Step")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::StepKeyword)),
            keyword_parse("Stop")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::StopKeyword)),
            keyword_parse("String")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::StringKeyword)),
            keyword_parse("Sub")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SubKeyword)),
            keyword_parse("Then")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ThenKeyword)),
            keyword_parse("Time")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::TimeKeyword)),
            keyword_parse("To").map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ToKeyword)),
            keyword_parse("True")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::TrueKeyword)),
            keyword_parse("Type")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::TypeKeyword)),
            keyword_parse("Unlock")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::UnlockKeyword)),
            keyword_parse("Until")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::UntilKeyword)),
            keyword_parse("Variant")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::VariantKeyword)),
            keyword_parse("Wend")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::WendKeyword)),
        )),
        alt((
            keyword_parse("While")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::WhileKeyword)),
            keyword_parse("Width")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::WidthKeyword)),
            // Switched so that `With` isn't selected before `WithEvents`.
            keyword_parse("WithEvents")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::WithEventsKeyword)),
            keyword_parse("With")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::WithKeyword)),
            keyword_parse("Write")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::WriteKeyword)),
            keyword_parse("Xor")
                .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::XorKeyword)),
        )),
    ))
    .parse_next(input)
}

fn vb6_symbol_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<(&'a str, VB6Token)> {
    // 'alt' only allows for a limited number of parsers to be passed in.
    // so we need to chain the 'alt' parsers together.
    alt((
        alt((
            "=".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::EqualityOperator)),
            "$".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DollarSign)),
            "_".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Underscore)),
            "&".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Ampersand)),
            "%".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Percent)),
            "#".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Octothorpe)),
            "<".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LessThanOperator)),
            ">".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::GreaterThanOperator)),
            "(".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LeftParentheses)),
            ")".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::RightParentheses)),
            ",".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Comma)),
            "+".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::AdditionOperator)),
            "-".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::SubtractionOperator)),
            "*".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::MultiplicationOperator)),
            "\\".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::BackwardSlashOperator)),
            "/".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::DivisionOperator)),
            ".".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::PeriodOperator)),
            ":".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ColonOperator)),
            "^".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ExponentiationOperator)),
        )),
        alt((
            "!".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::ExclamationMark)),
            "[".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::LeftSquareBracket)),
            "]".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::RightSquareBracket)),
            ";".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Semicolon)),
            "@".map(|token: &BStr| (token.to_str().unwrap(), VB6Token::AtSign)),
        )),
    ))
    .parse_next(input)
}

fn vb6_token_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<(&'a str, VB6Token)> {
    // 'alt' only allows for a limited number of parsers to be passed in.
    // so we need to chain the 'alt' parsers together.
    alt((
        (line_comment_parse)
            .map(|token: &BStr| (token.to_str().unwrap(), VB6Token::EndOfLineComment)),
        vb6_keyword_parse,
        vb6_symbol_parse,
        alt((
            digit1.map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Number)),
            variable_name_parse.map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Identifier)),
            space1.map(|token: &BStr| (token.to_str().unwrap(), VB6Token::Whitespace)),
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
        assert_eq!(tokens[0], ("Dim", VB6Token::DimKeyword));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("x", VB6Token::Identifier));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("As", VB6Token::AsKeyword));
        assert_eq!(tokens[5], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[6], ("Integer", VB6Token::IntegerKeyword));
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
