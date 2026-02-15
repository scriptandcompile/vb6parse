//! Tokenization module for VB6 source code.
//!
//! Provides functionality to tokenize VB6 source code into a stream of tokens.
//!
//! # Example
//! ```rust
//! use vb6parse::language::Token;
//! use vb6parse::lexer::tokenize;
//! use vb6parse::io::SourceStream;
//! let mut input = SourceStream::new("test.bas", "Dim x As Integer");
//! let result = tokenize(&mut input);
//! if result.has_failures() {
//!    for failure in result.failures() {
//!       failure.print();
//!   }
//!   panic!("Failed to parse vb6 code.");
//! }
//! let tokens = result.unwrap();
//! assert_eq!(tokens.len(), 7);
//! assert_eq!(tokens[0], ("Dim", Token::DimKeyword));
//! assert_eq!(tokens[1], (" ", Token::Whitespace));
//! assert_eq!(tokens[2], ("x", Token::Identifier));
//! assert_eq!(tokens[3], (" ", Token::Whitespace));
//! assert_eq!(tokens[4], ("As", Token::AsKeyword));
//! assert_eq!(tokens[5], (" ", Token::Whitespace));
//! assert_eq!(tokens[6], ("Integer", Token::IntegerKeyword));
//! ```
//!
//! # Overview
//!
//! The `tokenize` module provides functionality to parse VB6 source code into a stream of tokens.
//! This is a crucial step in the parsing process, as it breaks down the source code into manageable pieces
//! that can be further analyzed and processed.
//!
//! The main function in this module is `tokenize`, which takes a `SourceStream` as input
//! and returns a `ParseResult` containing a `TokenStream` and/or a list of errors.
//!
//! The module uses lookup tables to efficiently identify keywords and symbols in the VB6 language.
//! These tables map strings to their corresponding `Token` enum variants, allowing for quick
//! identification during the tokenization process.
//!
//! The tokenization process handles various types of tokens, including keywords, symbols,
//! identifiers, literals (string, numeric, date), comments, and whitespace.
//!
//! # See Also
//!
//! - [`SourceStream`]: Low-level character stream with offset tracking and line/column info
//! - [`TokenStream`]: Tokenized stream of VB6 tokens
//! - [`ParseResult`]: Result type for parsing operations, including errors
//! - [`Token`]: Enum representing VB6 tokens
//! - [`ErrorDetails`](crate::errors::ErrorDetails): Detailed error information for parsing operations
//!

pub mod token_stream;

pub use crate::language::Token;
pub use token_stream::TokenStream;

use phf::{phf_ordered_map, OrderedMap};

use crate::{
    io::SourceStream,
    parsers::{Comparator, ParseResult},
    ErrorKind,
};

/// Lookup table for VB6 keywords to their corresponding tokens.
/// This table is used during the tokenization process to quickly identify
/// keywords in the source code.
static KEYWORD_TOKEN_LOOKUP_TABLE: OrderedMap<&'static str, Token> = phf_ordered_map! {
    "AdressOf" => Token::AddressOfKeyword,
    "Access" => Token::AccessKeyword,
    "Alias" => Token::AliasKeyword,
    "And" => Token::AndKeyword,
    "AppActivate" => Token::AppActivateKeyword,
    "Append" => Token::AppendKeyword,
    "Attribute" => Token::AttributeKeyword,
    "As" => Token::AsKeyword,
    "Base" => Token::BaseKeyword,
    "Beep" => Token::BeepKeyword,
    "Begin" => Token::BeginKeyword,
    "Binary" => Token::BinaryKeyword,
    "Boolean" => Token::BooleanKeyword,
    "ByRef" => Token::ByRefKeyword,
    "Byte" => Token::ByteKeyword,
    "ByVal" => Token::ByValKeyword,
    "Call" => Token::CallKeyword,
    "Case" => Token::CaseKeyword,
    "ChDir" => Token::ChDirKeyword,
    "ChDrive" => Token::ChDriveKeyword,
    "Class" => Token::ClassKeyword,
    "Close" => Token::CloseKeyword,
    "Compare" => Token::CompareKeyword,
    "Const" => Token::ConstKeyword,
    "Currency" => Token::CurrencyKeyword,
    "Date" => Token::DateKeyword,
    "Decimal" => Token::DecimalKeyword,
    "Declare" => Token::DeclareKeyword,
    "DefBool" => Token::DefBoolKeyword,
    "DefByte" => Token::DefByteKeyword,
    "DefCur" => Token::DefCurKeyword,
    "DefDate" => Token::DefDateKeyword,
    "DefDbl" => Token::DefDblKeyword,
    "DefDec" => Token::DefDecKeyword,
    "DefInt" => Token::DefIntKeyword,
    "DefLng" => Token::DefLngKeyword,
    "DefObj" => Token::DefObjKeyword,
    "DefSng" => Token::DefSngKeyword,
    "DefStr" => Token::DefStrKeyword,
    "DefVar" => Token::DefVarKeyword,
    "DeleteSetting" => Token::DeleteSettingKeyword,
    "Dim" => Token::DimKeyword,
    // switched so that `Do` isn't selected for `Double`.
    "Double" => Token::DoubleKeyword,
    "Do" => Token::DoKeyword,
    "Each" => Token::EachKeyword,
    // switched so that `Else` isn't selected for `ElseIf`.
    "ElseIf" => Token::ElseIfKeyword,
    "Else" => Token::ElseKeyword,
    "Empty" => Token::EmptyKeyword,
    "End" => Token::EndKeyword,
    "Enum" => Token::EnumKeyword,
    "Eqv" => Token::EqvKeyword,
    "Erase" => Token::EraseKeyword,
    "Error" => Token::ErrorKeyword,
    "Event" => Token::EventKeyword,
    "Exit" => Token::ExitKeyword,
    "Explicit" => Token::ExplicitKeyword,
    "False" => Token::FalseKeyword,
    "FileCopy" => Token::FileCopyKeyword,
    "For" => Token::ForKeyword,
    "Friend" => Token::FriendKeyword,
    "Function" => Token::FunctionKeyword,
    "Get" => Token::GetKeyword,
    "GoSub" => Token::GoSubKeyword,
    "Goto" => Token::GotoKeyword,
    "If" => Token::IfKeyword,
    // switched so that `Imp` isn't selected for `Implements`.
    "Implements" => Token::ImplementsKeyword,
    "Imp" => Token::ImpKeyword,
    "In" => Token::InKeyword,
    "Input" => Token::InputKeyword,
    "Integer" => Token::IntegerKeyword,
    "Is" => Token::IsKeyword,
    "Kill" => Token::KillKeyword,
    "Len" => Token::LenKeyword,
    "Let" => Token::LetKeyword,
    "Lib" => Token::LibKeyword,
    "Line" => Token::LineKeyword,
    "Lock" => Token::LockKeyword,
    "Load" => Token::LoadKeyword,
    "Unload" => Token::UnloadKeyword,
    "Long" => Token::LongKeyword,
    "Loop" => Token::LoopKeyword,
    "LSet" => Token::LSetKeyword,
    "Me" => Token::MeKeyword,
    "Mid" => Token::MidKeyword,
    "MidB" => Token::MidBKeyword,
    "MkDir" => Token::MkDirKeyword,
    "Module" => Token::ModuleKeyword,
    "Mod" => Token::ModKeyword,
    "Name" => Token::NameKeyword,
    "New" => Token::NewKeyword,
    "Next" => Token::NextKeyword,
    "Not" => Token::NotKeyword,
    "Output" => Token::OutputKeyword,
    "Null" => Token::NullKeyword,
    "Object" => Token::ObjectKeyword,
    "On" => Token::OnKeyword,
    "Open" => Token::OpenKeyword,
    // Switched so that `Option` isn't selected for `Optional`.
    "Optional" => Token::OptionalKeyword,
    "Option" => Token::OptionKeyword,
    "Or" => Token::OrKeyword,
    "ParamArray" => Token::ParamArrayKeyword,
    "Preserve" => Token::PreserveKeyword,
    "Print" => Token::PrintKeyword,
    "Private" => Token::PrivateKeyword,
    "Property" => Token::PropertyKeyword,
    "Public" => Token::PublicKeyword,
    "Put" => Token::PutKeyword,
    "RaiseEvent" => Token::RaiseEventKeyword,
    "Random" => Token::RandomKeyword,
    "Randomize" => Token::RandomizeKeyword,
    "Read" => Token::ReadKeyword,
    "ReDim" => Token::ReDimKeyword,
    "Reset" => Token::ResetKeyword,
    "Resume" => Token::ResumeKeyword,
    "Return" => Token::ReturnKeyword,
    "RmDir" => Token::RmDirKeyword,
    "RSet" => Token::RSetKeyword,
    "SavePicture" => Token::SavePictureKeyword,
    "SaveSetting" => Token::SaveSettingKeyword,
    "Seek" => Token::SeekKeyword,
    "Select" => Token::SelectKeyword,
    "SendKeys" => Token::SendKeysKeyword,
    // Switched so that `Set` isn't selected for `SetAttr`.
    "SetAttr" => Token::SetAttrKeyword,
    "Set" => Token::SetKeyword,
    "Single" => Token::SingleKeyword,
    "Static" => Token::StaticKeyword,
    "Step" => Token::StepKeyword,
    "Stop" => Token::StopKeyword,
    "String" => Token::StringKeyword,
    "Sub" => Token::SubKeyword,
    "Text" => Token::TextKeyword,
    "Database" => Token::DatabaseKeyword,
    "Then" => Token::ThenKeyword,
    "Time" => Token::TimeKeyword,
    "To" => Token::ToKeyword,
    "True" => Token::TrueKeyword,
    "Type" => Token::TypeKeyword,
    "Unlock" => Token::UnlockKeyword,
    "Until" => Token::UntilKeyword,
    "Variant" => Token::VariantKeyword,
    "Version" => Token::VersionKeyword,
    "Wend" => Token::WendKeyword,
    "While" => Token::WhileKeyword,
    "Width" => Token::WidthKeyword,
    // Switched so that `With` isn't selected for `WithEvents`.
    "WithEvents" => Token::WithEventsKeyword,
    "With" => Token::WithKeyword,
    "Write" => Token::WriteKeyword,
    "Xor" => Token::XorKeyword,
};

/// Lookup table for VB6 symbols to their corresponding tokens.
/// This table is used during the tokenization process to quickly identify
/// symbols in the source code.
static SYMBOL_TOKEN_LOOKUP_TABLE: OrderedMap<&'static str, Token> = phf_ordered_map! {
    "<>" => Token::InequalityOperator,
    "<=" => Token::LessThanOrEqualOperator,
    ">=" => Token::GreaterThanOrEqualOperator,
    "=" => Token::EqualityOperator,
    "$" => Token::DollarSign,
    "_" => Token::Underscore,
    "&" => Token::Ampersand,
    "%" => Token::Percent,
    "#" => Token::Octothorpe,
    "<" => Token::LessThanOperator,
    ">" => Token::GreaterThanOperator,
    "(" => Token::LeftParenthesis,
    ")" => Token::RightParenthesis,
    "{" => Token::LeftCurlyBrace,
    "}" => Token::RightCurlyBrace,
    "," => Token::Comma,
    "+" => Token::AdditionOperator,
    "-" => Token::SubtractionOperator,
    "*" => Token::MultiplicationOperator,
    "\\" => Token::BackwardSlashOperator,
    "/" => Token::DivisionOperator,
    "." => Token::PeriodOperator,
    ":" => Token::ColonOperator,
    "^" => Token::ExponentiationOperator,
    "!" => Token::ExclamationMark,
    "[" => Token::LeftSquareBracket,
    "]" => Token::RightSquareBracket,
    ";" => Token::Semicolon,
    "@" => Token::AtSign,
};

/// Type alias for a tuple representing text and its corresponding token.
///
/// The first element of the tuple is the text slice, and the second element is the associated `Token`.
pub type TextTokenTuple<'a> = (&'a str, Token);

/// Type alias for a tuple representing a line comment and an optional newline token.
/// The first element of the tuple is another tuple containing the comment text and its corresponding token.
/// The second element is an optional tuple containing the newline text and its corresponding token.
pub type LineCommentTuple<'a> = (TextTokenTuple<'a>, Option<TextTokenTuple<'a>>);

/// Parses VB6 code into a token stream.
///
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// A vector of tuples containing the text and VB6 token type.
///
/// # Errors
///
/// If the parser encounters an unknown token, it will return an error.
///
/// # Example
///
/// ```rust
/// use vb6parse::language::Token;
/// use vb6parse::lexer::tokenize;
/// use vb6parse::io::SourceStream;
///
///
/// let mut input = SourceStream::new("test.bas", "Dim x As Integer");
/// let result = tokenize(&mut input);
///
/// let (Some(tokens), failures) = result.unpack() else {
///    panic!("Failed to read VB6 code.");
/// };
///
/// if !failures.is_empty() {
///     for failure in failures {
///         failure.print();
///     }
///
///     panic!("Failed to parse vb6 code.");
/// }
///
/// assert_eq!(tokens.len(), 7);
/// assert_eq!(tokens[0], ("Dim", Token::DimKeyword));
/// assert_eq!(tokens[1], (" ", Token::Whitespace));
/// assert_eq!(tokens[2], ("x", Token::Identifier));
/// assert_eq!(tokens[3], (" ", Token::Whitespace));
/// assert_eq!(tokens[4], ("As", Token::AsKeyword));
/// assert_eq!(tokens[5], (" ", Token::Whitespace));
/// assert_eq!(tokens[6], ("Integer", Token::IntegerKeyword));
/// ```
pub fn tokenize<'a>(
    input: &mut SourceStream<'a>,
) -> ParseResult<'a, TokenStream<'a>> {
    let mut failures = vec![];
    let mut tokens = Vec::new();

    // Always start from the beginning of the source file.
    // Some files may have already been partially parsed (eg, to extract
    // attribute statements) so we need to reset the stream since we want
    // these tokens included in the final token stream.
    input.reset_to_start();

    loop {
        if input.is_empty() {
            break;
        }

        if let Some(token) = input.take_newline() {
            tokens.push((token, Token::Newline));
            continue;
        }

        if let Some((comment_tuple, newline_optional)) = take_line_comment(input) {
            tokens.push(comment_tuple);

            if let Some(newline_tuple) = newline_optional {
                tokens.push(newline_tuple);
            }
            continue;
        }

        if let Some((comment_tuple, newline_optional)) = take_rem_comment(input) {
            tokens.push(comment_tuple);

            if let Some(newline_tuple) = newline_optional {
                tokens.push(newline_tuple);
            }
            continue;
        }

        if let Some(string_literal_tuple) = take_string_literal(input) {
            tokens.push(string_literal_tuple);
            continue;
        }

        // Try to parse date/ time literal #date# or #date time#
        if let Some((date_text, date_token)) = take_date_time_literal(input) {
            tokens.push((date_text, date_token));
            continue;
        }

        // Try to parse time only datetime literal #time#
        if let Some((date_text, date_token)) = take_time_literal(input) {
            tokens.push((date_text, date_token));
            continue;
        }

        if let Some((keyword_text, keyword_token)) = take_keyword(input) {
            tokens.push((keyword_text, keyword_token));
            continue;
        }

        if let Some((symbol_text, symbol_token)) = take_symbol(input) {
            tokens.push((symbol_text, symbol_token));
            continue;
        }

        // Try to parse numeric literal with type suffix
        if let Some((literal_text, literal_token)) = take_numeric_literal(input) {
            tokens.push((literal_text, literal_token));
            continue;
        }

        if let Some((identifier_text, identifier_token)) = take_variable_name(input) {
            tokens.push((identifier_text, identifier_token));
            continue;
        }

        if let Some(whitespace_text) = input.take_ascii_whitespaces() {
            tokens.push((whitespace_text, Token::Whitespace));
            continue;
        }

        if let Some(token_text) = input.take_count(1) {
            let error = input.generate_error(ErrorKind::UnknownToken {
                token: token_text.into(),
            });

            failures.push(error);
        }
    }

    let token_stream = TokenStream::new(input.file_name.clone(), tokens);
    (token_stream, failures).into()
}

/// Parses VB6 code into a token stream, excluding whitespace tokens.
///
/// This function first tokenizes the input, then filters out all whitespace tokens
/// from the resulting token stream.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// A `ParseResult` containing the token stream without whitespace tokens, or a list of errors.
///
/// # Errors
///
/// If the tokenizer encounters any errors, they will be included in the returned `ParseResult`.
pub fn tokenize_without_whitespaces<'a>(
    input: &mut SourceStream<'a>,
) -> ParseResult<'a, TokenStream<'a>> {
    let parse_result = tokenize(input);

    if parse_result.has_failures() {
        return parse_result;
    }

    let (token_stream_opt, failures) = parse_result.unpack();

    let Some(token_stream) = token_stream_opt else {
        return ParseResult::new(None, failures);
    };

    let tokens_without_whitespaces: Vec<(&str, Token)> = token_stream
        .tokens()
        .iter()
        .filter(|&&(_, token)| token != Token::Whitespace)
        .copied()
        .collect();

    let filtered_stream = TokenStream::new(
        token_stream.file_name().to_string(),
        tokens_without_whitespaces,
    );
    ParseResult::new(Some(filtered_stream), vec![])
}

/// Parses a VB6 to-end-of-the-line comment.
///
/// The comment starts with a single quote and continues until the end of the
/// line. It includes the single quote, but excludes the newline character(s)
/// in the comment token. If a newline exists at the end of the line (ie, it is
/// not the end of the stream) then the second token will be the newline token.
///
///
/// # Arguments
///
/// * `input` - The input to parse.
///
///
/// # Returns
///
/// `Some()` with a tuple where the first element is the comment token, including
/// the single quote while the second element is an optional newline token.
/// The only time this optional token should be None is if the line comment
/// ends at the end of the stream.
///
/// None if there is no line comment at the current position in the stream.
fn take_line_comment<'a>(input: &mut SourceStream<'a>) -> Option<LineCommentTuple<'a>> {
    input.peek_text("'", crate::io::Comparator::CaseInsensitive)?;

    match input.take_until_newline() {
        None => None,
        Some((comment, newline_optional)) => {
            let comment_tuple = (comment, Token::EndOfLineComment);

            match newline_optional {
                None => Some((comment_tuple, None)),
                Some(newline) => Some((comment_tuple, Some((newline, Token::Newline)))),
            }
        }
    }
}

/// Parses a VB6 REM-to-end-of-the-line comment.
///
/// The comment starts at the start of the line with 'REM ' and continues
/// until the end of the line. It includes the 'REM ' characters, but excludes
/// the newline character(s) in the comment token. If a newline exists at the
/// end of the line (ie, it is not the end of the stream) then the second
/// token will be the newline token.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// `Some()` with a tuple, the the first element is the comment token
/// including the 'REM ' characters at the start of the comment. The second
/// is an optional token for the newline (it's only None if the comment is
/// at the of the stream).
///
/// None if there is no REM comment at the current position in the stream.
fn take_rem_comment<'a>(input: &mut SourceStream<'a>) -> Option<LineCommentTuple<'a>> {
    input.peek_text("REM", crate::io::Comparator::CaseInsensitive)?;

    match input.take_until_newline() {
        None => None,
        Some((comment, newline_optional)) => {
            let comment_tuple = (comment, Token::RemComment);

            match newline_optional {
                None => Some((comment_tuple, None)),
                Some(newline) => Some((comment_tuple, Some((newline, Token::Newline)))),
            }
        }
    }
}

/// Parses a VB6 numeric literal with optional type suffix from the input stream.
///
/// Recognizes:
/// - Integer literals: `42%`
/// - Long literals: `42&`
/// - Single literals: `3.14!`
/// - Double literals: `3.14#`
/// - Decimal literals: `12.50@`
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Returns
///
/// `Some()` with a tuple containing the matched numeric literal text and its corresponding VB6 token
/// if a numeric literal is found at the current position in the stream; otherwise, `None`.
fn take_numeric_literal<'a>(input: &mut SourceStream<'a>) -> Option<TextTokenTuple<'a>> {
    let start_offset = input.offset;

    // Parse the numeric part (digits, optional decimal point, optional exponent)
    let _digits = input.take_ascii_digits()?;

    let mut has_decimal = false;
    let mut has_exponent = false;

    // Check for decimal point followed by more digits
    if input.peek_text(".", Comparator::CaseInsensitive).is_some() {
        // Peek ahead to see if there are digits after the period
        let _ = input.take_count(1); // consume '.'
        if input
            .peek(1)
            .and_then(|s| s.chars().next())
            .is_some_and(|c| c.is_ascii_digit())
        {
            input.take_ascii_digits(); // fractional part
            has_decimal = true;
        }
    }

    // Check for exponent (E or D followed by optional sign and digits)
    if input.peek_text("E", Comparator::CaseInsensitive).is_some()
        || input.peek_text("D", Comparator::CaseInsensitive).is_some()
    {
        let _ = input.take_count(1); // consume 'E' or 'D'
        if input.peek_text("+", Comparator::CaseInsensitive).is_some()
            || input.peek_text("-", Comparator::CaseInsensitive).is_some()
        {
            let _ = input.take_count(1); // optional sign
        }
        input.take_ascii_digits(); // exponent digits
        has_exponent = true;
    }

    // Check for type suffix
    let token_type = if input.peek_text("%", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        Token::IntegerLiteral
    } else if input.peek_text("&", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        Token::LongLiteral
    } else if input.peek_text("!", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        Token::SingleLiteral
    } else if input.peek_text("#", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        Token::DoubleLiteral
    } else if input.peek_text("@", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        Token::DecimalLiteral
    } else if has_decimal || has_exponent {
        // No explicit suffix, but has decimal point or exponent -> Single (VB6 default)
        Token::SingleLiteral
    } else {
        // No suffix, no decimal, no exponent -> Integer
        Token::IntegerLiteral
    };

    let end_offset = input.offset;
    let literal_text = &input.contents[start_offset..end_offset];

    Some((literal_text, token_type))
}

/// We only check the month digits of the stream without actually taking anything.
/// This will return the month if it's found or None if it doesn't match the correct format.
fn check_month_digits(input: &mut SourceStream) -> Option<u8> {
    // Snag the first digit of the month then check based on the digit.
    let Some(month_digit_peek) = input.peek(1) else {
        // Failed parse. This is not a date/time literal. Reset and return None.
        return None;
    };

    // Months can *not* start with a zero digit.
    if month_digit_peek == "0" {
        return None;
    }

    if month_digit_peek != "1" {
        // Likely a single digit month. Parse it and report it if it worked.
        if let Ok(single_digit_month) = str::parse::<u8>(month_digit_peek) {
            // It's between February (2) to September (9)
            return Some(single_digit_month);
        }

        return None;
    }

    // so, this could be 1 as in january, or it could be
    // october, november, or december, so we need to check
    // the next item as well.
    let Some(two_digit_month) = input.peek(2) else {
        // Failed parse. This is not a date/time literal. Reset and return None.
        return None;
    };

    if two_digit_month == "1/" {
        // looks like it might be the start of date/time literal with january as the month.
        return Some(1u8);
    }

    if let Ok(month) = str::parse::<u8>(two_digit_month) {
        // two digit month! 10, 11, or 12.
        return Some(month);
    }

    None
}

fn check_day_digits(input: &mut SourceStream) -> Option<u8> {
    // Snag the first digit of the day then check based on the digit.
    let Some(day_digit_peek) = input.peek(1) else {
        // Failed parse. This is not a date/time literal. Reset and return None.
        return None;
    };

    // days can *not* start with a zero digit.
    if day_digit_peek == "0" {
        return None;
    }

    if day_digit_peek != "1" && day_digit_peek != "2" && day_digit_peek != "3" {
        // Likely a single digit day. Parse it and report it if it worked.
        if let Ok(single_digit_day) = str::parse::<u8>(day_digit_peek) {
            // It's between 2 & 9
            return Some(single_digit_day);
        }

        return None;
    }

    // so, this could be 1 as in 1x, 2 as 2x, 3 as in 3x.
    let Some(two_digit_day) = input.peek(2) else {
        // Failed parse. This is not a date/time literal. Reset and return None.
        return None;
    };

    if two_digit_day == "1/" {
        // looks like it might be the start of date/time literal on the 1st.
        return Some(1u8);
    }

    if two_digit_day == "2/" {
        // looks like it might be the start of date/time literal on the 2nd.
        return Some(2u8);
    }

    if two_digit_day == "3/" {
        // looks like it might be the start of date/time literal on the 3rd.
        return Some(3u8);
    }

    if let Ok(day) = str::parse::<u8>(two_digit_day) {
        // If it's larger than 31, just bail out on the parsing.
        // Seriously, vb6 doesn't care if the month / day combo doesn't make sense.
        // All it cares about is that the day is 1 to 31.
        if day > 31 {
            return None;
        }

        return Some(day);
    }

    None
}

fn check_year_digits(input: &mut SourceStream) -> Option<u32> {
    // the year must be between 100 and 9999

    let four_digit_year_peek = input.peek(4)?;

    if let Ok(day) = str::parse::<u32>(four_digit_year_peek) {
        // looks like the four digit number parses so it's between 1000 and 9999
        return Some(day);
    }

    let three_digit_year_peek = input.peek(3)?;

    if let Ok(day) = str::parse::<u32>(three_digit_year_peek) {
        // looks like the three digit number parses so it's between 100 and 999
        return Some(day);
    }

    None
}

fn check_hour_digits(input: &mut SourceStream) -> Option<u8> {
    let Some(hour_double_digits) = input.peek(2) else {
        // Failed parse. This is not a date/time literal
        return None;
    };

    if let Ok(hour) = str::parse::<u8>(hour_double_digits) {
        // We got a double digit hour so now we need to be sure it's
        // betwene 12 and 1. Technically we should only have been able
        // to get 12, 11, or 10 here, but it doesn't hurt anything to expand the check.
        if (1..=12).contains(&hour) {
            return Some(hour);
        }

        return None;
    }

    // might be a single digit hour so check just a single digit.
    let Some(hour_single_digits) = input.peek(1) else {
        // Failed parse. This is not a date/time literal.
        // Technically, it shouldn't be possible to hit this
        // since we pulled *two* digits right above, but whatever.
        return None;
    };

    if let Ok(hour) = str::parse::<u8>(hour_single_digits) {
        // We got a single digit hour so now we need to be sure it's
        // betwene 12 and 1. Technically we should only have been able
        // to get 1 to 9 here, but it doesn't hurt anything to expand the check.
        if (1..=12).contains(&hour) {
            return Some(hour);
        }

        return None;
    }

    None
}

fn check_minute_or_second_digits(input: &mut SourceStream) -> Option<u8> {
    let Some(double_digits) = input.peek(2) else {
        // Failed parse. This is not a date/time literal
        return None;
    };

    if let Ok(time) = str::parse::<u8>(double_digits) {
        // We got a double digit time so now we need to be sure it's
        // less than 59.
        if time <= 59 {
            return Some(time);
        }

        return None;
    }

    None
}

/// Parses a VB6 date literal from the input stream.
///
/// Date literals are enclosed in # characters, e.g., `#1/1/2000#`, `#1/1/2000 12:30:00 PM#`
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Returns
///
/// `Some()` with a tuple containing the matched date literal text and its corresponding VB6 token
/// if a date literal is found at the current position in the stream; otherwise, `None`.
fn take_date_time_literal<'a>(input: &mut SourceStream<'a>) -> Option<TextTokenTuple<'a>> {
    let start_offset = input.offset;

    // Date format literals come in a few different formats:
    // #MM/dd/yyyy#
    // #MM/d/yyyy#
    // #M/dd/yyyy#
    // #M/d/yyyy#
    // #MM/dd/yyyy HH:mm::yyyy AM#
    // #MM/d/yyyy HH:mm::yyyy AM#
    // #M/dd/yyyy HH:mm::yyyy AM#
    // #M/d/yyyy HH:mm::yyyy AM#
    // #MM/dd/yyyy HH:mm::yyyy PM#
    // #MM/d/yyyy HH:mm::yyyy PM#
    // #M/dd/yyyy HH:mm::yyyy PM#
    // #M/d/yyyy HH:mm::yyyy PM#
    // #HH:mm::yyyy PM#
    // #H:mm::yyyy PM#
    // #HH:mm::yyyy AM#
    // #H:mm::yyyy AM#

    // Must start with #
    let Some(_) = input.take("#", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    let _month = match check_month_digits(input) {
        None => {
            // reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(month) => {
            // grab the single or double digit(s) of month
            if month >= 10 {
                let _ = input.take_count(2);
            } else {
                let _ = input.take_count(1);
            }
            month
        }
    };

    // take the day divider.
    let Some(_) = input.take("/", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    let _day = match check_day_digits(input) {
        None => {
            //reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(day) => {
            // grab the single or double digit(s) of the day
            if day >= 10 {
                let _ = input.take_count(2);
            } else {
                let _ = input.take_count(1);
            }
            day
        }
    };

    // take the year divider.
    let Some(_) = input.take("/", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    let _year = match check_year_digits(input) {
        None => {
            // reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(year) => {
            if year < 100 {
                // I don't think it's possible to get a year less than 100
                // but it's not that hard to check against it so we should.
                // reset and return since we failed the parse.
                input.offset = start_offset;
                return None;
            } else if (100..=999).contains(&year) {
                let _ = input.take_count(3);
            } else if (1000..=9999).contains(&year) {
                let _ = input.take_count(4);
            }

            year
        }
    };

    let Some(end_year_divider_peek) = input.peek(1) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };
    if end_year_divider_peek == "#" {
        // looks like it's just a date not a date/time literal.
        let _ = input.take_count(1);

        let end_offset = input.offset;
        let date_text = &input.contents[start_offset..end_offset];

        return Some((date_text, Token::DateTimeLiteral));
    }

    if end_year_divider_peek != " " {
        // This needs to be a space since it's a date & time literal.
        // Since we have something besides a " " or "#" it's not a date/time literal.

        input.offset = start_offset;
        return None;
    }

    // looks like this is a date time literal with a time section.
    // grab the space character and move on to handle the hours.
    let Some(_) = input.take(" ", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    // Grab the hours.
    let _hour = match check_hour_digits(input) {
        None => {
            // reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(hour) => {
            if hour > 12 || hour == 0 {
                // shouldn't be possible to get an hour outside the range,
                // but checking won't cause us any issue either.
                input.offset = start_offset;
                return None;
            } else if (10..=12).contains(&hour) {
                let _ = input.take_count(2);
            } else if (1..9).contains(&hour) {
                let _ = input.take_count(1);
            }

            hour
        }
    };

    let Some(end_hour_divider_peek) = input.peek(1) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };
    if end_hour_divider_peek != ":" {
        // looks like it's not a date / time parse for the hour.
        input.offset = start_offset;
        return None;
    }

    // eat the ":"
    let Some(_) = input.take(":", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    // Grab the minutes.
    let _minute = match check_minute_or_second_digits(input) {
        None => {
            // reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(minute) => {
            if minute > 59 {
                // shouldn't be possible to get a minute outside the range,
                // but checking won't cause us any issue either.
                input.offset = start_offset;
                return None;
            }

            let _ = input.take_count(2);
            minute
        }
    };

    let Some(end_minute_divider_peek) = input.peek(1) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };
    if end_minute_divider_peek != ":" {
        // looks like it's not a date / time parse for the hour.
        input.offset = start_offset;
        return None;
    }

    // eat the ":"
    let Some(_) = input.take(":", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    // Grab the seconds.
    let _seconds = match check_minute_or_second_digits(input) {
        None => {
            // reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(seconds) => {
            if seconds > 59 {
                // shouldn't be possible to get a second outside the range,
                // but checking won't cause us any issue either.
                input.offset = start_offset;
                return None;
            }

            let _ = input.take_count(2);
            seconds
        }
    };

    let Some(end_second_divider_peek) = input.peek(4) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };
    if end_second_divider_peek != " PM#" && end_second_divider_peek != " AM#" {
        // looks like it's not a date / time parse for the hour.
        input.offset = start_offset;
        return None;
    }

    // eat the " *M#"
    let Some(_) = input.take_count(4) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    let end_offset = input.offset;
    let date_text = &input.contents[start_offset..end_offset];

    Some((date_text, Token::DateTimeLiteral))
}

/// Parses a VB6 date/time literal with only a time component from the input stream.
///
/// Date literals are enclosed in # characters, e.g., `#12:30:00 PM#`
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Returns
///
/// `Some()` with a tuple containing the matched date literal text and its corresponding VB6 token
/// if a date literal is found at the current position in the stream; otherwise, `None`.
fn take_time_literal<'a>(input: &mut SourceStream<'a>) -> Option<TextTokenTuple<'a>> {
    let start_offset = input.offset;

    // Date format literals come in a few different formats:
    // #HH:mm::yyyy PM#
    // #H:mm::yyyy PM#
    // #HH:mm::yyyy AM#
    // #H:mm::yyyy AM#

    // Must start with #
    let Some(_) = input.take("#", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    // Grab the hours.
    let _hour = match check_hour_digits(input) {
        None => {
            // reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(hour) => {
            if hour > 12 || hour == 0 {
                // shouldn't be possible to get an hour outside the range,
                // but checking won't cause us any issue either.
                input.offset = start_offset;
                return None;
            } else if (10..=12).contains(&hour) {
                let _ = input.take_count(2);
            } else if (1..9).contains(&hour) {
                let _ = input.take_count(1);
            }

            hour
        }
    };

    let Some(end_hour_divider_peek) = input.peek(1) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };
    if end_hour_divider_peek != ":" {
        // looks like it's not a date / time parse for the hour.
        input.offset = start_offset;
        return None;
    }

    // eat the ":"
    let Some(_) = input.take(":", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    // Grab the minutes.
    let _minute = match check_minute_or_second_digits(input) {
        None => {
            // reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(minute) => {
            if minute > 59 {
                // shouldn't be possible to get a minute outside the range,
                // but checking won't cause us any issue either.
                input.offset = start_offset;
                return None;
            }

            let _ = input.take_count(2);
            minute
        }
    };

    let Some(end_minute_divider_peek) = input.peek(1) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };
    if end_minute_divider_peek != ":" {
        // looks like it's not a date / time parse for the hour.
        input.offset = start_offset;
        return None;
    }

    // eat the ":"
    let Some(_) = input.take(":", Comparator::CaseInsensitive) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    // Grab the seconds.
    let _seconds = match check_minute_or_second_digits(input) {
        None => {
            // reset and return since we failed the parse.
            input.offset = start_offset;
            return None;
        }
        Some(seconds) => {
            if seconds > 59 {
                // shouldn't be possible to get a second outside the range,
                // but checking won't cause us any issue either.
                input.offset = start_offset;
                return None;
            }

            let _ = input.take_count(2);
            seconds
        }
    };

    let Some(end_second_divider_peek) = input.peek(4) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };
    if end_second_divider_peek != " PM#" && end_second_divider_peek != " AM#" {
        // looks like it's not a date / time parse for the hour.
        input.offset = start_offset;
        return None;
    }

    // eat the " *M#"
    let Some(_) = input.take_count(4) else {
        // reset and return since we failed the parse.
        input.offset = start_offset;
        return None;
    };

    let end_offset = input.offset;
    let date_text = &input.contents[start_offset..end_offset];

    Some((date_text, Token::DateTimeLiteral))
}

/// Parses a VB6 string literal from the input stream.
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Returns
///
/// `Some()` with a tuple containing the matched string literal text and its corresponding VB6 token
/// if a string literal is found at the current position in the stream; otherwise, `None`.
fn take_string_literal<'a>(input: &mut SourceStream<'a>) -> Option<TextTokenTuple<'a>> {
    input.peek_text("\"", crate::io::Comparator::CaseInsensitive)?;

    let mut quote_character_count = 0;
    let take_string = |next_character| match next_character {
        // it doesn't matter what the character is if it is right after
        // the second quote character.
        '\"' if quote_character_count == 2 => {
            quote_character_count = 1;
            false
        }
        _ if quote_character_count == 2 => true,
        '\"' if quote_character_count < 2 => {
            quote_character_count += 1;
            false
        }
        _ => false,
    };

    input
        .take_until_lambda(take_string, false)
        .map(|text| (text, Token::StringLiteral))
}

/// Parses a VB6 keyword from the input stream.
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Returns
///
/// `Some()` with a tuple containing the matched keyword text and its corresponding VB6 token
/// if a keyword is found at the current position in the stream; otherwise, `None`.
fn take_keyword<'a>(input: &mut SourceStream<'a>) -> Option<TextTokenTuple<'a>> {
    for entry in KEYWORD_TOKEN_LOOKUP_TABLE.entries() {
        if let Some(matching_text) = take_matching_text(input, *entry.0) {
            return Some((matching_text, *entry.1));
        }
    }

    None
}

/// Parses a VB6 symbol from the input stream.
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Returns
///
/// `Some()` with a tuple containing the matched symbol text and its corresponding VB6 token
/// if a symbol is found at the current position in the stream; otherwise, `None`.
fn take_symbol<'a>(input: &mut SourceStream<'a>) -> Option<TextTokenTuple<'a>> {
    for entry in SYMBOL_TOKEN_LOOKUP_TABLE.entries() {
        if let Some(matching_text) = input.take(*entry.0, Comparator::CaseSensitive) {
            return Some((matching_text, *entry.1));
        }
    }

    None
}

/// Attempts to take a matching text from the input stream, ensuring that
/// the match is not part of a larger identifier.
///
/// # Arguments
///
/// * `input` - The input stream to parse.
/// * `keyword` - The keyword text to match.
///
/// # Returns
///
/// `Some()` with the matched text if it is found and not part of a larger identifier; otherwise, `None`.
pub fn take_matching_text<'a>(
    input: &mut SourceStream<'a>,
    keyword: impl Into<&'a str>,
) -> Option<&'a str> {
    let keyword_match_text = keyword.into();
    let len = keyword_match_text.len();

    let content_left_len = input.contents.len() - input.offset();
    // If we are at the end of the stream and we just so happen to match the
    // length of the keyword, we need to check if we have an exact match.
    if content_left_len == len {
        return input.take(keyword_match_text, Comparator::CaseInsensitive);
    }

    // The stream doesn't have enough characters for the keyword so we can't
    // possibly match on it.
    if content_left_len < len {
        return None;
    }

    // We already handled the case where the stream has exactly the match we
    // care about. Now we need to check in the case where the contents has
    // at least one more character than the keyword.
    //
    // We care about this last general case because we need to peek to check
    // that the last character in the match *isn't* an alphanumeric character
    // or underscore, except if that last character is a space.
    //
    // This will keep us from matching 'Timer' as the keyword 'Time' with a
    // left over of 'r' as well as keep us from matching 'char_' as 'Char'
    // with a leftover of '_'
    if content_left_len < len + 1 {
        return None;
    }

    if let Some(peek_text) = input.peek(len + 1) {
        match peek_text.chars().last() {
            None => return None,
            Some(last) => {
                if last.is_alphanumeric() || last == '_' && last != ' ' {
                    return None;
                }

                return input.take(keyword_match_text, Comparator::CaseInsensitive);
            }
        }
    }

    None
}

/// Parses a VB6 variable name (identifier) from the input stream.
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Returns
///
/// `Some()` with a tuple containing the matched identifier text and its corresponding VB6 token
/// if an identifier is found at the current position in the stream; otherwise, `None`.
fn take_variable_name<'a>(input: &mut SourceStream<'a>) -> Option<TextTokenTuple<'a>> {
    if input.peek(1)?.chars().next()?.is_ascii_alphabetic() {
        let variable_text = input.take_ascii_underscore_alphanumerics()?;

        return Some((variable_text, Token::Identifier));
    }

    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vb6_double_qoute_start_containing_string() {
        let content = r#"r = """ " 'Also a comment"#;
        let mut input = SourceStream::new("", content);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("r", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("\"\"\" \"", Token::StringLiteral));
        assert_eq!(tokens[5], (" ", Token::Whitespace));
        assert_eq!(tokens[6], ("'Also a comment", Token::EndOfLineComment));
    }

    #[test]
    fn vb6_double_qoute_mid_string() {
        let content = r#"r = " "" " 'Also a comment"#;
        let mut input = SourceStream::new("", content);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("r", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("\" \"\" \"", Token::StringLiteral));
        assert_eq!(tokens[5], (" ", Token::Whitespace));
        assert_eq!(tokens[6], ("'Also a comment", Token::EndOfLineComment));
    }

    #[test]
    fn vb6_double_qoute_end_string() {
        let content = r#"r = " """ 'Also a comment"#;
        let mut input = SourceStream::new("", content);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("r", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("\" \"\"\"", Token::StringLiteral));
        assert_eq!(tokens[5], (" ", Token::Whitespace));
        assert_eq!(tokens[6], ("'Also a comment", Token::EndOfLineComment));
    }

    #[test]
    fn vb6_double_qoute_doubled_string() {
        let content = r#"r = " "" "" " 'Also a comment"#;
        let mut input = SourceStream::new("", content);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("r", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("\" \"\" \"\" \"", Token::StringLiteral));
        assert_eq!(tokens[5], (" ", Token::Whitespace));
        assert_eq!(tokens[6], ("'Also a comment", Token::EndOfLineComment));
    }

    #[test]
    fn vb6_quad_qoute_mid_string() {
        let content = r#"r = " """" " 'Also a comment"#;
        let mut input = SourceStream::new("", content);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("r", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("\" \"\"\"\" \"", Token::StringLiteral));
        assert_eq!(tokens[5], (" ", Token::Whitespace));
        assert_eq!(tokens[6], ("'Also a comment", Token::EndOfLineComment));
    }

    #[test]
    fn vb6_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "Dim x As Integer");
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("Dim", Token::DimKeyword));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("x", Token::Identifier));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("As", Token::AsKeyword));
        assert_eq!(tokens[5], (" ", Token::Whitespace));
        assert_eq!(tokens[6], ("Integer", Token::IntegerKeyword));
        assert_eq!(tokens.len(), 7);
    }

    #[test]
    fn vb6_string_as_end_of_stream_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", r#"x = "Test""#);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens.len(), 5);
        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("\"Test\"", Token::StringLiteral));
    }

    #[test]
    fn vb6_string_at_start_of_stream_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", r#""Text""#);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0], ("\"Text\"", Token::StringLiteral));
    }

    #[test]
    fn vb6_string_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", r#"x = "Test" 'This is a comment."#);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens.len(), 7);
        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("\"Test\"", Token::StringLiteral));
        assert_eq!(tokens[5], (" ", Token::Whitespace));
        assert_eq!(tokens[6], ("'This is a comment.", Token::EndOfLineComment));
    }

    #[allow(clippy::too_many_lines)]
    #[test]
    fn class_file_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let source_code = "VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior = 0  'vbNone
    MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = \"Something\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
";

        let mut input = SourceStream::new("", source_code);
        let result = tokenize(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let mut tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens.len(), 96);
        assert_eq!(tokens.next().unwrap(), ("VERSION", Token::VersionKeyword));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("1.0", Token::SingleLiteral));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("CLASS", Token::ClassKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("BEGIN", Token::BeginKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("MultiUse", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("-", Token::SubtractionOperator));
        assert_eq!(tokens.next().unwrap(), ("1", Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("'True", Token::EndOfLineComment));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("Persistable", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("0", Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("'NotPersistable", Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("DataBindingBehavior", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("0", Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("'vbNone", Token::EndOfLineComment));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("DataSourceBehavior", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("0", Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("'vbNone", Token::EndOfLineComment));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("MTSTransactionMode", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("0", Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("'NotAnMTSObject", Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("END", Token::EndKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("VB_Name", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("\"Something\"", Token::StringLiteral)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_GlobalNameSpace", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("False", Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("VB_Creatable", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("True", Token::TrueKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_PredeclaredId", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("False", Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("VB_Exposed", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("False", Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert!(tokens.next().is_none());
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn class_file_tokenize_without_whitespace() {
        use super::tokenize_without_whitespaces;
        use crate::io::SourceStream;

        let source_code = "VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior = 0  'vbNone
    MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = \"Something\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
";

        let mut input = SourceStream::new("", source_code);
        let result = tokenize_without_whitespaces(&mut input);

        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let mut tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens.len(), 59);
        assert_eq!(tokens.next().unwrap(), ("VERSION", Token::VersionKeyword));
        assert_eq!(tokens.next().unwrap(), ("1.0", Token::SingleLiteral));
        assert_eq!(tokens.next().unwrap(), ("CLASS", Token::ClassKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("BEGIN", Token::BeginKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("MultiUse", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("-", Token::SubtractionOperator));
        assert_eq!(tokens.next().unwrap(), ("1", Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("'True", Token::EndOfLineComment));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("Persistable", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("0", Token::IntegerLiteral));
        assert_eq!(
            tokens.next().unwrap(),
            ("'NotPersistable", Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("DataBindingBehavior", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("0", Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("'vbNone", Token::EndOfLineComment));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("DataSourceBehavior", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("0", Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("'vbNone", Token::EndOfLineComment));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("MTSTransactionMode", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("0", Token::IntegerLiteral));
        assert_eq!(
            tokens.next().unwrap(),
            ("'NotAnMTSObject", Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("END", Token::EndKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), ("VB_Name", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(
            tokens.next().unwrap(),
            ("\"Something\"", Token::StringLiteral)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_GlobalNameSpace", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("False", Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), ("VB_Creatable", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("True", Token::TrueKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_PredeclaredId", Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("False", Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), ("VB_Exposed", Token::Identifier));
        assert_eq!(tokens.next().unwrap(), ("=", Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("False", Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", Token::Newline));
        assert!(tokens.next().is_none());
    }

    #[test]
    fn integer_literal_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "x = 42%");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("42%", Token::IntegerLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn long_literal_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "x = 123456&");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("123456&", Token::LongLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn single_literal_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "x = 3.14!");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("3.14!", Token::SingleLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn double_literal_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "x = 3.14159265#");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("3.14159265#", Token::DoubleLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn decimal_literal_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "price = 12.50@");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("price", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("12.50@", Token::DecimalLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn date_literal_only_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "d = #1/1/2000#");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("d", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("#1/1/2000#", Token::DateTimeLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn date_literal_with_time_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "d = #12/31/1999 11:59:59 PM#");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("d", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(
            tokens[4],
            ("#12/31/1999 11:59:59 PM#", Token::DateTimeLiteral)
        );
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn date_literal_with_am_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "d = #12/31/1999 11:59:59 AM#");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("d", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(
            tokens[4],
            ("#12/31/1999 11:59:59 AM#", Token::DateTimeLiteral)
        );
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn date_literal_single_month_with_am_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "d = #1/15/2000 10:30:45 AM#");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("d", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(
            tokens[4],
            ("#1/15/2000 10:30:45 AM#", Token::DateTimeLiteral)
        );
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn time_only_literal_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "d = #10:20:45 AM#");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("d", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("#10:20:45 AM#", Token::DateTimeLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn date_literal_with_pm_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "d = #1/1/100 1:00:00 PM#");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("d", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("#1/1/100 1:00:00 PM#", Token::DateTimeLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn date_literal_with_largest_time_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "d = #12/31/9999 12:59:59 PM#");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("d", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(
            tokens[4],
            ("#12/31/9999 12:59:59 PM#", Token::DateTimeLiteral)
        );
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn plain_number_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "x = 42");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("42", Token::IntegerLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn decimal_number_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "x = 3.14");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("3.14", Token::SingleLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn exponent_number_tokenize() {
        use super::tokenize;
        use crate::io::SourceStream;

        let mut input = SourceStream::new("", "x = 1.5E+10");
        let result = tokenize(&mut input);
        let (tokens_opt, failures) = result.unpack();

        if !failures.is_empty() {
            for failure in failures {
                failure.eprint();
            }
        }

        let tokens = tokens_opt.expect("Expected tokens");

        assert_eq!(tokens[0], ("x", Token::Identifier));
        assert_eq!(tokens[1], (" ", Token::Whitespace));
        assert_eq!(tokens[2], ("=", Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", Token::Whitespace));
        assert_eq!(tokens[4], ("1.5E+10", Token::SingleLiteral));
        assert_eq!(tokens.len(), 5);
    }
}
