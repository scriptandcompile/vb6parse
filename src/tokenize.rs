use phf::{phf_ordered_map, OrderedMap};

use crate::{
    language::VB6Token,
    parsers::{Comparator, ParseResult, SourceStream},
    tokenstream::TokenStream,
    VB6CodeErrorKind,
};

static KEYWORD_TOKEN_LOOKUP_TABLE: OrderedMap<&'static str, VB6Token> = phf_ordered_map! {
    "AdressOf" => VB6Token::AddressOfKeyword,
    "Access" => VB6Token::AccessKeyword,
    "Alias" => VB6Token::AliasKeyword,
    "And" => VB6Token::AndKeyword,
    "AppActivate" => VB6Token::AppActivateKeyword,
    "Append" => VB6Token::AppendKeyword,
    "Attribute" => VB6Token::AttributeKeyword,
    "As" => VB6Token::AsKeyword,
    "Base" => VB6Token::BaseKeyword,
    "Beep" => VB6Token::BeepKeyword,
    "Begin" => VB6Token::BeginKeyword,
    "Binary" => VB6Token::BinaryKeyword,
    "Boolean" => VB6Token::BooleanKeyword,
    "ByRef" => VB6Token::ByRefKeyword,
    "Byte" => VB6Token::ByteKeyword,
    "ByVal" => VB6Token::ByValKeyword,
    "Call" => VB6Token::CallKeyword,
    "Case" => VB6Token::CaseKeyword,
    "ChDir" => VB6Token::ChDirKeyword,
    "ChDrive" => VB6Token::ChDriveKeyword,
    "Class" => VB6Token::ClassKeyword,
    "Close" => VB6Token::CloseKeyword,
    "Compare" => VB6Token::CompareKeyword,
    "Const" => VB6Token::ConstKeyword,
    "Currency" => VB6Token::CurrencyKeyword,
    "Date" => VB6Token::DateKeyword,
    "Decimal" => VB6Token::DecimalKeyword,
    "Declare" => VB6Token::DeclareKeyword,
    "DefBool" => VB6Token::DefBoolKeyword,
    "DefByte" => VB6Token::DefByteKeyword,
    "DefCur" => VB6Token::DefCurKeyword,
    "DefDate" => VB6Token::DefDateKeyword,
    "DefDbl" => VB6Token::DefDblKeyword,
    "DefDec" => VB6Token::DefDecKeyword,
    "DefInt" => VB6Token::DefIntKeyword,
    "DefLng" => VB6Token::DefLngKeyword,
    "DefObj" => VB6Token::DefObjKeyword,
    "DefSng" => VB6Token::DefSngKeyword,
    "DefStr" => VB6Token::DefStrKeyword,
    "DefVar" => VB6Token::DefVarKeyword,
    "DeleteSetting" => VB6Token::DeleteSettingKeyword,
    "Dim" => VB6Token::DimKeyword,
    // switched so that `Do` isn't selected for `Double`.
    "Double" => VB6Token::DoubleKeyword,
    "Do" => VB6Token::DoKeyword,
    "Each" => VB6Token::EachKeyword,
    // switched so that `Else` isn't selected for `ElseIf`.
    "ElseIf" => VB6Token::ElseIfKeyword,
    "Else" => VB6Token::ElseKeyword,
    "Empty" => VB6Token::EmptyKeyword,
    "End" => VB6Token::EndKeyword,
    "Enum" => VB6Token::EnumKeyword,
    "Eqv" => VB6Token::EqvKeyword,
    "Erase" => VB6Token::EraseKeyword,
    "Error" => VB6Token::ErrorKeyword,
    "Event" => VB6Token::EventKeyword,
    "Exit" => VB6Token::ExitKeyword,
    "Explicit" => VB6Token::ExplicitKeyword,
    "False" => VB6Token::FalseKeyword,
    "FileCopy" => VB6Token::FileCopyKeyword,
    "For" => VB6Token::ForKeyword,
    "Friend" => VB6Token::FriendKeyword,
    "Function" => VB6Token::FunctionKeyword,
    "Get" => VB6Token::GetKeyword,
    "GoSub" => VB6Token::GoSubKeyword,
    "Goto" => VB6Token::GotoKeyword,
    "If" => VB6Token::IfKeyword,
    // switched so that `Imp` isn't selected for `Implements`.
    "Implements" => VB6Token::ImplementsKeyword,
    "Imp" => VB6Token::ImpKeyword,
    "In" => VB6Token::InKeyword,
    "Input" => VB6Token::InputKeyword,
    "Integer" => VB6Token::IntegerKeyword,
    "Is" => VB6Token::IsKeyword,
    "Kill" => VB6Token::KillKeyword,
    "Len" => VB6Token::LenKeyword,
    "Let" => VB6Token::LetKeyword,
    "Lib" => VB6Token::LibKeyword,
    "Line" => VB6Token::LineKeyword,
    "Lock" => VB6Token::LockKeyword,
    "Load" => VB6Token::LoadKeyword,
    "Unload" => VB6Token::UnloadKeyword,
    "Long" => VB6Token::LongKeyword,
    "Loop" => VB6Token::LoopKeyword,
    "LSet" => VB6Token::LSetKeyword,
    "Me" => VB6Token::MeKeyword,
    "Mid" => VB6Token::MidKeyword,
    "MidB" => VB6Token::MidBKeyword,
    "MkDir" => VB6Token::MkDirKeyword,
    "Module" => VB6Token::ModuleKeyword,
    "Mod" => VB6Token::ModKeyword,
    "Name" => VB6Token::NameKeyword,
    "New" => VB6Token::NewKeyword,
    "Next" => VB6Token::NextKeyword,
    "Not" => VB6Token::NotKeyword,
    "Output" => VB6Token::OutputKeyword,
    "Null" => VB6Token::NullKeyword,
    "Object" => VB6Token::ObjectKeyword,
    "On" => VB6Token::OnKeyword,
    "Open" => VB6Token::OpenKeyword,
    // Switched so that `Option` isn't selected for `Optional`.
    "Optional" => VB6Token::OptionalKeyword,
    "Option" => VB6Token::OptionKeyword,
    "Or" => VB6Token::OrKeyword,
    "ParamArray" => VB6Token::ParamArrayKeyword,
    "Preserve" => VB6Token::PreserveKeyword,
    "Print" => VB6Token::PrintKeyword,
    "Private" => VB6Token::PrivateKeyword,
    "Property" => VB6Token::PropertyKeyword,
    "Public" => VB6Token::PublicKeyword,
    "Put" => VB6Token::PutKeyword,
    "RaiseEvent" => VB6Token::RaiseEventKeyword,
    "Random" => VB6Token::RandomKeyword,
    "Randomize" => VB6Token::RandomizeKeyword,
    "Read" => VB6Token::ReadKeyword,
    "ReDim" => VB6Token::ReDimKeyword,
    "Reset" => VB6Token::ResetKeyword,
    "Resume" => VB6Token::ResumeKeyword,
    "Return" => VB6Token::ReturnKeyword,
    "RmDir" => VB6Token::RmDirKeyword,
    "RSet" => VB6Token::RSetKeyword,
    "SavePicture" => VB6Token::SavePictureKeyword,
    "SaveSetting" => VB6Token::SaveSettingKeyword,
    "Seek" => VB6Token::SeekKeyword,
    "Select" => VB6Token::SelectKeyword,
    "SendKeys" => VB6Token::SendKeysKeyword,
    // Switched so that `Set` isn't selected for `SetAttr`.
    "SetAttr" => VB6Token::SetAttrKeyword,
    "Set" => VB6Token::SetKeyword,
    "Single" => VB6Token::SingleKeyword,
    "Static" => VB6Token::StaticKeyword,
    "Step" => VB6Token::StepKeyword,
    "Stop" => VB6Token::StopKeyword,
    "String" => VB6Token::StringKeyword,
    "Sub" => VB6Token::SubKeyword,
    "Text" => VB6Token::TextKeyword,
    "Database" => VB6Token::DatabaseKeyword,
    "Then" => VB6Token::ThenKeyword,
    "Time" => VB6Token::TimeKeyword,
    "To" => VB6Token::ToKeyword,
    "True" => VB6Token::TrueKeyword,
    "Type" => VB6Token::TypeKeyword,
    "Unlock" => VB6Token::UnlockKeyword,
    "Until" => VB6Token::UntilKeyword,
    "Variant" => VB6Token::VariantKeyword,
    "Version" => VB6Token::VersionKeyword,
    "Wend" => VB6Token::WendKeyword,
    "While" => VB6Token::WhileKeyword,
    "Width" => VB6Token::WidthKeyword,
    // Switched so that `With` isn't selected for `WithEvents`.
    "WithEvents" => VB6Token::WithEventsKeyword,
    "With" => VB6Token::WithKeyword,
    "Write" => VB6Token::WriteKeyword,
    "Xor" => VB6Token::XorKeyword,
};

static SYMBOL_TOKEN_LOOKUP_TABLE: OrderedMap<&'static str, VB6Token> = phf_ordered_map! {
    "<>" => VB6Token::InequalityOperator,
    "<=" => VB6Token::LessThanOrEqualOperator,
    ">=" => VB6Token::GreaterThanOrEqualOperator,
    "=" => VB6Token::EqualityOperator,
    "$" => VB6Token::DollarSign,
    "_" => VB6Token::Underscore,
    "&" => VB6Token::Ampersand,
    "%" => VB6Token::Percent,
    "#" => VB6Token::Octothorpe,
    "<" => VB6Token::LessThanOperator,
    ">" => VB6Token::GreaterThanOperator,
    "(" => VB6Token::LeftParenthesis,
    ")" => VB6Token::RightParenthesis,
    "," => VB6Token::Comma,
    "+" => VB6Token::AdditionOperator,
    "-" => VB6Token::SubtractionOperator,
    "*" => VB6Token::MultiplicationOperator,
    "\\" => VB6Token::BackwardSlashOperator,
    "/" => VB6Token::DivisionOperator,
    "." => VB6Token::PeriodOperator,
    ":" => VB6Token::ColonOperator,
    "^" => VB6Token::ExponentiationOperator,
    "!" => VB6Token::ExclamationMark,
    "[" => VB6Token::LeftSquareBracket,
    "]" => VB6Token::RightSquareBracket,
    ";" => VB6Token::Semicolon,
    "@" => VB6Token::AtSign,
};

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
/// use vb6parse::language::VB6Token;
/// use vb6parse::tokenize::tokenize;
/// use vb6parse::SourceStream;
///
///
/// let mut input = SourceStream::new("test.bas", "Dim x As Integer");
/// let result = tokenize(&mut input);
///
/// if result.has_failures() {
///     for failure in result.failures {
///         failure.print();
///     }
///     panic!("Failed to parse vb6 code.");
/// }
///
/// let tokens = result.unwrap();
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
pub fn tokenize<'a>(
    input: &mut SourceStream<'a>,
) -> ParseResult<'a, TokenStream<'a>, VB6CodeErrorKind> {
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
            tokens.push((token, VB6Token::Newline));
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

        // Try to parse date literal first (#date#)
        if let Some((date_text, date_token)) = take_date_literal(input) {
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
            tokens.push((whitespace_text, VB6Token::Whitespace));
            continue;
        }

        if let Some(token_text) = input.take_count(1) {
            let error = input.generate_error(VB6CodeErrorKind::UnknownToken {
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
) -> ParseResult<'a, TokenStream<'a>, VB6CodeErrorKind> {
    let parse_result = tokenize(input);

    if parse_result.has_failures() {
        return parse_result;
    }

    let token_stream = parse_result.result.unwrap();
    let tokens_without_whitespaces: Vec<(&str, VB6Token)> = token_stream
        .tokens
        .into_iter()
        .filter(|(_, token)| !matches!(token, VB6Token::Whitespace))
        .collect();

    let filtered_stream = TokenStream::new(token_stream.source_file, tokens_without_whitespaces);
    ParseResult {
        result: Some(filtered_stream),
        failures: vec![],
    }
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
fn take_line_comment<'a>(
    input: &mut SourceStream<'a>,
) -> Option<((&'a str, VB6Token), Option<(&'a str, VB6Token)>)> {
    input.peek_text("'", super::Comparator::CaseInsensitive)?;

    match input.take_until_newline() {
        None => None,
        Some((comment, newline_optional)) => {
            let comment_tuple = (comment, VB6Token::EndOfLineComment);

            match newline_optional {
                None => Some((comment_tuple, None)),
                Some(newline) => Some((comment_tuple, Some((newline, VB6Token::Newline)))),
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
fn take_rem_comment<'a>(
    input: &mut SourceStream<'a>,
) -> Option<((&'a str, VB6Token), Option<(&'a str, VB6Token)>)> {
    input.peek_text("REM", super::Comparator::CaseInsensitive)?;

    match input.take_until_newline() {
        None => None,
        Some((comment, newline_optional)) => {
            let comment_tuple = (comment, VB6Token::RemComment);

            match newline_optional {
                None => Some((comment_tuple, None)),
                Some(newline) => Some((comment_tuple, Some((newline, VB6Token::Newline)))),
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
fn take_numeric_literal<'a>(input: &mut SourceStream<'a>) -> Option<(&'a str, VB6Token)> {
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
        VB6Token::IntegerLiteral
    } else if input.peek_text("&", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        VB6Token::LongLiteral
    } else if input.peek_text("!", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        VB6Token::SingleLiteral
    } else if input.peek_text("#", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        VB6Token::DoubleLiteral
    } else if input.peek_text("@", Comparator::CaseInsensitive).is_some() {
        let _ = input.take_count(1);
        VB6Token::DecimalLiteral
    } else if has_decimal || has_exponent {
        // No explicit suffix, but has decimal point or exponent -> Single (VB6 default)
        VB6Token::SingleLiteral
    } else {
        // No suffix, no decimal, no exponent -> Integer
        VB6Token::IntegerLiteral
    };

    let end_offset = input.offset;
    let literal_text = &input.contents[start_offset..end_offset];

    Some((literal_text, token_type))
}

/// Parses a VB6 date literal from the input stream.
///
/// Date literals are enclosed in # characters, e.g., `#1/1/2000#`, `#12:30:00 PM#`
///
/// # Arguments
///
/// * `input` - The input stream to parse.
///
/// # Returns
///
/// `Some()` with a tuple containing the matched date literal text and its corresponding VB6 token
/// if a date literal is found at the current position in the stream; otherwise, `None`.
fn take_date_literal<'a>(input: &mut SourceStream<'a>) -> Option<(&'a str, VB6Token)> {
    let start_offset = input.offset;

    // Must start with #
    input.peek_text("#", Comparator::CaseInsensitive)?;
    let _ = input.take_count(1);

    // Take everything until the closing # or newline
    let found_closing = loop {
        if let Some(peek) = input.peek(1) {
            let ch = peek.chars().next()?;
            if ch == '#' {
                let _ = input.take_count(1);
                break true;
            } else if ch == '\r' || ch == '\n' {
                // Date literals cannot span multiple lines
                break false;
            }

            let _ = input.take_count(1);
        } else {
            // End of input
            break false;
        }
    };

    if !found_closing {
        // Rollback if we didn't find a closing #
        input.offset = start_offset;
        return None;
    }

    let end_offset = input.offset;
    let date_text = &input.contents[start_offset..end_offset];

    Some((date_text, VB6Token::DateLiteral))
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
fn take_string_literal<'a>(input: &mut SourceStream<'a>) -> Option<(&'a str, VB6Token)> {
    input.peek_text("\"", super::Comparator::CaseInsensitive)?;

    // TODO: Need to handle error reporting of incorrect escape sequences as well
    // as string literals that hit a newline before the second quote character.
    let mut escape_sequence_started = false;
    let mut quote_character_count = 0;
    let take_string = |next_character| match next_character {
        // it doesn't matter what the character is if it is right after
        // the second quote character.
        _ if quote_character_count == 2 => true,
        '\r' | '\n' => true,
        '\\' => {
            escape_sequence_started = true;
            false
        }
        '\"' if !escape_sequence_started && quote_character_count < 2 => {
            quote_character_count += 1;
            false
        }
        _ if escape_sequence_started => {
            escape_sequence_started = false;
            false
        }
        _ => false,
    };

    input
        .take_until_lambda(take_string, false)
        .map(|text| (text, VB6Token::StringLiteral))
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
fn take_keyword<'a>(input: &mut SourceStream<'a>) -> Option<(&'a str, VB6Token)> {
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
fn take_symbol<'a>(input: &mut SourceStream<'a>) -> Option<(&'a str, VB6Token)> {
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
fn take_variable_name<'a>(input: &mut SourceStream<'a>) -> Option<(&'a str, VB6Token)> {
    if input.peek(1)?.chars().next()?.is_ascii_alphabetic() {
        let variable_text = input.take_ascii_underscore_alphanumerics()?;

        return Some((variable_text, VB6Token::Identifier));
    }

    None
}

#[cfg(test)]
mod test {
    use super::*;

    #[test]
    fn vb6_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "Dim x As Integer");
        let result = tokenize(&mut input);

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("Dim", VB6Token::DimKeyword));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("x", VB6Token::Identifier));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("As", VB6Token::AsKeyword));
        assert_eq!(tokens[5], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[6], ("Integer", VB6Token::IntegerKeyword));
        assert_eq!(tokens.len(), 7);
    }

    #[test]
    fn vb6_string_as_end_of_stream_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", r#"x = "Test""#);
        let result = tokenize(&mut input);

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let tokens = result.result.unwrap();

        assert_eq!(tokens.len(), 5);
        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("\"Test\"", VB6Token::StringLiteral));
    }

    #[test]
    fn vb6_string_at_start_of_stream_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", r#""Text""#);
        let result = tokenize(&mut input);

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let tokens = result.result.unwrap();

        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0], ("\"Text\"", VB6Token::StringLiteral));
    }

    #[test]
    fn vb6_string_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", r#"x = "Test" 'This is a comment."#);
        let result = tokenize(&mut input);

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let tokens = result.result.unwrap();

        assert_eq!(tokens.len(), 7);
        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("\"Test\"", VB6Token::StringLiteral));
        assert_eq!(tokens[5], (" ", VB6Token::Whitespace));
        assert_eq!(
            tokens[6],
            ("'This is a comment.", VB6Token::EndOfLineComment)
        );
    }

    #[test]
    fn class_file_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

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

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let mut tokens = result.result.unwrap().into_iter();

        assert_eq!(tokens.len(), 96);
        assert_eq!(
            tokens.next().unwrap(),
            ("VERSION", VB6Token::VersionKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("1.0", VB6Token::SingleLiteral));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("CLASS", VB6Token::ClassKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("BEGIN", VB6Token::BeginKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("MultiUse", VB6Token::Identifier));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("-", VB6Token::SubtractionOperator));
        assert_eq!(tokens.next().unwrap(), ("1", VB6Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("'True", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("Persistable", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("0", VB6Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("'NotPersistable", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("DataBindingBehavior", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("0", VB6Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("'vbNone", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("DataSourceBehavior", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("0", VB6Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("'vbNone", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("    ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("MTSTransactionMode", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("0", VB6Token::IntegerLiteral));
        assert_eq!(tokens.next().unwrap(), ("  ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("'NotAnMTSObject", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("END", VB6Token::EndKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("VB_Name", VB6Token::Identifier));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("\"Something\"", VB6Token::StringLiteral)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_GlobalNameSpace", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("False", VB6Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_Creatable", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("True", VB6Token::TrueKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_PredeclaredId", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("False", VB6Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("VB_Exposed", VB6Token::Identifier));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), (" ", VB6Token::Whitespace));
        assert_eq!(tokens.next().unwrap(), ("False", VB6Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert!(tokens.next().is_none());
    }

    #[test]
    fn class_file_tokenize_without_whitespace() {
        use crate::tokenize::tokenize_without_whitespaces;
        use crate::SourceStream;

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

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let mut tokens = result.result.unwrap().into_iter();

        assert_eq!(tokens.len(), 59);
        assert_eq!(
            tokens.next().unwrap(),
            ("VERSION", VB6Token::VersionKeyword)
        );
        assert_eq!(tokens.next().unwrap(), ("1.0", VB6Token::SingleLiteral));
        assert_eq!(tokens.next().unwrap(), ("CLASS", VB6Token::ClassKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("BEGIN", VB6Token::BeginKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("MultiUse", VB6Token::Identifier));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("-", VB6Token::SubtractionOperator));
        assert_eq!(tokens.next().unwrap(), ("1", VB6Token::IntegerLiteral));
        assert_eq!(
            tokens.next().unwrap(),
            ("'True", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Persistable", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("0", VB6Token::IntegerLiteral));
        assert_eq!(
            tokens.next().unwrap(),
            ("'NotPersistable", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("DataBindingBehavior", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("0", VB6Token::IntegerLiteral));
        assert_eq!(
            tokens.next().unwrap(),
            ("'vbNone", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("DataSourceBehavior", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("0", VB6Token::IntegerLiteral));
        assert_eq!(
            tokens.next().unwrap(),
            ("'vbNone", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("MTSTransactionMode", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("0", VB6Token::IntegerLiteral));
        assert_eq!(
            tokens.next().unwrap(),
            ("'NotAnMTSObject", VB6Token::EndOfLineComment)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(tokens.next().unwrap(), ("END", VB6Token::EndKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), ("VB_Name", VB6Token::Identifier));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(
            tokens.next().unwrap(),
            ("\"Something\"", VB6Token::StringLiteral)
        );
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_GlobalNameSpace", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("False", VB6Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_Creatable", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("True", VB6Token::TrueKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(
            tokens.next().unwrap(),
            ("VB_PredeclaredId", VB6Token::Identifier)
        );
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("False", VB6Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert_eq!(
            tokens.next().unwrap(),
            ("Attribute", VB6Token::AttributeKeyword)
        );
        assert_eq!(tokens.next().unwrap(), ("VB_Exposed", VB6Token::Identifier));
        assert_eq!(tokens.next().unwrap(), ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens.next().unwrap(), ("False", VB6Token::FalseKeyword));
        assert_eq!(tokens.next().unwrap(), ("\n", VB6Token::Newline));
        assert!(tokens.next().is_none());
    }

    #[test]
    fn integer_literal_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "x = 42%");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("42%", VB6Token::IntegerLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn long_literal_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "x = 123456&");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("123456&", VB6Token::LongLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn single_literal_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "x = 3.14!");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("3.14!", VB6Token::SingleLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn double_literal_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "x = 3.14159265#");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("3.14159265#", VB6Token::DoubleLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn decimal_literal_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "price = 12.50@");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("price", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("12.50@", VB6Token::DecimalLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn date_literal_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "d = #1/1/2000#");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("d", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("#1/1/2000#", VB6Token::DateLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn date_literal_with_time_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "d = #12/31/1999 11:59:59 PM#");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("d", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(
            tokens[4],
            ("#12/31/1999 11:59:59 PM#", VB6Token::DateLiteral)
        );
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn plain_number_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "x = 42");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("42", VB6Token::IntegerLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn decimal_number_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "x = 3.14");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("3.14", VB6Token::SingleLiteral));
        assert_eq!(tokens.len(), 5);
    }

    #[test]
    fn exponent_number_tokenize() {
        use crate::tokenize::tokenize;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "x = 1.5E+10");
        let result = tokenize(&mut input);

        assert!(!result.has_failures());
        let tokens = result.result.unwrap();

        assert_eq!(tokens[0], ("x", VB6Token::Identifier));
        assert_eq!(tokens[1], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[2], ("=", VB6Token::EqualityOperator));
        assert_eq!(tokens[3], (" ", VB6Token::Whitespace));
        assert_eq!(tokens[4], ("1.5E+10", VB6Token::SingleLiteral));
        assert_eq!(tokens.len(), 5);
    }
}
