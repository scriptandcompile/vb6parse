use phf::{OrderedMap, phf_ordered_map};

use crate::{
    language::VB6Token,
    parsers::{Comparator, ParseResult, SourceStream},
    VB6CodeErrorKind,
};

pub trait VB6Tokenizer<'a> {
    fn take_line_comment(self) -> Option<(VB6Token<'a>, Option<VB6Token<'a>>)>;
    fn take_rem_comment(self) -> Option<(VB6Token<'a>, Option<VB6Token<'a>>)>;
    fn take_string_literal(self) -> Option<VB6Token<'a>>;
    fn take_keyword(self) -> Option<VB6Token<'a>>;
    fn take_matching_text(self, keyword: impl Into<&'a str>) -> Option<&'a str>;
    fn take_symbol(self) -> Option<VB6Token<'a>>;
    fn take_variable_name(self) -> Option<VB6Token<'a>>;
}

static KEYWORD_TOKEN_LOOKUP_TABLE: OrderedMap<&'static str, for<'a> fn(&'a str) -> VB6Token<'a>> = phf_ordered_map! {
    "AdressOf" => |matching_text| VB6Token::AddressOfKeyword(matching_text),
    "Alias" => |matching_text| VB6Token::AliasKeyword(matching_text),
    "And" => |matching_text| VB6Token::AndKeyword(matching_text),
    "AppActivate" => |matching_text| VB6Token::AppActivateKeyword(matching_text),
    "As" => |matching_text| VB6Token::AsKeyword(matching_text),
    "Base" => |matching_text| VB6Token::BaseKeyword(matching_text),
    "Beep" => |matching_text| VB6Token::BeepKeyword(matching_text),
    "Binary" => |matching_text| VB6Token::BinaryKeyword(matching_text),
    "Boolean" => |matching_text| VB6Token::BooleanKeyword(matching_text),
    "ByRef" => |matching_text| VB6Token::ByRefKeyword(matching_text),
    "Byte" => |matching_text| VB6Token::ByteKeyword(matching_text),
    "ByVal" => |matching_text| VB6Token::ByValKeyword(matching_text),
    "Call" => |matching_text| VB6Token::CallKeyword(matching_text),
    "Case" => |matching_text| VB6Token::CaseKeyword(matching_text),
    "ChDir" => |matching_text| VB6Token::ChDirKeyword(matching_text),
    "ChDrive" => |matching_text| VB6Token::ChDriveKeyword(matching_text),
    "Close" => |matching_text| VB6Token::CloseKeyword(matching_text),
    "Compare" => |matching_text| VB6Token::CompareKeyword(matching_text),
    "Const" => |matching_text| VB6Token::ConstKeyword(matching_text),
    "Currency" => |matching_text| VB6Token::CurrencyKeyword(matching_text),
    "Date" => |matching_text| VB6Token::DateKeyword(matching_text),
    "Decimal" => |matching_text| VB6Token::DecimalKeyword(matching_text),
    "Declare" => |matching_text| VB6Token::DeclareKeyword(matching_text),
    "DefBool" => |matching_text| VB6Token::DefBoolKeyword(matching_text),
    "DefByte" => |matching_text| VB6Token::DefByteKeyword(matching_text),
    "DefCur" => |matching_text| VB6Token::DefCurKeyword(matching_text),
    "DefDate" => |matching_text| VB6Token::DefDateKeyword(matching_text),
    "DefDbl" => |matching_text| VB6Token::DefDblKeyword(matching_text),
    "DefDec" => |matching_text| VB6Token::DefDecKeyword(matching_text),
    "DefInt" => |matching_text| VB6Token::DefIntKeyword(matching_text),
    "DefLng" => |matching_text| VB6Token::DefLngKeyword(matching_text),
    "DefObj" => |matching_text| VB6Token::DefObjKeyword(matching_text),
    "DefSng" => |matching_text| VB6Token::DefSngKeyword(matching_text),
    "DefStr" => |matching_text| VB6Token::DefStrKeyword(matching_text),
    "DefVar" => |matching_text| VB6Token::DefVarKeyword(matching_text),
    "DeleteSetting" => |matching_text| VB6Token::DeleteSettingKeyword(matching_text),
    "Dim" => |matching_text| VB6Token::DimKeyword(matching_text),
    // switched so that `Do` isn't selected for `Double`.
    "Double" => |matching_text| VB6Token::DoubleKeyword(matching_text),
    "Do" => |matching_text| VB6Token::DoKeyword(matching_text),
    // switched so that `Else` isn't selected for `ElseIf`.
    "ElseIf" => |matching_text| VB6Token::ElseIfKeyword(matching_text),
    "Else" => |matching_text| VB6Token::ElseKeyword(matching_text),
    "Empty" => |matching_text| VB6Token::EmptyKeyword(matching_text),
    "End" => |matching_text| VB6Token::EndKeyword(matching_text),
    "Enum" => |matching_text| VB6Token::EnumKeyword(matching_text),
    "Eqv" => |matching_text| VB6Token::EqvKeyword(matching_text),
    "Erase" => |matching_text| VB6Token::EraseKeyword(matching_text),
    "Error" => |matching_text| VB6Token::ErrorKeyword(matching_text),
    "Event" => |matching_text| VB6Token::EventKeyword(matching_text),
    "Exit" => |matching_text| VB6Token::ExitKeyword(matching_text),
    "Explicit" => |matching_text| VB6Token::ExplicitKeyword(matching_text),
    "False" => |matching_text| VB6Token::FalseKeyword(matching_text),
    "FileCopy" => |matching_text| VB6Token::FileCopyKeyword(matching_text),
    "For" => |matching_text| VB6Token::ForKeyword(matching_text),
    "Friend" => |matching_text| VB6Token::FriendKeyword(matching_text),
    "Function" => |matching_text| VB6Token::FunctionKeyword(matching_text),
    "Get" => |matching_text| VB6Token::GetKeyword(matching_text),
    "Goto" => |matching_text| VB6Token::GotoKeyword(matching_text),
    "If" => |matching_text| VB6Token::IfKeyword(matching_text),
    // switched so that `Imp` isn't selected for `Implements`.
    "Implements" => |matching_text| VB6Token::ImplementsKeyword(matching_text),
    "Imp" => |matching_text| VB6Token::ImpKeyword(matching_text),
    "Integer" => |matching_text| VB6Token::IntegerKeyword(matching_text),
    "Is" => |matching_text| VB6Token::IsKeyword(matching_text),
    "Kill" => |matching_text| VB6Token::KillKeyword(matching_text),
    "Len" => |matching_text| VB6Token::LenKeyword(matching_text),
    "Let" => |matching_text| VB6Token::LetKeyword(matching_text),
    "Lib" => |matching_text| VB6Token::LibKeyword(matching_text),
    "Line" => |matching_text| VB6Token::LineKeyword(matching_text),
    "Lock" => |matching_text| VB6Token::LockKeyword(matching_text),
    "Load" => |matching_text| VB6Token::LoadKeyword(matching_text),
    "Long" => |matching_text| VB6Token::LongKeyword(matching_text),
    "LSet" => |matching_text| VB6Token::LSetKeyword(matching_text),
    "Me" => |matching_text| VB6Token::MeKeyword(matching_text),
    "Mid" => |matching_text| VB6Token::MidKeyword(matching_text),
    "MkDir" => |matching_text| VB6Token::MkDirKeyword(matching_text),
    "Mod" => |matching_text| VB6Token::ModKeyword(matching_text),
    "Name" => |matching_text| VB6Token::NameKeyword(matching_text),
    "New" => |matching_text| VB6Token::NewKeyword(matching_text),
    "Next" => |matching_text| VB6Token::NextKeyword(matching_text),
    "Not" => |matching_text| VB6Token::NotKeyword(matching_text),
    "Null" => |matching_text| VB6Token::NullKeyword(matching_text),
    "Object" => |matching_text| VB6Token::ObjectKeyword(matching_text),
    "On" => |matching_text| VB6Token::OnKeyword(matching_text),
    "Open" => |matching_text| VB6Token::OpenKeyword(matching_text),
    // Switched so that `Option` isn't selected for `Optional`.
    "Optional" => |matching_text| VB6Token::OptionalKeyword(matching_text),
    "Option" => |matching_text| VB6Token::OptionKeyword(matching_text),
    "Or" => |matching_text| VB6Token::OrKeyword(matching_text),
    "ParamArray" => |matching_text| VB6Token::ParamArrayKeyword(matching_text),
    "Preserve" => |matching_text| VB6Token::PreserveKeyword(matching_text),
    "Print" => |matching_text| VB6Token::PrintKeyword(matching_text),
    "Private" => |matching_text| VB6Token::PrivateKeyword(matching_text),
    "Property" => |matching_text| VB6Token::PropertyKeyword(matching_text),
    "Public" => |matching_text| VB6Token::PublicKeyword(matching_text),
    "Put" => |matching_text| VB6Token::PutKeyword(matching_text),
    "RaiseEvent" => |matching_text| VB6Token::RaiseEventKeyword(matching_text),
    "Randomize" => |matching_text| VB6Token::RandomizeKeyword(matching_text),
    "ReDim" => |matching_text| VB6Token::ReDimKeyword(matching_text),
    "Reset" => |matching_text| VB6Token::ResetKeyword(matching_text),
    "Resume" => |matching_text| VB6Token::ResumeKeyword(matching_text),
    "RmDir" => |matching_text| VB6Token::RmDirKeyword(matching_text),
    "RSet" => |matching_text| VB6Token::RSetKeyword(matching_text),
    "SavePicture" => |matching_text| VB6Token::SavePictureKeyword(matching_text),
    "SaveSetting" => |matching_text| VB6Token::SaveSettingKeyword(matching_text),
    "Seek" => |matching_text| VB6Token::SeekKeyword(matching_text),
    "Select" => |matching_text| VB6Token::SelectKeyword(matching_text),
    "SendKeys" => |matching_text| VB6Token::SendKeysKeyword(matching_text),
    // Switched so that `Set` isn't selected for `SetAttr`.
    "SetAttr" => |matching_text| VB6Token::SetAttrKeyword(matching_text),
    "Set" => |matching_text| VB6Token::SetKeyword(matching_text),
    "Single" => |matching_text| VB6Token::SingleKeyword(matching_text),
    "Static" => |matching_text| VB6Token::StaticKeyword(matching_text),
    "Step" => |matching_text| VB6Token::StepKeyword(matching_text),
    "Stop" => |matching_text| VB6Token::StopKeyword(matching_text),
    "String" => |matching_text| VB6Token::StringKeyword(matching_text),
    "Sub" => |matching_text| VB6Token::SubKeyword(matching_text),
    "Then" => |matching_text| VB6Token::ThenKeyword(matching_text),
    "Time" => |matching_text| VB6Token::TimeKeyword(matching_text),
    "To" => |matching_text| VB6Token::ToKeyword(matching_text),
    "True" => |matching_text| VB6Token::TrueKeyword(matching_text),
    "Type" => |matching_text| VB6Token::TypeKeyword(matching_text),
    "Unlock" => |matching_text| VB6Token::UnlockKeyword(matching_text),
    "Until" => |matching_text| VB6Token::UntilKeyword(matching_text),
    "Variant" => |matching_text| VB6Token::VariantKeyword(matching_text),
    "Wend" => |matching_text| VB6Token::WendKeyword(matching_text),
    "While" => |matching_text| VB6Token::WhileKeyword(matching_text),
    "Width" => |matching_text| VB6Token::WidthKeyword(matching_text),
    // Switched so that `With` isn't selected for `WithEvents`.
    "WithEvents" => |matching_text| VB6Token::WithEventsKeyword(matching_text),
    "With" => |matching_text| VB6Token::WithKeyword(matching_text),
    "Write" => |matching_text| VB6Token::WriteKeyword(matching_text),
    "Xor" => |matching_text| VB6Token::XorKeyword(matching_text),
};

static SYMBOL_TOKEN_LOOKUP_TABLE: OrderedMap<&'static str, for<'a> fn(&'a str) -> VB6Token<'a>> = phf_ordered_map! {
    "=" => |matching_text| VB6Token::EqualityOperator(matching_text),
    "$" => |matching_text| VB6Token::DollarSign(matching_text),
    "_" => |matching_text| VB6Token::Underscore(matching_text),
    "&" => |matching_text| VB6Token::Ampersand(matching_text),
    "%" => |matching_text| VB6Token::Percent(matching_text),
    "#" => |matching_text| VB6Token::Octothorpe(matching_text),
    "<" => |matching_text| VB6Token::LessThanOperator(matching_text),
    ">" => |matching_text| VB6Token::GreaterThanOperator(matching_text),
    "(" => |matching_text| VB6Token::LeftParentheses(matching_text),
    ")" => |matching_text| VB6Token::RightParentheses(matching_text),
    "," => |matching_text| VB6Token::Comma(matching_text),
    "+" => |matching_text| VB6Token::AdditionOperator(matching_text),
    "-" => |matching_text| VB6Token::SubtractionOperator(matching_text),
    "*" => |matching_text| VB6Token::MultiplicationOperator(matching_text),
    "\\" => |matching_text| VB6Token::BackwardSlashOperator(matching_text),
    "/" => |matching_text| VB6Token::DivisionOperator(matching_text),
    "." => |matching_text| VB6Token::PeriodOperator(matching_text),
    ":" => |matching_text| VB6Token::ColonOperator(matching_text),
    "^" => |matching_text| VB6Token::ExponentiationOperator(matching_text),
    "!" => |matching_text| VB6Token::ExclamationMark(matching_text),
    "[" => |matching_text| VB6Token::LeftSquareBracket(matching_text),
    "]" => |matching_text| VB6Token::RightSquareBracket(matching_text),
    ";" => |matching_text| VB6Token::Semicolon(matching_text),
    "@" => |matching_text| VB6Token::AtSign(matching_text),
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
/// use vb6parse::parsers::vb6_code_parse;
/// use vb6parse::SourceStream;
///
///
/// let mut input = SourceStream::new("test.bas", "Dim x As Integer");
/// let result = vb6_code_parse(&mut input);
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
/// assert_eq!(tokens[0], VB6Token::DimKeyword("Dim".into()));
/// assert_eq!(tokens[1], VB6Token::Whitespace(" ".into()));
/// assert_eq!(tokens[2], VB6Token::Identifier("x".into()));
/// assert_eq!(tokens[3], VB6Token::Whitespace(" ".into()));
/// assert_eq!(tokens[4], VB6Token::AsKeyword("As".into()));
/// assert_eq!(tokens[5], VB6Token::Whitespace(" ".into()));
/// assert_eq!(tokens[6], VB6Token::IntegerKeyword("Integer".into()));
/// ```
pub fn vb6_code_parse<'a>(
    input: &mut SourceStream<'a>,
) -> ParseResult<'a, Vec<VB6Token<'a>>, VB6CodeErrorKind> {
    let mut failures = vec![];
    let mut tokens = Vec::new();

    while !input.is_empty() {
        if let Some(token) = input.take_newline() {
            let token = VB6Token::Newline(token.into());
            tokens.push(token);
            continue;
        }

        if let Some((comment_token, newline_optional)) = input.take_line_comment() {
            tokens.push(comment_token);

            if let Some(newline_token) = newline_optional {
                tokens.push(newline_token);
            }
            continue;
        }

        if let Some((comment_token, newline_optional)) = input.take_rem_comment() {
            tokens.push(comment_token);

            if let Some(newline_token) = newline_optional {
                tokens.push(newline_token);
            }
            continue;
        }

        if let Some(string_literal_token) = input.take_string_literal() {
            tokens.push(string_literal_token);
            continue;
        }

        if let Some(keyword_token) = input.take_keyword() {
            tokens.push(keyword_token);
            continue;
        }

        if let Some(symbol_token) = input.take_symbol() {
            tokens.push(symbol_token);
            continue;
        }

        if let Some(digit_characters) = input.take_ascii_digits() {
            tokens.push(VB6Token::Number(digit_characters.into()));
            continue;
        }

        if let Some(identifier_token) = input.take_variable_name() {
            tokens.push(identifier_token);
            continue;
        }

        if let Some(whitespace_text) = input.take_ascii_whitespaces() {
            tokens.push(VB6Token::Whitespace(whitespace_text.into()));
            continue;
        }

        if let Some(token_text) = input.take_count(1) {
            let error = input.generate_error(VB6CodeErrorKind::UnknownToken {
                token: token_text.into(),
            });

            failures.push(error);
            continue;
        }
    }

    (tokens, failures).into()
}

impl<'a> VB6Tokenizer<'a> for &mut SourceStream<'a> {
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
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::*;
    ///
    /// let mut input = SourceStream::new("line_comment.bas".to_owned(), "' This is a comment\r\n");
    /// let Some((comment, Some(newline))) = input.take_line_comment() else {
    ///     panic!("rem comment failed to parse correctly.")
    /// };
    ///
    /// assert_eq!(comment, VB6Token::Comment("' This is a comment".into()));
    /// assert_eq!(newline, VB6Token::Newline("\r\n".into()));
    /// ```
    fn take_line_comment(self) -> Option<(VB6Token<'a>, Option<VB6Token<'a>>)> {
        self.peek_text("'", super::Comparator::CaseInsensitive)?;

        match self.take_until_newline() {
            None => None,
            Some((comment, newline_optional)) => {
                let comment_token = VB6Token::Comment(comment.into());

                match newline_optional {
                    None => Some((comment_token, None)),
                    Some(newline) => Some((comment_token, Some(VB6Token::Newline(newline.into())))),
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
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::*;
    ///
    /// let mut input = SourceStream::new("line_comment.bas".to_owned(), "REM This is a comment\r\n");
    /// let Some((comment, Some(newline))) = input.take_rem_comment() else {
    ///     panic!("rem comment failed to parse correctly.")
    /// };
    ///
    /// assert_eq!(comment, VB6Token::RemComment("REM This is a comment".into()));
    /// assert_eq!(newline, VB6Token::Newline("\r\n".into()));
    /// ```
    fn take_rem_comment(self) -> Option<(VB6Token<'a>, Option<VB6Token<'a>>)> {
        self.peek_text("REM", super::Comparator::CaseInsensitive)?;

        match self.take_until_newline() {
            None => None,
            Some((comment, newline_optional)) => {
                let comment_token = VB6Token::RemComment(comment.into());

                match newline_optional {
                    None => Some((comment_token, None)),
                    Some(newline) => Some((comment_token, Some(VB6Token::Newline(newline.into())))),
                }
            }
        }
    }

    fn take_string_literal(self) -> Option<VB6Token<'a>> {
        self.peek_text("\"", super::Comparator::CaseInsensitive)?;

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

        self.take_until_lambda(take_string, false)
            .map(|quoted_text| VB6Token::StringLiteral(quoted_text.into()))
    }

    fn take_symbol(self) -> Option<VB6Token<'a>> {
        for entry in SYMBOL_TOKEN_LOOKUP_TABLE.entries()
        {
            if let Some(matching_text) = self.take(*entry.0, Comparator::CaseSensitive) {
                return Some(entry.1(matching_text));
            }
        }

        None
    }

    fn take_matching_text(self, keyword: impl Into<&'a str>) -> Option<&'a str> {
        let keyword_match_text = keyword.into();
        let len = keyword_match_text.len();

        let content_left_len = self.contents.len() - self.offset();

        // If we are at the end of the stream and we just so happen to match the
        // length of the keyword, we need to check if we have an exact match.
        if content_left_len == len {
            return self.take(keyword_match_text, Comparator::CaseInsensitive);
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

        if let Some(peek_text) = self.peek(len + 1) {
            match peek_text.chars().last() {
                None => return None,
                Some(last) => {
                    if last.is_alphanumeric() || last == '_' && last != ' ' {
                        return None;
                    } else {
                        return self.take(keyword_match_text, Comparator::CaseInsensitive);
                    }
                }
            }
        }

        None
    }

    fn take_keyword(self) -> Option<VB6Token<'a>> {

        for entry in KEYWORD_TOKEN_LOOKUP_TABLE.entries()
        {
            if let Some(matching_text) = self.take_matching_text(*entry.0) {
                return Some(entry.1(matching_text));
            }
        }

        None
    }

    fn take_variable_name(self) -> Option<VB6Token<'a>> {
        if self.peek(1)?.chars().next()?.is_ascii_alphabetic() {
            let variable_text = self.take_ascii_underscore_alphanumerics()?;

            return Some(VB6Token::Identifier(variable_text.into()));
        }

        None
    }
}

#[cfg(test)]
mod test {
    use super::*;

    #[test]
    fn vb6_parse() {
        use crate::vb6code::vb6_code_parse;
        use crate::SourceStream;

        let mut input = SourceStream::new("", "Dim x As Integer");
        let result = vb6_code_parse(&mut input);

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let tokens = result.result.unwrap();

        assert_eq!(tokens.len(), 7);
        assert_eq!(tokens[0], VB6Token::DimKeyword("Dim".into()));
        assert_eq!(tokens[1], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[2], VB6Token::Identifier("x".into()));
        assert_eq!(tokens[3], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[4], VB6Token::AsKeyword("As".into()));
        assert_eq!(tokens[5], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[6], VB6Token::IntegerKeyword("Integer".into()));
    }

    #[test]
    fn vb6_string_as_end_of_stream_parse() {
        use crate::vb6code::vb6_code_parse;
        use crate::SourceStream;

        let mut input = SourceStream::new("", r#"x = "Test""#);
        let result = vb6_code_parse(&mut input);

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let tokens = result.result.unwrap();

        assert_eq!(tokens.len(), 5);
        assert_eq!(tokens[0], VB6Token::Identifier("x".into()));
        assert_eq!(tokens[1], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[2], VB6Token::EqualityOperator("=".into()));
        assert_eq!(tokens[3], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[4], VB6Token::StringLiteral("\"Test\"".into()));
    }

    #[test]
    fn vb6_string_at_start_of_stream_parse() {
        use crate::vb6code::vb6_code_parse;
        use crate::SourceStream;

        let mut input = SourceStream::new("", r#""Text""#);
        let result = vb6_code_parse(&mut input);

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let tokens = result.result.unwrap();

        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0], VB6Token::StringLiteral("\"Text\"".into()));
    }

    #[test]
    fn vb6_string_parse() {
        use crate::vb6code::vb6_code_parse;
        use crate::SourceStream;

        let mut input = SourceStream::new("", r#"x = "Test" 'This is a comment."#);
        let result = vb6_code_parse(&mut input);

        if result.has_failures() {
            result.failures[0].eprint();
        };

        let tokens = result.result.unwrap();

        assert_eq!(tokens.len(), 7);
        assert_eq!(tokens[0], VB6Token::Identifier("x".into()));
        assert_eq!(tokens[1], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[2], VB6Token::EqualityOperator("=".into()));
        assert_eq!(tokens[3], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[4], VB6Token::StringLiteral("\"Test\"".into()));
        assert_eq!(tokens[5], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[6], VB6Token::Comment("'This is a comment.".into()));
    }
}
