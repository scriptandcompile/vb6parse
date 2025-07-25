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
    /// the single qoute while the second element is an optional newline token.
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
        // as string literals that hit a newline before the second qoute character.
        let mut escape_sequence_started = false;
        let mut qoute_character_count = 0;
        let take_string = |next_character| match next_character {
            // it doesn't matter what the character is if it is right after
            // the second qoute character.
            _ if qoute_character_count == 2 => true,
            '\r' | '\n' => true,
            '\\' => {
                escape_sequence_started = true;
                false
            }
            '\"' if !escape_sequence_started && qoute_character_count < 2 => {
                qoute_character_count += 1;
                false
            }
            _ if escape_sequence_started => {
                escape_sequence_started = false;
                false
            }
            _ => false,
        };

        self.take_until_lambda(take_string, false)
            .map(|qouted_text| VB6Token::StringLiteral(qouted_text.into()))
    }

    fn take_symbol(self) -> Option<VB6Token<'a>> {
        if let Some(token) = self.take("=", Comparator::CaseInsensitive) {
            return Some(VB6Token::EqualityOperator(token.into()));
        } else if let Some(token) = self.take("$", Comparator::CaseInsensitive) {
            return Some(VB6Token::DollarSign(token.into()));
        } else if let Some(token) = self.take("_", Comparator::CaseInsensitive) {
            return Some(VB6Token::Underscore(token.into()));
        } else if let Some(token) = self.take("&", Comparator::CaseInsensitive) {
            return Some(VB6Token::Ampersand(token.into()));
        } else if let Some(token) = self.take("%", Comparator::CaseInsensitive) {
            return Some(VB6Token::Percent(token.into()));
        } else if let Some(token) = self.take("#", Comparator::CaseInsensitive) {
            return Some(VB6Token::Octothorpe(token.into()));
        } else if let Some(token) = self.take("<", Comparator::CaseInsensitive) {
            return Some(VB6Token::LessThanOperator(token.into()));
        } else if let Some(token) = self.take(">", Comparator::CaseInsensitive) {
            return Some(VB6Token::GreaterThanOperator(token.into()));
        } else if let Some(token) = self.take("(", Comparator::CaseInsensitive) {
            return Some(VB6Token::LeftParanthesis(token.into()));
        } else if let Some(token) = self.take(")", Comparator::CaseInsensitive) {
            return Some(VB6Token::RightParanthesis(token.into()));
        } else if let Some(token) = self.take(",", Comparator::CaseInsensitive) {
            return Some(VB6Token::Comma(token.into()));
        } else if let Some(token) = self.take("+", Comparator::CaseInsensitive) {
            return Some(VB6Token::AdditionOperator(token.into()));
        } else if let Some(token) = self.take("-", Comparator::CaseInsensitive) {
            return Some(VB6Token::SubtractionOperator(token.into()));
        } else if let Some(token) = self.take("*", Comparator::CaseInsensitive) {
            return Some(VB6Token::MultiplicationOperator(token.into()));
        } else if let Some(token) = self.take("\\", Comparator::CaseInsensitive) {
            return Some(VB6Token::BackwardSlashOperator(token.into()));
        } else if let Some(token) = self.take("/", Comparator::CaseInsensitive) {
            return Some(VB6Token::DivisionOperator(token.into()));
        } else if let Some(token) = self.take(".", Comparator::CaseInsensitive) {
            return Some(VB6Token::PeriodOperator(token.into()));
        } else if let Some(token) = self.take(":", Comparator::CaseInsensitive) {
            return Some(VB6Token::ColonOperator(token.into()));
        } else if let Some(token) = self.take("^", Comparator::CaseInsensitive) {
            return Some(VB6Token::ExponentiationOperator(token.into()));
        } else if let Some(token) = self.take("!", Comparator::CaseInsensitive) {
            return Some(VB6Token::ExclamationMark(token.into()));
        } else if let Some(token) = self.take("[", Comparator::CaseInsensitive) {
            return Some(VB6Token::LeftSquareBracket(token.into()));
        } else if let Some(token) = self.take("]", Comparator::CaseInsensitive) {
            return Some(VB6Token::RightSquareBracket(token.into()));
        } else if let Some(token) = self.take(";", Comparator::CaseInsensitive) {
            return Some(VB6Token::Semicolon(token.into()));
        } else if let Some(token) = self.take("@", Comparator::CaseInsensitive) {
            return Some(VB6Token::AtSign(token.into()));
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
        if let Some(keyword) = self.take_matching_text("AddressOf") {
            return Some(VB6Token::AddressOfKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Alias") {
            return Some(VB6Token::AliasKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("And") {
            return Some(VB6Token::AndKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("AppActivate") {
            return Some(VB6Token::AppActivateKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("As") {
            return Some(VB6Token::AsKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Base") {
            return Some(VB6Token::BaseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Beep") {
            return Some(VB6Token::BeepKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Binary") {
            return Some(VB6Token::BinaryKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Boolean") {
            return Some(VB6Token::BooleanKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("ByRef") {
            return Some(VB6Token::ByRefKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Byte") {
            return Some(VB6Token::ByteKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("ByVal") {
            return Some(VB6Token::ByValKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Call") {
            return Some(VB6Token::CallKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Case") {
            return Some(VB6Token::CaseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("ChDir") {
            return Some(VB6Token::ChDirKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("ChDrive") {
            return Some(VB6Token::ChDriveKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Close") {
            return Some(VB6Token::CloseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Compare") {
            return Some(VB6Token::CompareKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Const") {
            return Some(VB6Token::ConstKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Currency") {
            return Some(VB6Token::CurrencyKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Date") {
            return Some(VB6Token::DateKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Decimal") {
            return Some(VB6Token::DecimalKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Declare") {
            return Some(VB6Token::DeclareKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefBool") {
            return Some(VB6Token::DefBoolKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefByte") {
            return Some(VB6Token::DefByteKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefCur") {
            return Some(VB6Token::DefCurKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefDate") {
            return Some(VB6Token::DefDateKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefDbl") {
            return Some(VB6Token::DefDblKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefDec") {
            return Some(VB6Token::DefDecKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefInt") {
            return Some(VB6Token::DefIntKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefLng") {
            return Some(VB6Token::DefLngKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefObj") {
            return Some(VB6Token::DefObjKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefSng") {
            return Some(VB6Token::DefSngKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefStr") {
            return Some(VB6Token::DefStrKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DefVar") {
            return Some(VB6Token::DefVarKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("DeleteSetting") {
            return Some(VB6Token::DeleteSettingKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Dim") {
            return Some(VB6Token::DimKeyword(keyword.into()));
        }
        // switched so that `Do` isn't selected for `Double`.
        else if let Some(keyword) = self.take_matching_text("Double") {
            return Some(VB6Token::DoubleKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Do") {
            return Some(VB6Token::DoKeyword(keyword.into()));
        }
        // switched so that `Else` isn't selected for `ElseIf`.
        else if let Some(keyword) = self.take_matching_text("ElseIf") {
            return Some(VB6Token::ElseIfKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Else") {
            return Some(VB6Token::ElseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Empty") {
            return Some(VB6Token::EmptyKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("End") {
            return Some(VB6Token::EndKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Enum") {
            return Some(VB6Token::EnumKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Eqv") {
            return Some(VB6Token::EqvKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Erase") {
            return Some(VB6Token::EraseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Error") {
            return Some(VB6Token::ErrorKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Event") {
            return Some(VB6Token::EventKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Exit") {
            return Some(VB6Token::ExitKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Explicit") {
            return Some(VB6Token::ExplicitKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("False") {
            return Some(VB6Token::FalseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("FileCopy") {
            return Some(VB6Token::FileCopyKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("For") {
            return Some(VB6Token::ForKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Friend") {
            return Some(VB6Token::FriendKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Function") {
            return Some(VB6Token::FunctionKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Get") {
            return Some(VB6Token::GetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Goto") {
            return Some(VB6Token::GotoKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("If") {
            return Some(VB6Token::IfKeyword(keyword.into()));
        }
        // switched so that `Imp` isn't selected for `Implements`.
        else if let Some(keyword) = self.take_matching_text("Implements") {
            return Some(VB6Token::ImplementsKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Imp") {
            return Some(VB6Token::ImpKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Integer") {
            return Some(VB6Token::IntegerKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Is") {
            return Some(VB6Token::IsKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Kill") {
            return Some(VB6Token::KillKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Len") {
            return Some(VB6Token::LenKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Let") {
            return Some(VB6Token::LetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Lib") {
            return Some(VB6Token::LibKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Line") {
            return Some(VB6Token::LineKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Lock") {
            return Some(VB6Token::LockKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Load") {
            return Some(VB6Token::LoadKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Long") {
            return Some(VB6Token::LongKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("LSet") {
            return Some(VB6Token::LSetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Me") {
            return Some(VB6Token::MeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Mid") {
            return Some(VB6Token::MidKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("MkDir") {
            return Some(VB6Token::MkDirKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Mod") {
            return Some(VB6Token::ModKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Name") {
            return Some(VB6Token::NameKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("New") {
            return Some(VB6Token::NewKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Next") {
            return Some(VB6Token::NextKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Not") {
            return Some(VB6Token::NotKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Null") {
            return Some(VB6Token::NullKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Object") {
            return Some(VB6Token::ObjectKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("On") {
            return Some(VB6Token::OnKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Open") {
            return Some(VB6Token::OpenKeyword(keyword.into()));
        }
        // Switched so that `Option` isn't selected for `Optional`.
        else if let Some(keyword) = self.take_matching_text("Optional") {
            return Some(VB6Token::OptionalKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Option") {
            return Some(VB6Token::OptionKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Or") {
            return Some(VB6Token::OrKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("ParamArray") {
            return Some(VB6Token::ParamArrayKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Preserve") {
            return Some(VB6Token::PreserveKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Print") {
            return Some(VB6Token::PrintKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Private") {
            return Some(VB6Token::PrivateKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Property") {
            return Some(VB6Token::PropertyKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Public") {
            return Some(VB6Token::PublicKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Put") {
            return Some(VB6Token::PutKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("RaiseEvent") {
            return Some(VB6Token::RaiseEventKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Randomize") {
            return Some(VB6Token::RandomizeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("ReDim") {
            return Some(VB6Token::ReDimKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Reset") {
            return Some(VB6Token::ResetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Resume") {
            return Some(VB6Token::ResumeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("RmDir") {
            return Some(VB6Token::RmDirKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("RSet") {
            return Some(VB6Token::RSetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("SavePicture") {
            return Some(VB6Token::SavePictureKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("SaveSetting") {
            return Some(VB6Token::SaveSettingKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Seek") {
            return Some(VB6Token::SeekKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Select") {
            return Some(VB6Token::SelectKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("SendKeys") {
            return Some(VB6Token::SendKeysKeyword(keyword.into()));
        }
        // Switched so that `Set` isn't selected for `SetAttr`.
        else if let Some(keyword) = self.take_matching_text("SetAttr") {
            return Some(VB6Token::SetAttrKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Set") {
            return Some(VB6Token::SetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Single") {
            return Some(VB6Token::SingleKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Static") {
            return Some(VB6Token::StaticKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Step") {
            return Some(VB6Token::StepKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Stop") {
            return Some(VB6Token::StopKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("String") {
            return Some(VB6Token::StringKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Sub") {
            return Some(VB6Token::SubKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Then") {
            return Some(VB6Token::ThenKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Time") {
            return Some(VB6Token::TimeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("To") {
            return Some(VB6Token::ToKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("True") {
            return Some(VB6Token::TrueKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Type") {
            return Some(VB6Token::TypeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Unlock") {
            return Some(VB6Token::UnlockKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Until") {
            return Some(VB6Token::UntilKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Variant") {
            return Some(VB6Token::VariantKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Wend") {
            return Some(VB6Token::WendKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("While") {
            return Some(VB6Token::WhileKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Width") {
            return Some(VB6Token::WidthKeyword(keyword.into()));
        }
        // Switched so that `With` isn't selected for `WithEvents`.
        else if let Some(keyword) = self.take_matching_text("WithEvents") {
            return Some(VB6Token::WithEventsKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("With") {
            return Some(VB6Token::WithKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Write") {
            return Some(VB6Token::WriteKeyword(keyword.into()));
        } else if let Some(keyword) = self.take_matching_text("Xor") {
            return Some(VB6Token::XorKeyword(keyword.into()));
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
