use crate::{
    language::VB6Token,
    parsers::{Comparator, SourceStream, Success},
    VB6CodeErrorKind,
};

pub trait VB6Tokenizer<'a> {
    fn take_line_comment(self) -> Option<(VB6Token<'a>, Option<VB6Token<'a>>)>;
    fn take_rem_comment(self) -> Option<(VB6Token<'a>, Option<VB6Token<'a>>)>;
    fn take_string_literal(self) -> Option<VB6Token<'a>>;
    fn take_keyword(self) -> Option<VB6Token<'a>>;
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
/// let Ok(success) = vb6_code_parse(&mut input) else {
///     panic!("Failed to parse vb6 code.");
/// };
///
/// let tokens = success.value();
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
) -> Result<Success<Vec<VB6Token<'a>>, VB6CodeErrorKind>, VB6CodeErrorKind> {
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

        if let Some(token_text) = input.peek(1) {
            let error_kind = VB6CodeErrorKind::UnknownToken {
                token: token_text.into(),
            };

            return Err(error_kind);
        } else {
            return Err(VB6CodeErrorKind::UnexpectedEndOfStream);
        }
    }

    return Ok(Success::Value(tokens));
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
    /// * Some() with a tuple where the first element is the comment token, including
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
        if self
            .peek_text("'", super::Comparator::CaseInsensitive)
            .is_none()
        {
            return None;
        }

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
    ///
    /// # Arguments
    ///
    /// * `input` - The input to parse.
    ///
    /// # Returns
    ///
    /// * Some() with a tuple, the the first element is the comment token
    /// including the 'REM ' characters at the start of the comment. The second
    /// is an optional token for the newline (it's only None if the comment is
    /// at the of the stream).
    ///
    /// * None if there is no REM comment at the current position in the stream.
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
        if self
            .peek_text("REM", super::Comparator::CaseInsensitive)
            .is_none()
        {
            return None;
        }

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
        if self.peek_text("\"", super::Comparator::CaseInsensitive) == None {
            return None;
        }

        // TODO: Need to handle error reporting of incorrect escape sequences as well
        // as string literals that hit a newline before the second qoute character.
        let mut first_qoute = false;
        let take_string = |next_character| match next_character {
            '\r' | '\n' => false,
            '\"' if !first_qoute => {
                first_qoute = true;

                true
            }
            '\"' if first_qoute => false,
            _ => true,
        };

        match self.take_until_lambda(take_string) {
            None => None,
            Some(qouted_text) => Some(VB6Token::StringLiteral(qouted_text.into())),
        }
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

    fn take_keyword(self) -> Option<VB6Token<'a>> {
        if let Some(keyword) = self.take("AddressOf", Comparator::CaseInsensitive) {
            return Some(VB6Token::AddressOfKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Alias", Comparator::CaseInsensitive) {
            return Some(VB6Token::AliasKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("And", Comparator::CaseInsensitive) {
            return Some(VB6Token::AndKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("AppActivate", Comparator::CaseInsensitive) {
            return Some(VB6Token::AppActivateKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("As", Comparator::CaseInsensitive) {
            return Some(VB6Token::AsKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Base", Comparator::CaseInsensitive) {
            return Some(VB6Token::BaseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Beep", Comparator::CaseInsensitive) {
            return Some(VB6Token::BeepKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Binary", Comparator::CaseInsensitive) {
            return Some(VB6Token::BinaryKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Boolean", Comparator::CaseInsensitive) {
            return Some(VB6Token::BooleanKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("ByRef", Comparator::CaseInsensitive) {
            return Some(VB6Token::ByRefKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Byte", Comparator::CaseInsensitive) {
            return Some(VB6Token::ByteKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("ByVal", Comparator::CaseInsensitive) {
            return Some(VB6Token::ByValKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Call", Comparator::CaseInsensitive) {
            return Some(VB6Token::CallKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Case", Comparator::CaseInsensitive) {
            return Some(VB6Token::CaseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("ChDir", Comparator::CaseInsensitive) {
            return Some(VB6Token::ChDirKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("ChDrive", Comparator::CaseInsensitive) {
            return Some(VB6Token::ChDriveKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Close", Comparator::CaseInsensitive) {
            return Some(VB6Token::CloseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Compare", Comparator::CaseInsensitive) {
            return Some(VB6Token::CompareKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Const", Comparator::CaseInsensitive) {
            return Some(VB6Token::ConstKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Currency", Comparator::CaseInsensitive) {
            return Some(VB6Token::CurrencyKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Date", Comparator::CaseInsensitive) {
            return Some(VB6Token::DateKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Decimal", Comparator::CaseInsensitive) {
            return Some(VB6Token::DecimalKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Declare", Comparator::CaseInsensitive) {
            return Some(VB6Token::DeclareKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefBool", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefBoolKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefByte", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefByteKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefCur", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefCurKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefDate", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefDateKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefDbl", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefDblKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefDec", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefDecKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefInt", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefIntKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefLng", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefLngKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefObj", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefObjKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefSng", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefSngKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefStr", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefStrKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DefVar", Comparator::CaseInsensitive) {
            return Some(VB6Token::DefVarKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("DeleteSetting", Comparator::CaseInsensitive) {
            return Some(VB6Token::DeleteSettingKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Dim", Comparator::CaseInsensitive) {
            return Some(VB6Token::DimKeyword(keyword.into()));
        }
        // switched so that `Do` isn't selected for `Double`.
        else if let Some(keyword) = self.take("Double", Comparator::CaseInsensitive) {
            return Some(VB6Token::DoubleKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Do", Comparator::CaseInsensitive) {
            return Some(VB6Token::DoKeyword(keyword.into()));
        }
        // switched so that `Else` isn't selected for `ElseIf`.
        else if let Some(keyword) = self.take("ElseIf", Comparator::CaseInsensitive) {
            return Some(VB6Token::ElseIfKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Else", Comparator::CaseInsensitive) {
            return Some(VB6Token::ElseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Empty", Comparator::CaseInsensitive) {
            return Some(VB6Token::EmptyKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("End", Comparator::CaseInsensitive) {
            return Some(VB6Token::EndKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Enum", Comparator::CaseInsensitive) {
            return Some(VB6Token::EnumKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Eqv", Comparator::CaseInsensitive) {
            return Some(VB6Token::EqvKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Erase", Comparator::CaseInsensitive) {
            return Some(VB6Token::EraseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Error", Comparator::CaseInsensitive) {
            return Some(VB6Token::ErrorKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Event", Comparator::CaseInsensitive) {
            return Some(VB6Token::EventKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Exit", Comparator::CaseInsensitive) {
            return Some(VB6Token::ExitKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Explicit", Comparator::CaseInsensitive) {
            return Some(VB6Token::ExplicitKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("False", Comparator::CaseInsensitive) {
            return Some(VB6Token::FalseKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("FileCopy", Comparator::CaseInsensitive) {
            return Some(VB6Token::FileCopyKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("For", Comparator::CaseInsensitive) {
            return Some(VB6Token::ForKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Friend", Comparator::CaseInsensitive) {
            return Some(VB6Token::FriendKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Function", Comparator::CaseInsensitive) {
            return Some(VB6Token::FunctionKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Get", Comparator::CaseInsensitive) {
            return Some(VB6Token::GetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Goto", Comparator::CaseInsensitive) {
            return Some(VB6Token::GotoKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("If", Comparator::CaseInsensitive) {
            return Some(VB6Token::IfKeyword(keyword.into()));
        }
        // switched so that `Imp` isn't selected for `Implements`.
        else if let Some(keyword) = self.take("Implements", Comparator::CaseInsensitive) {
            return Some(VB6Token::ImplementsKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Imp", Comparator::CaseInsensitive) {
            return Some(VB6Token::ImpKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Integer", Comparator::CaseInsensitive) {
            return Some(VB6Token::IntegerKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Is", Comparator::CaseInsensitive) {
            return Some(VB6Token::IsKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Kill", Comparator::CaseInsensitive) {
            return Some(VB6Token::KillKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Len", Comparator::CaseInsensitive) {
            return Some(VB6Token::LenKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Let", Comparator::CaseInsensitive) {
            return Some(VB6Token::LetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Lib", Comparator::CaseInsensitive) {
            return Some(VB6Token::LibKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Line", Comparator::CaseInsensitive) {
            return Some(VB6Token::LineKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Lock", Comparator::CaseInsensitive) {
            return Some(VB6Token::LockKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Load", Comparator::CaseInsensitive) {
            return Some(VB6Token::LoadKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Long", Comparator::CaseInsensitive) {
            return Some(VB6Token::LongKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("LSet", Comparator::CaseInsensitive) {
            return Some(VB6Token::LSetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Me", Comparator::CaseInsensitive) {
            return Some(VB6Token::MeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Mid", Comparator::CaseInsensitive) {
            return Some(VB6Token::MidKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("MkDir", Comparator::CaseInsensitive) {
            return Some(VB6Token::MkDirKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Mod", Comparator::CaseInsensitive) {
            return Some(VB6Token::ModKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Name", Comparator::CaseInsensitive) {
            return Some(VB6Token::NameKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("New", Comparator::CaseInsensitive) {
            return Some(VB6Token::NewKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Next", Comparator::CaseInsensitive) {
            return Some(VB6Token::NextKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Not", Comparator::CaseInsensitive) {
            return Some(VB6Token::NotKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Null", Comparator::CaseInsensitive) {
            return Some(VB6Token::NullKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Object", Comparator::CaseInsensitive) {
            return Some(VB6Token::ObjectKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("On", Comparator::CaseInsensitive) {
            return Some(VB6Token::OnKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Open", Comparator::CaseInsensitive) {
            return Some(VB6Token::OpenKeyword(keyword.into()));
        }
        // Switched so that `Option` isn't selected for `Optional`.
        else if let Some(keyword) = self.take("Optional", Comparator::CaseInsensitive) {
            return Some(VB6Token::OptionalKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Option", Comparator::CaseInsensitive) {
            return Some(VB6Token::OptionKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Or", Comparator::CaseInsensitive) {
            return Some(VB6Token::OrKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("ParamArray", Comparator::CaseInsensitive) {
            return Some(VB6Token::ParamArrayKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Preserve", Comparator::CaseInsensitive) {
            return Some(VB6Token::PreserveKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Print", Comparator::CaseInsensitive) {
            return Some(VB6Token::PrintKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Private", Comparator::CaseInsensitive) {
            return Some(VB6Token::PrivateKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Property", Comparator::CaseInsensitive) {
            return Some(VB6Token::PropertyKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Public", Comparator::CaseInsensitive) {
            return Some(VB6Token::PublicKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Put", Comparator::CaseInsensitive) {
            return Some(VB6Token::PutKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("RaiseEvent", Comparator::CaseInsensitive) {
            return Some(VB6Token::RaiseEventKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Randomize", Comparator::CaseInsensitive) {
            return Some(VB6Token::RandomizeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("ReDim", Comparator::CaseInsensitive) {
            return Some(VB6Token::ReDimKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Reset", Comparator::CaseInsensitive) {
            return Some(VB6Token::ResetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Resume", Comparator::CaseInsensitive) {
            return Some(VB6Token::ResumeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("RmDir", Comparator::CaseInsensitive) {
            return Some(VB6Token::RmDirKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("RSet", Comparator::CaseInsensitive) {
            return Some(VB6Token::RSetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("SavePicture", Comparator::CaseInsensitive) {
            return Some(VB6Token::SavePictureKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("SaveSetting", Comparator::CaseInsensitive) {
            return Some(VB6Token::SaveSettingKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Seek", Comparator::CaseInsensitive) {
            return Some(VB6Token::SeekKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Select", Comparator::CaseInsensitive) {
            return Some(VB6Token::SelectKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("SendKeys", Comparator::CaseInsensitive) {
            return Some(VB6Token::SendKeysKeyword(keyword.into()));
        }
        // Switched so that `Set` isn't selected for `SetAttr`.
        else if let Some(keyword) = self.take("SetAttr", Comparator::CaseInsensitive) {
            return Some(VB6Token::SetAttrKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Set", Comparator::CaseInsensitive) {
            return Some(VB6Token::SetKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Single", Comparator::CaseInsensitive) {
            return Some(VB6Token::SingleKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Static", Comparator::CaseInsensitive) {
            return Some(VB6Token::StaticKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Step", Comparator::CaseInsensitive) {
            return Some(VB6Token::StepKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Stop", Comparator::CaseInsensitive) {
            return Some(VB6Token::StopKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("String", Comparator::CaseInsensitive) {
            return Some(VB6Token::StringKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Sub", Comparator::CaseInsensitive) {
            return Some(VB6Token::SubKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Then", Comparator::CaseInsensitive) {
            return Some(VB6Token::ThenKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Time", Comparator::CaseInsensitive) {
            return Some(VB6Token::TimeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("To", Comparator::CaseInsensitive) {
            return Some(VB6Token::ToKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("True", Comparator::CaseInsensitive) {
            return Some(VB6Token::TrueKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Type", Comparator::CaseInsensitive) {
            return Some(VB6Token::TypeKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Unlock", Comparator::CaseInsensitive) {
            return Some(VB6Token::UnlockKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Until", Comparator::CaseInsensitive) {
            return Some(VB6Token::UntilKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Variant", Comparator::CaseInsensitive) {
            return Some(VB6Token::VariantKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Wend", Comparator::CaseInsensitive) {
            return Some(VB6Token::WendKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("While", Comparator::CaseInsensitive) {
            return Some(VB6Token::WhileKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Width", Comparator::CaseInsensitive) {
            return Some(VB6Token::WidthKeyword(keyword.into()));
        }
        // Switched so that `With` isn't selected for `WithEvents`.
        else if let Some(keyword) = self.take("WithEvents", Comparator::CaseInsensitive) {
            return Some(VB6Token::WithEventsKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("With", Comparator::CaseInsensitive) {
            return Some(VB6Token::WithKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Write", Comparator::CaseInsensitive) {
            return Some(VB6Token::WriteKeyword(keyword.into()));
        } else if let Some(keyword) = self.take("Xor", Comparator::CaseInsensitive) {
            return Some(VB6Token::XorKeyword(keyword.into()));
        }

        None
    }

    fn take_variable_name(self) -> Option<VB6Token<'a>> {
        if self.peek(1)?.chars().next()?.is_ascii_alphabetic() {
            let variable_text = self.take_ascii_alphanumerics()?;

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

        let Ok(success) = result else {
            let error = result.err().unwrap();
            panic!("vb6_code_parse errored: {error:?}");
        };

        if success.has_failures() {
            panic!("vb6_code_parse succeded but with warnings.");
        };

        let tokens = success.value();

        assert_eq!(tokens.len(), 7);
        assert_eq!(tokens[0], VB6Token::DimKeyword("Dim".into()));
        assert_eq!(tokens[1], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[2], VB6Token::Identifier("x".into()));
        assert_eq!(tokens[3], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[4], VB6Token::AsKeyword("As".into()));
        assert_eq!(tokens[5], VB6Token::Whitespace(" ".into()));
        assert_eq!(tokens[6], VB6Token::IntegerKeyword("Integer".into()));
    }
}
