//! Enum statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Enum (enumeration) statements.
//!
//! Enum statement syntax:
//!
//! \[ Public | Private \] Enum name
//! membername \[= constantexpression\]
//! membername \[= constantexpression\]
//! ...
//! End Enum
//!
//! Enumerations provide a convenient way to work with sets of related constants
//! and to associate constant values with names.
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/enum-statement)

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Visual Basic 6 Enum statement with syntax:
    ///
    /// \[ Public | Private \] Enum name
    /// membername \[= constantexpression\]
    /// membername \[= constantexpression\]
    /// ...
    /// End Enum
    ///
    /// The Enum statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public      | Optional | Indicates that the Enum type is accessible to all other procedures in all modules. If used in a module that contains an Option Private statement, the Enum is not available outside the project. |
    /// | Private     | Optional | Indicates that the Enum type is accessible only to other procedures in the module where it is declared. |
    /// | name        | Required | Name of the Enum type; follows standard variable naming conventions. |
    /// | membername  | Required | Name of the enumeration member; follows standard variable naming conventions. |
    /// | constantexpression | Optional | Value to be assigned to the member (evaluates to a Long). If no constantexpression is specified, the value assigned is either zero (if it is the first membername), or 1 greater than the value of the immediately preceding membername. |
    ///
    /// Remarks:
    /// - Enumeration variables are variables declared with an Enum type.
    /// - Both variables and properties can be declared with an Enum type.
    /// - The values of Enum members are initialized to constant values within the Enum statement.
    /// - Values can't be modified at run time.
    /// - Enum values are Long integers.
    /// - By default, the first member is initialized to 0, and subsequent members are initialized to 1 more than the previous member.
    /// - You can assign specific values to members using the = operator.
    ///
    /// Examples:
    /// ```vb
    /// Public Enum SecurityLevel
    ///     IllegalEntry = -1
    ///     SecurityLevel1 = 0
    ///     SecurityLevel2 = 1
    /// End Enum
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/enum-statement)
    pub(super) fn parse_enum_statement(&mut self) {
        // if we are now parsing an enum statement, we are no longer in the header.
        self.parsing_header = false;
        self.builder.start_node(SyntaxKind::EnumStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume optional Public/Private keyword
        if self.at_token(Token::PublicKeyword) || self.at_token(Token::PrivateKeyword) {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume "Enum" keyword
        self.consume_token();

        // Consume any whitespace after "Enum"
        self.consume_whitespace();

        // Consume enum name (keywords can be used as enum names in VB6)
        if self.at_token(Token::Identifier) {
            self.consume_token();
        } else if self.at_keyword() {
            self.consume_token_as_identifier();
        }

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(Token::Newline);

        // Parse enum members until "End Enum"
        while !self.is_at_end() {
            // Check if we've reached "End Enum"
            if self.at_token(Token::EndKeyword)
                && self.peek_next_keyword() == Some(Token::EnumKeyword)
            {
                break;
            }

            // Consume enum member lines (identifier [= expression])
            // This includes whitespace, comments, identifiers, operators, and newlines
            match self.current_token() {
                Some(
                    Token::Whitespace
                    | Token::Newline
                    | Token::EndOfLineComment
                    | Token::RemComment
                    | Token::Identifier
                    | Token::EqualityOperator
                    | Token::IntegerLiteral
                    | Token::LongLiteral
                    | Token::SingleLiteral
                    | Token::DoubleLiteral
                    | Token::SubtractionOperator
                    | Token::AdditionOperator
                    | Token::MultiplicationOperator
                    | Token::DivisionOperator
                    | Token::LeftParenthesis
                    | Token::RightParenthesis
                    | Token::Ampersand
                    | Token::Comma,
                ) => {
                    self.consume_token();
                }
                _ => {
                    // Unknown token in enum body, consume it
                    self.consume_token_as_unknown();
                }
            }
        }

        // Consume "End Enum" and trailing tokens
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Enum"
            self.consume_whitespace();

            // Consume "Enum"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // EnumStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn enum_simple() {
        let source = r"
Enum Colors
    Red
    Green
    Blue
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_with_values() {
        let source = r"
Enum SecurityLevel
    IllegalEntry = -1
    SecurityLevel1 = 0
    SecurityLevel2 = 1
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_public() {
        let source = r"
Public Enum Status
    Active = 1
    Inactive = 0
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_private() {
        let source = r"
Private Enum InternalState
    Pending = 0
    Processing = 1
    Complete = 2
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_at_module_level() {
        let source = r"Enum Direction
    North = 0
    South = 1
    East = 2
    West = 3
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_with_comments() {
        let source = r"
Enum Priority
    Low = 0      ' Lowest priority
    Medium = 5   ' Medium priority
    High = 10    ' Highest priority
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_preserves_whitespace() {
        let source = "    Enum Test\n        Value1 = 1\n    End Enum\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_with_expressions() {
        let source = r"
Enum Flags
    None = 0
    Read = 1
    Write = 2
    ReadWrite = Read + Write
    All = &HFF
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_empty() {
        let source = r"
Enum EmptyEnum
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_multiple_in_module() {
        let source = r"
Public Enum Color
    Red = 1
    Green = 2
    Blue = 3
End Enum

Private Enum Size
    Small = 0
    Medium = 1
    Large = 2
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_with_hex_values() {
        let source = r"
Enum FileAttributes
    ReadOnly = &H1
    Hidden = &H2
    System = &H4
    Archive = &H20
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_long_member_list() {
        let source = r"
Enum DayOfWeek
    Sunday = 1
    Monday = 2
    Tuesday = 3
    Wednesday = 4
    Thursday = 5
    Friday = 6
    Saturday = 7
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn enum_negative_values() {
        let source = r"
Enum Temperature
    FreezingPoint = -273
    Zero = 0
    BoilingPoint = 100
End Enum
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/enum_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
