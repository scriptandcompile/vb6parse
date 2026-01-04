//! Function statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Function statements with syntax:
//!
//! \[ Public | Private | Friend \] \[ Static \] Function name \[ ( arglist ) \] \[ As type \]
//! \[ statements \]
//! \[ name = expression \]
//! \[ Exit Function \]
//! \[ statements \]
//! \[ name = expression \]
//! End Function
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Visual Basic 6 `Function` with syntax:
    ///
    /// `\[ Public | Private | Friend \] \[ Static \] Function name \[ ( arglist ) \] \[ As type \]`
    /// `\[ statements \]`
    /// `\[ name = expression \]`
    /// `\[ Exit Function \]`
    /// `\[ statements \]`
    /// `\[ name = expression \]`
    /// `End Function`
    ///
    /// The `Function` statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public      | Optional | Indicates that the `Function` procedure is accessible to all other procedures in all modules. If used in a module that contains an Option Private, the procedure is not available outside the project. |
    /// | Private     | Optional | Indicates that the `Function` procedure is accessible only to other procedures in the module where it is declared. |
    /// | Friend      | Optional | Used only in a class module. Indicates that the `Function` procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
    /// | Static      | Optional | Indicates that the `Function` procedure's local variables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the `Function`, even if they are used in the procedure. |
    /// | name        | Required | Name of the `Function`; follows standard variable naming conventions. |
    /// | arglist     | Optional | List of variables representing arguments that are passed to the `Function` procedure when it is called. Multiple variables are separated by commas. |
    /// | type        | Optional | Data type of the value returned by the `Function` procedure; may be `Byte`, `Boolean`, `Integer`, `Long`, `Currency`, `Single`, `Double`, `Decimal` (not currently supported), `Date`, `String` (except fixed length), `Object`, `Variant`, or any user-defined type. |
    /// | statements  | Optional | Any group of statements to be executed within the `Function` procedure.
    /// | expression  | Optional | Return value of the `Function`. |
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// `\[ Optional \] \[ ByVal | ByRef \] \[ ParamArray \] varname \[ ( ) \] \[ As type \] \[ = defaultvalue \]`
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)
    pub(super) fn parse_function_statement(&mut self) {
        // if we are now parsing a function statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::FunctionStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume optional Public/Private/Friend keyword
        if self.at_token(Token::PublicKeyword)
            || self.at_token(Token::PrivateKeyword)
            || self.at_token(Token::FriendKeyword)
        {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume optional Static keyword
        if self.at_token(Token::StaticKeyword) {
            self.consume_token();

            // Consume any whitespace after Static
            self.consume_whitespace();
        }

        // Consume "Function" keyword
        self.consume_token();

        // Consume any whitespace after "Function"
        self.consume_whitespace();

        // Consume function name (keywords can be used as function names in VB6)
        if self.at_token(Token::Identifier) {
            self.consume_token();
        } else if self.at_keyword() {
            self.consume_token_as_identifier();
        }

        // Consume any whitespace before parameter list
        self.consume_whitespace();

        // Parse parameter list if present
        if self.at_token(Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(Token::Newline);

        // Parse body until "End Function"
        self.parse_statement_list(|parser| {
            parser.at_token(Token::EndKeyword)
                && parser.peek_next_keyword() == Some(Token::FunctionKeyword)
        });

        // Consume "End Function" and trailing tokens
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Function"
            self.consume_whitespace();

            // Consume "Function"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // FunctionStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn function_distinguishes_declarations_from_functions() {
        // Test that Private declaration and Private Function are correctly distinguished
        let source =
            "Private myVar As Integer\nPrivate Function GetVar() As Integer\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("myVar"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            FunctionStatement {
                PrivateKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("GetVar"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn public_function() {
        let source = "Public Function Test() As Integer\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_function() {
        let source = "Private Function Test() As Integer\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                PrivateKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn friend_function() {
        let source = "Friend Function Test() As Integer\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FriendKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn public_static_function() {
        let source = "Public Static Function Test() As Integer\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                StaticKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_static_function() {
        let source = "Private Static Function Test() As Integer\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                PrivateKeyword,
                Whitespace,
                StaticKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn friend_static_function() {
        let source = "Friend Static Function Test() As Integer\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FriendKeyword,
                Whitespace,
                StaticKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_private_static_with_args() {
        // Test Private Static Function
        let source = "Private Static Function Calculate(x As Long) As Long\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                PrivateKeyword,
                Whitespace,
                StaticKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Calculate"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_friend_as_string() {
        // Test Friend Function
        let source = "Friend Function ProcessData() As String\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FriendKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("ProcessData"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_with_line_continuation_in_params() {
        let source = r#"
Public Function Test( _
  ByVal x As Long _
) As String
    Test = "hello"
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Whitespace,
                    Underscore,
                    Newline,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Test"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"hello\""),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_with_line_continuation_after_open_paren() {
        // This is the exact pattern from audiostation modArgs.bas argGetSwitchArg
        let source = r#"
Public Function argGetSwitchArg( _
  ByRef Switch As String, _
  Optional ByRef Position As Long, _
  Optional ByVal UseWildcard As Boolean _
) As String
Dim I&
argGetSwitchArg = ""
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("argGetSwitchArg"),
                ParameterList {
                    LeftParenthesis,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Switch"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    OptionalKeyword,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Position"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    OptionalKeyword,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("UseWildcard"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    BooleanKeyword,
                    Whitespace,
                    Underscore,
                    Newline,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("I"),
                        Ampersand,
                        Newline,
                    },
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("argGetSwitchArg"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"\""),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_with_do_loop_before_end() {
        // Test that "End Function" after a DO loop is recognized correctly
        let source = r"
Public Function Test(ByVal x As Long) As String
Dim i As Long
Do
    i = i + 1
Loop
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    DoStatement {
                        DoKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_with_line_continuation_in_if_condition() {
        // Test from audiostation modArgs.bas - line continuation in IF condition
        let source = r#"
Public Function argGetArgs(ByRef argv() As String, ByRef argc As Long, _
 Optional ByVal Args As String)
Dim strArgTemp As String
Do Until strArgTemp = ""
  If InStr(1, strArgTemp, Chr$(34)) <> 0 And _
     InStr(1, strArgTemp, Chr$(34)) < InStr(1, strArgTemp, " ") Then
    strArgTemp = ""
  End If
Loop
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("argGetArgs"),
                ParameterList {
                    LeftParenthesis,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("argv"),
                    LeftParenthesis,
                    RightParenthesis,
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("argc"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    Underscore,
                    Newline,
                    Whitespace,
                    OptionalKeyword,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Args"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("strArgTemp"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    DoStatement {
                        DoKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("strArgTemp"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"\""),
                            },
                        },
                        Newline,
                        StatementList {
                            IfStatement {
                                Whitespace,
                                IfKeyword,
                                Whitespace,
                                BinaryExpression {
                                    BinaryExpression {
                                        CallExpression {
                                            Identifier ("InStr"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("1"),
                                                    },
                                                },
                                                Comma,
                                                Whitespace,
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("strArgTemp"),
                                                    },
                                                },
                                                Comma,
                                                Whitespace,
                                                Argument {
                                                    CallExpression {
                                                        Identifier ("Chr$"),
                                                        LeftParenthesis,
                                                        ArgumentList {
                                                            Argument {
                                                                NumericLiteralExpression {
                                                                    IntegerLiteral ("34"),
                                                                },
                                                            },
                                                        },
                                                        RightParenthesis,
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        InequalityOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                    Whitespace,
                                    AndKeyword,
                                    Whitespace,
                                    Underscore,
                                    Newline,
                                    Whitespace,
                                    BinaryExpression {
                                        CallExpression {
                                            Identifier ("InStr"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("1"),
                                                    },
                                                },
                                                Comma,
                                                Whitespace,
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("strArgTemp"),
                                                    },
                                                },
                                                Comma,
                                                Whitespace,
                                                Argument {
                                                    CallExpression {
                                                        Identifier ("Chr$"),
                                                        LeftParenthesis,
                                                        ArgumentList {
                                                            Argument {
                                                                NumericLiteralExpression {
                                                                    IntegerLiteral ("34"),
                                                                },
                                                            },
                                                        },
                                                        RightParenthesis,
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        LessThanOperator,
                                        Whitespace,
                                        CallExpression {
                                            Identifier ("InStr"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("1"),
                                                    },
                                                },
                                                Comma,
                                                Whitespace,
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("strArgTemp"),
                                                    },
                                                },
                                                Comma,
                                                Whitespace,
                                                Argument {
                                                    StringLiteralExpression {
                                                        StringLiteral ("\" \""),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                    },
                                },
                                Whitespace,
                                ThenKeyword,
                                Newline,
                                StatementList {
                                    Whitespace,
                                    AssignmentStatement {
                                        IdentifierExpression {
                                            Identifier ("strArgTemp"),
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        StringLiteralExpression {
                                            StringLiteral ("\"\""),
                                        },
                                        Newline,
                                    },
                                    Whitespace,
                                },
                                EndKeyword,
                                Whitespace,
                                IfKeyword,
                                Newline,
                            },
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_simple_no_params() {
        // Test simple function with no parameters
        let source = "Function GetValue() As Integer\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetValue"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_with_return_value() {
        // Test function with return value assignment
        let source = "Function GetValue() As Integer\n    GetValue = 42\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetValue"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("42"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_with_exit_function() {
        // Test function with Exit Function statement
        let source = "Function IsValid(x As Integer) As Boolean\n    If x < 0 Then\n        Exit Function\n    End If\n    IsValid = True\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsValid"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            ExitStatement {
                                Whitespace,
                                ExitKeyword,
                                Whitespace,
                                FunctionKeyword,
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("IsValid"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BooleanLiteralExpression {
                            TrueKeyword,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_no_return_type() {
        // Test function without explicit return type (defaults to Variant)
        let source = "Function GetData()\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetData"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn function_with_multiple_params() {
        // Test function with multiple parameters
        let source = "Function Add(ByVal x As Long, ByVal y As Long) As Long\nEnd Function\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("Add"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }
}
