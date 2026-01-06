//! Sub statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Sub statements with syntax:
//!
//! \[ Public | Private | Friend \] \[ Static \] Sub name \[ ( arglist ) \]
//! \[ statements \]
//! \[ Exit Sub \]
//! \[ statements \]
//! End Sub
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Visual Basic 6 subroutine with syntax:
    ///
    /// \[ Public | Private | Friend \] \[ Static \] Sub name \[ ( arglist ) \]
    /// \[ statements \]
    /// \[ Exit Sub \]
    /// \[ statements \]
    /// End Sub
    ///
    /// The Sub statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public      | Optional | Indicates that the Sub procedure is accessible to all other procedures in all modules. If used in a module that contains an Option Private statement, the procedure is not available outside the project. |
    /// | Private     | Optional | Indicates that the Sub procedure is accessible only to other procedures in the module where it is declared. |
    /// | Friend      | Optional | Used only in a class module. Indicates that the Sub procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
    /// | Static      | Optional | Indicates that the Sub procedure's local variables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the Sub, even if they are used in the procedure. |
    /// | name        | Required | Name of the Sub; follows standard variable naming conventions. |
    /// | arglist     | Optional | List of variables representing arguments that are passed to the Sub procedure when it is called. Multiple variables are separated by commas. |
    /// | statements  | Optional | Any group of statements to be executed within the Sub procedure.
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// \[ Optional \] \[ `ByVal` | `ByRef` \] \[ `ParamArray` \] varname \[ ( ) \] \[ As type \] \[ = defaultvalue \]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)
    pub(super) fn parse_sub_statement(&mut self) {
        // if we are now parsing a sub statement, we are no longer in the header.
        self.parsing_header = false;
        self.builder.start_node(SyntaxKind::SubStatement.to_raw());

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

        // Consume "Sub" keyword
        self.consume_token();

        // Consume any whitespace after "Sub"
        self.consume_whitespace();

        // Consume procedure name (keywords can be used as procedure names in VB6)
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

        // Parse body until "End Sub"
        self.parse_statement_list(|parser| {
            parser.at_token(Token::EndKeyword)
                && parser.peek_next_keyword() == Some(Token::SubKeyword)
        });

        // Consume "End Sub" and trailing tokens
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Sub"
            self.consume_whitespace();

            // Consume "Sub"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // SubStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn sub_public_static() {
        let source = "Public Static Sub Initialize()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                PublicKeyword,
                Whitespace,
                StaticKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_public() {
        let source = "Public Sub Initialize()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                PublicKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_private() {
        let source = "Private Sub Initialize()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_private_static() {
        let source = "Private Static Sub Initialize()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                PrivateKeyword,
                Whitespace,
                StaticKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_friend() {
        let source = "Friend Sub Initialize()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                FriendKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_friend_static() {
        let source = "Friend Static Sub Initialize()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                FriendKeyword,
                Whitespace,
                StaticKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_simple_no_params() {
        // Test simple sub with no parameters
        let source = "Sub DoSomething()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("DoSomething"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_with_params() {
        // Test sub with parameters
        let source = "Sub SetValue(ByVal x As Integer, ByVal y As Integer)\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("SetValue"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_with_exit_sub() {
        // Test sub with Exit Sub statement
        let source = "Sub Validate(x As Integer)\n    If x < 0 Then\n        Exit Sub\n    End If\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Validate"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
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
                                SubKeyword,
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
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_friend_modifier() {
        // Test Friend Sub
        let source = "Friend Sub ProcessData()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                FriendKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("ProcessData"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_static_modifier() {
        // Test Static Sub
        let source = "Static Sub Counter()\n    Dim count As Integer\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                StaticKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Counter"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("count"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_with_body() {
        // Test sub with body statements
        let source = "Sub Calculate()\n    Dim x As Integer\n    x = 10\n    MsgBox x\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Calculate"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("x"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        Identifier ("x"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_with_optional_params() {
        // Test sub with optional parameters
        let source = "Sub Process(ByVal x As Integer, Optional ByVal y As Integer = 0)\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Process"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    OptionalKeyword,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("0"),
                    },
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sub_with_keyword_as_name() {
        let source = r"Sub Text()
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Text"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
