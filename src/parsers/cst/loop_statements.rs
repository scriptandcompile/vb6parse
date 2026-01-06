//! Do/Loop statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 loop statements:
//! - Do While...Loop
//! - Do Until...Loop
//! - Do...Loop While
//! - Do...Loop Until
//! - Do...Loop (infinite loop)
//! - While...Wend

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Do...Loop statement.
    ///
    /// VB6 supports several forms of Do loops:
    /// - Do While condition...Loop
    /// - Do Until condition...Loop
    /// - Do...Loop While condition
    /// - Do...Loop Until condition
    /// - Do...Loop (infinite loop, requires Exit Do)
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doloop-statement)
    pub(super) fn parse_do_statement(&mut self) {
        // if we are now parsing a do statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::DoStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "Do" keyword
        self.consume_token();

        // Consume whitespace after Do
        self.consume_whitespace();

        // Check if we have While or Until after Do
        let has_top_condition =
            self.at_token(Token::WhileKeyword) || self.at_token(Token::UntilKeyword);

        if has_top_condition {
            // Consume While or Until
            self.consume_token();

            // Consume any whitespace after While or Until
            self.consume_whitespace();

            // Parse condition - consume everything until newline
            self.parse_expression();
        }

        // Consume newline after Do line
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Loop"
        self.parse_statement_list(|parser| parser.at_token(Token::LoopKeyword));

        // Consume "Loop" keyword
        if self.at_token(Token::LoopKeyword) {
            self.consume_token();

            // Consume whitespace after Loop
            self.consume_whitespace();

            // Check if we have While or Until after Loop
            if self.at_token(Token::WhileKeyword) || self.at_token(Token::UntilKeyword) {
                // Consume While or Until
                self.consume_token();

                // Consume any whitespace after While or Until
                self.consume_whitespace();

                // Parse condition - consume everything until newline
                self.parse_expression();
            }

            // Consume newline after Loop
            if self.at_token(Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // DoStatement
    }

    /// Parse a While...Wend statement.
    ///
    /// VB6 While...Wend loop syntax:
    /// - While condition
    ///   ...statements...
    ///   Wend
    ///
    /// While...Wend statement syntax:
    ///
    /// | Part      | Description |
    /// |-----------|-------------|
    /// | condition | Required. Numeric or String expression that evaluates to True or False. If condition is Null, condition is treated as False. |
    /// | statements| Optional. One or more statements executed while condition is True. |
    ///
    /// Remarks:
    /// - If condition is True, all statements are executed until the Wend statement is encountered.
    /// - Control then returns to the While statement and condition is again checked.
    /// - If condition is still True, the process is repeated. If it's not True, execution resumes with the statement following the Wend statement.
    /// - While...Wend loops can be nested to any level. Each Wend matches the most recent While.
    /// - Note: The Do...Loop statement provides a more structured and flexible way to perform looping.
    /// - Tip: While...Wend is provided for compatibility with earlier versions of Visual Basic. Consider using Do...Loop instead for new code.
    ///
    /// Examples:
    /// ```vb
    /// Dim counter As Integer
    /// counter = 0
    /// While counter < 20
    ///     counter = counter + 1
    ///     Debug.Print counter
    /// Wend
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/whilewend-statement)
    /// Parse a While...Wend statement.
    ///
    /// While...Wend is a legacy VB6 loop construct that executes a block of
    /// statements while a condition is true. It has been superseded by Do While...Loop
    /// but is still supported for backward compatibility.
    ///
    /// Syntax:
    /// ```vb6
    /// While condition
    ///     statements
    /// Wend
    /// ```
    ///
    /// Example:
    /// ```vb6
    /// While x < 10
    ///     x = x + 1
    /// Wend
    /// ```
    ///
    /// The condition is evaluated before each iteration. If the condition is
    /// initially false, the loop body will not execute at all.
    pub(super) fn parse_while_statement(&mut self) {
        // if we are now parsing a while statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::WhileStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "While" keyword
        self.consume_token();

        // Consume whitespace after While
        self.consume_whitespace();

        // Parse condition - consume everything until newline
        self.parse_expression();

        // Consume newline after While line
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Wend"
        self.parse_statement_list(|parser| parser.at_token(Token::WendKeyword));

        // Consume "Wend" keyword
        if self.at_token(Token::WendKeyword) {
            self.consume_token();

            // Consume newline after Wend
            if self.at_token(Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // WhileStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // While...Wend statement tests

    #[test]
    fn while_simple() {
        let source = r"
Sub Test()
    While x < 10
        x = x + 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("x"),
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
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_at_module_level() {
        let source = r"
While x < 5
    x = x + 1
Wend
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WhileStatement {
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("x"),
                    },
                    Whitespace,
                    LessThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("5"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
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
                WendKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn while_with_string_condition() {
        let source = r#"
Sub Test()
    While inputText <> ""
        ProcessInput
    Wend
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("inputText"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"\""),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("ProcessInput"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_not_eof() {
        let source = r"
Sub Test()
    While Not EOF(1)
        Line Input #1, textLine
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        UnaryExpression {
                            NotKeyword,
                            Whitespace,
                            CallExpression {
                                Identifier ("EOF"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                        StatementList {
                            LineInputStatement {
                                Whitespace,
                                LineKeyword,
                                Whitespace,
                                InputKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("textLine"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_nested() {
        let source = r"
Sub Test()
    While i < 10
        While j < 5
            j = j + 1
        Wend
        i = i + 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            WhileStatement {
                                Whitespace,
                                WhileKeyword,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("j"),
                                    },
                                    Whitespace,
                                    LessThanOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
                                    },
                                },
                                Newline,
                                StatementList {
                                    Whitespace,
                                    AssignmentStatement {
                                        IdentifierExpression {
                                            Identifier ("j"),
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("j"),
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
                                    Whitespace,
                                },
                                WendKeyword,
                                Newline,
                            },
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
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_exit() {
        let source = r"
Sub Test()
    While True
        If x > 100 Then Exit Do
        x = x + 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BooleanLiteralExpression {
                            TrueKeyword,
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
                                    GreaterThanOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
                                    },
                                },
                                Whitespace,
                                ThenKeyword,
                                Whitespace,
                                ExitStatement {
                                    ExitKeyword,
                                    Whitespace,
                                    DoKeyword,
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
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("x"),
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
                                Whitespace,
                                WendKeyword,
                                Newline,
                            },
                            Unknown,
                            Whitespace,
                            Unknown,
                            Newline,
                        },
                    },
                },
            },
        ]);
    }

    #[test]
    fn while_with_comment() {
        let source = r"
Sub Test()
    While count < limit ' Loop until limit
        count = count + 1
    Wend ' End of loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("count"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("limit"),
                            },
                        },
                        StatementList {
                            Whitespace,
                            EndOfLineComment,
                            Newline,
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("count"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("count"),
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
                            Whitespace,
                        },
                        WendKeyword,
                    },
                    Whitespace,
                    EndOfLineComment,
                    Newline,
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn while_with_complex_condition() {
        let source = r"
Sub Test()
    While (x < 10 And y > 0) Or z = 5
        Process
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("x"),
                                        },
                                        Whitespace,
                                        LessThanOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("10"),
                                        },
                                    },
                                    Whitespace,
                                    AndKeyword,
                                    Whitespace,
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("y"),
                                        },
                                        Whitespace,
                                        GreaterThanOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            OrKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("z"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("5"),
                                },
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Process"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_function_call() {
        let source = r"
Sub Test()
    While IsValid(data)
        ProcessData data
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        CallExpression {
                            Identifier ("IsValid"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("data"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("ProcessData"),
                                Whitespace,
                                Identifier ("data"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_property_access() {
        let source = r"
Sub Test()
    While rs.EOF = False
        rs.MoveNext
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("rs"),
                                PeriodOperator,
                                Identifier ("EOF"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            BooleanLiteralExpression {
                                FalseKeyword,
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("rs"),
                                PeriodOperator,
                                Identifier ("MoveNext"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_empty_body() {
        let source = r"
Sub Test()
    While condition
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("condition"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_doevents() {
        let source = r"
Sub Test()
    While processing
        DoEvents
        CheckStatus
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("processing"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("DoEvents"),
                                Newline,
                            },
                            Whitespace,
                            CallStatement {
                                Identifier ("CheckStatus"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_string_length_check() {
        let source = r"
Sub Test()
    While Len(text) > 0
        text = Mid(text, 2)
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                LenKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            TextKeyword,
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    TextKeyword,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    MidKeyword,
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                TextKeyword,
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("2"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_array_access() {
        let source = r"
Sub Test()
    While arr(index) <> 0
        index = index + 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("arr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("index"),
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
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("index"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("index"),
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
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_if_statement() {
        let source = r"
Sub Test()
    While active
        If condition Then
            Process
        End If
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("active"),
                        },
                        Newline,
                        StatementList {
                            IfStatement {
                                Whitespace,
                                IfKeyword,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("condition"),
                                },
                                Whitespace,
                                ThenKeyword,
                                Newline,
                                StatementList {
                                    Whitespace,
                                    CallStatement {
                                        Identifier ("Process"),
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
                        },
                        WendKeyword,
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
    fn while_with_select_case() {
        let source = r"
Sub Test()
    While running
        Select Case action
            Case 1
                DoAction1
            Case 2
                DoAction2
        End Select
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("running"),
                        },
                        Newline,
                        StatementList {
                            SelectCaseStatement {
                                Whitespace,
                                SelectKeyword,
                                Whitespace,
                                CaseKeyword,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("action"),
                                },
                                Newline,
                                Whitespace,
                                CaseClause {
                                    CaseKeyword,
                                    Whitespace,
                                    IntegerLiteral ("1"),
                                    Newline,
                                    StatementList {
                                        Whitespace,
                                        CallStatement {
                                            Identifier ("DoAction1"),
                                            Newline,
                                        },
                                        Whitespace,
                                    },
                                },
                                CaseClause {
                                    CaseKeyword,
                                    Whitespace,
                                    IntegerLiteral ("2"),
                                    Newline,
                                    StatementList {
                                        Whitespace,
                                        CallStatement {
                                            Identifier ("DoAction2"),
                                            Newline,
                                        },
                                        Whitespace,
                                    },
                                },
                                EndKeyword,
                                Whitespace,
                                SelectKeyword,
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_for_loop() {
        let source = r"
Sub Test()
    While outerCondition
        For i = 1 To 10
            Process i
        Next i
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("outerCondition"),
                        },
                        Newline,
                        StatementList {
                            ForStatement {
                                Whitespace,
                                ForKeyword,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("1"),
                                },
                                Whitespace,
                                ToKeyword,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
                                },
                                Newline,
                                StatementList {
                                    Whitespace,
                                    CallStatement {
                                        Identifier ("Process"),
                                        Whitespace,
                                        Identifier ("i"),
                                        Newline,
                                    },
                                    Whitespace,
                                },
                                NextKeyword,
                                Whitespace,
                                Identifier ("i"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_triple_nested() {
        let source = r"
Sub Test()
    While a < 10
        While b < 5
            While c < 3
                c = c + 1
            Wend
            b = b + 1
        Wend
        a = a + 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("a"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            WhileStatement {
                                Whitespace,
                                WhileKeyword,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("b"),
                                    },
                                    Whitespace,
                                    LessThanOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
                                    },
                                },
                                Newline,
                                StatementList {
                                    WhileStatement {
                                        Whitespace,
                                        WhileKeyword,
                                        Whitespace,
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("c"),
                                            },
                                            Whitespace,
                                            LessThanOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("3"),
                                            },
                                        },
                                        Newline,
                                        StatementList {
                                            Whitespace,
                                            AssignmentStatement {
                                                IdentifierExpression {
                                                    Identifier ("c"),
                                                },
                                                Whitespace,
                                                EqualityOperator,
                                                Whitespace,
                                                BinaryExpression {
                                                    IdentifierExpression {
                                                        Identifier ("c"),
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
                                            Whitespace,
                                        },
                                        WendKeyword,
                                        Newline,
                                    },
                                    Whitespace,
                                    AssignmentStatement {
                                        IdentifierExpression {
                                            Identifier ("b"),
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("b"),
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
                                    Whitespace,
                                },
                                WendKeyword,
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("a"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("a"),
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
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_comparison_operators() {
        let source = r"
Sub Test()
    While value <= maxValue
        value = value * 2
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                            Whitespace,
                            LessThanOrEqualOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("maxValue"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("value"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("value"),
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("2"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_boolean_literal() {
        let source = r"
Sub Test()
    While True
        If userQuit Then Exit Sub
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BooleanLiteralExpression {
                            TrueKeyword,
                        },
                        Newline,
                        StatementList {
                            IfStatement {
                                Whitespace,
                                IfKeyword,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("userQuit"),
                                },
                                Whitespace,
                                ThenKeyword,
                                Whitespace,
                                ExitStatement {
                                    ExitKeyword,
                                    Whitespace,
                                    SubKeyword,
                                    Newline,
                                },
                                Whitespace,
                                WendKeyword,
                                Newline,
                            },
                            Unknown,
                            Whitespace,
                            Unknown,
                            Newline,
                        },
                    },
                },
            },
        ]);
    }

    #[test]
    fn while_with_multiple_statements() {
        let source = r"
Sub Test()
    While counter < 100
        counter = counter + 1
        sum = sum + counter
        average = sum / counter
        Display average
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("counter"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("100"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("counter"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("counter"),
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
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("sum"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("sum"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("counter"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("average"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("sum"),
                                    },
                                    Whitespace,
                                    DivisionOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("counter"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                            CallStatement {
                                Identifier ("Display"),
                                Whitespace,
                                Identifier ("average"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_reading_file() {
        let source = r#"
Sub ReadFile()
    Open "data.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, dataLine
        ProcessLine dataLine
    Wend
    Close #1
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ReadFile"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OpenStatement {
                        Whitespace,
                        OpenKeyword,
                        Whitespace,
                        StringLiteral ("\"data.txt\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Newline,
                    },
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        UnaryExpression {
                            NotKeyword,
                            Whitespace,
                            CallExpression {
                                Identifier ("EOF"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                        StatementList {
                            LineInputStatement {
                                Whitespace,
                                LineKeyword,
                                Whitespace,
                                InputKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("dataLine"),
                                Newline,
                            },
                            Whitespace,
                            CallStatement {
                                Identifier ("ProcessLine"),
                                Whitespace,
                                Identifier ("dataLine"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
                        Newline,
                    },
                    CloseStatement {
                        Whitespace,
                        CloseKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
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
    fn while_with_parenthesized_condition() {
        let source = r"
Sub Test()
    While (counter < limit)
        counter = counter + 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("counter"),
                                },
                                Whitespace,
                                LessThanOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("limit"),
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("counter"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("counter"),
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
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_recordset_iteration() {
        let source = r#"
Sub Test()
    Set rs = db.OpenRecordset("Table1")
    While Not rs.EOF
        Debug.Print rs!FieldName
        rs.MoveNext
    Wend
    rs.Close
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SetStatement {
                        Whitespace,
                        SetKeyword,
                        Whitespace,
                        Identifier ("rs"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("db"),
                        PeriodOperator,
                        Identifier ("OpenRecordset"),
                        LeftParenthesis,
                        StringLiteral ("\"Table1\""),
                        RightParenthesis,
                        Newline,
                    },
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        UnaryExpression {
                            NotKeyword,
                            Whitespace,
                            MemberAccessExpression {
                                Identifier ("rs"),
                                PeriodOperator,
                                Identifier ("EOF"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("rs"),
                                ExclamationMark,
                                Identifier ("FieldName"),
                                Newline,
                            },
                            Whitespace,
                            CallStatement {
                                Identifier ("rs"),
                                PeriodOperator,
                                Identifier ("MoveNext"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("rs"),
                        PeriodOperator,
                        CloseKeyword,
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
    fn while_timer_based() {
        let source = r"
Sub Test()
    startTime = Timer
    While Timer - startTime < 5
        DoEvents
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("startTime"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Timer"),
                        },
                        Newline,
                    },
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("Timer"),
                                },
                                Whitespace,
                                SubtractionOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("startTime"),
                                },
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("5"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("DoEvents"),
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_with_msgbox() {
        let source = r#"
Sub Test()
    While confirm = vbYes
        Process
        confirm = MsgBox("Continue?", vbYesNo)
    Wend
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("confirm"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("vbYes"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Process"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("confirm"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("MsgBox"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            StringLiteralExpression {
                                                StringLiteral ("\"Continue?\""),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("vbYesNo"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_collection_count() {
        let source = r"
Sub Test()
    While col.Count > 0
        col.Remove 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("col"),
                                PeriodOperator,
                                Identifier ("Count"),
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("col"),
                                PeriodOperator,
                                RemComment,
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_dir_function() {
        let source = r#"
Sub Test()
    fileName = Dir("*.txt")
    While fileName <> ""
        ProcessFile fileName
        fileName = Dir
    Wend
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fileName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Dir"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"*.txt\""),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("fileName"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"\""),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("ProcessFile"),
                                Whitespace,
                                Identifier ("fileName"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("fileName"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("Dir"),
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_instr_search() {
        let source = r"
Sub Test()
    position = InStr(text, searchTerm)
    While position > 0
        FoundAt position
        position = InStr(position + 1, text, searchTerm)
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("position"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("InStr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("searchTerm"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("position"),
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("FoundAt"),
                                Whitespace,
                                Identifier ("position"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("position"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("InStr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("position"),
                                                },
                                                Whitespace,
                                                AdditionOperator,
                                                Whitespace,
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("1"),
                                                },
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                TextKeyword,
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("searchTerm"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
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
    fn while_preserves_whitespace() {
        let source = "    While    x <    10    \n        x = x + 1\n    Wend    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            WhileStatement {
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("x"),
                    },
                    Whitespace,
                    LessThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("10"),
                    },
                },
                StatementList {
                    Whitespace,
                    Newline,
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
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
                    Whitespace,
                },
                WendKeyword,
            },
            Whitespace,
            Newline,
        ]);
    }

    // Do...Loop statement tests

    #[test]
    fn do_while_loop() {
        let source = r"
Sub Test()
    Do While x < 10
        x = x + 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("x"),
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
                            Whitespace,
                        },
                        LoopKeyword,
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
    fn do_until_loop() {
        let source = r"
Sub Test()
    Do Until x >= 10
        x = x + 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("x"),
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
                            Whitespace,
                        },
                        LoopKeyword,
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
    fn do_loop_while() {
        let source = r"
Sub Test()
    Do
        x = x + 1
    Loop While x < 10
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("x"),
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
                            Whitespace,
                        },
                        LoopKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
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
    fn do_loop_until() {
        let source = r"
Sub Test()
    Do
        x = x + 1
    Loop Until x >= 10
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("x"),
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
                            Whitespace,
                        },
                        LoopKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
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
    fn do_loop_infinite() {
        let source = r"
Sub Test()
    Do
        If x > 10 Then Exit Do
        x = x + 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
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
                                    GreaterThanOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
                                    },
                                },
                                Whitespace,
                                ThenKeyword,
                                Whitespace,
                                ExitStatement {
                                    ExitKeyword,
                                    Whitespace,
                                    DoKeyword,
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
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("x"),
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
                                Whitespace,
                                LoopKeyword,
                                Newline,
                            },
                            Unknown,
                            Whitespace,
                            Unknown,
                            Newline,
                        },
                    },
                },
            },
        ]);
    }

    #[test]
    fn nested_do_loops() {
        let source = r"
Sub Test()
    Do While i < 10
        Do While j < 5
            j = j + 1
        Loop
        i = i + 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            DoStatement {
                                Whitespace,
                                DoKeyword,
                                Whitespace,
                                WhileKeyword,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("j"),
                                    },
                                    Whitespace,
                                    LessThanOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
                                    },
                                },
                                Newline,
                                StatementList {
                                    Whitespace,
                                    AssignmentStatement {
                                        IdentifierExpression {
                                            Identifier ("j"),
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("j"),
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
                                    Whitespace,
                                },
                                LoopKeyword,
                                Newline,
                            },
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
                            Whitespace,
                        },
                        LoopKeyword,
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
    fn do_while_with_complex_condition() {
        let source = r"
Sub Test()
    Do While x < 10 And y > 0
        x = x + 1
        y = y - 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                LessThanOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
                                },
                            },
                            Whitespace,
                            AndKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("y"),
                                },
                                Whitespace,
                                GreaterThanOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("x"),
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
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("y"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("y"),
                                    },
                                    Whitespace,
                                    SubtractionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
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
    fn do_loop_preserves_whitespace() {
        let source = r"
Sub Test()
    Do  While  x < 10
        x = x + 1
    Loop  While  y > 0
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("x"),
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
                            Whitespace,
                        },
                        LoopKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("y"),
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
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
    fn function_with_do_loop_ending_at_end_function() {
        let source = r"Function Test()
Do
Loop
End Function
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        DoKeyword,
                        Newline,
                        StatementList,
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
    fn function_with_do_until_loop() {
        let source = r#"Function Test()
Do Until x = ""
  y = z
Loop
End Function
"#;

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        DoKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
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
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("y"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("z"),
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
}
