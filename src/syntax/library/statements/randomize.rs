use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    /// Parse a VB6 Randomize statement.
    ///
    /// # Syntax
    ///
    /// ```vb
    /// Randomize [number]
    /// ```
    ///
    /// # Arguments
    ///
    /// | Part | Optional / Required | Description |
    /// |------|---------------------|-------------|
    /// | number | Optional | A Variant or any valid numeric expression that is used as the new seed value to initialize the random number generator. |
    ///
    /// # Remarks
    ///
    /// - The Randomize statement initializes the random-number generator, giving it a new seed value.
    /// - If you omit number, the value returned by the system timer is used as the new seed value.
    /// - If Randomize is not used, the Rnd function (with no arguments) uses the same number as a seed the first time it is called, and thereafter uses the last generated number as a seed value.
    /// - To repeat sequences of random numbers, call Rnd with a negative argument immediately before using Randomize with a numeric argument.
    /// - Using Randomize with the same value for number does not repeat the previous sequence.
    ///
    /// # Examples
    ///
    /// ```vb
    /// ' Initialize random number generator
    /// Randomize
    /// x = Int((100 * Rnd) + 1)
    ///
    /// ' Initialize with specific seed
    /// Randomize 42
    /// x = Rnd
    ///
    /// ' Use timer as seed
    /// Randomize Timer
    /// ```
    ///
    /// # References
    ///
    /// [Microsoft VBA Language Reference - Randomize Statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/randomize-statement)
    pub(crate) fn parse_randomize_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::RandomizeStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn randomize_simple() {
        let source = r"
Sub Test()
    Randomize
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
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
    fn randomize_with_seed() {
        let source = r"
Sub Test()
    Randomize 42
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        IntegerLiteral ("42"),
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
    fn randomize_with_timer() {
        let source = r"
Sub Test()
    Randomize Timer
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        Identifier ("Timer"),
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
    fn randomize_with_expression() {
        let source = r"
Sub Test()
    Randomize x * 100 + 42
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        Identifier ("x"),
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("100"),
                        Whitespace,
                        AdditionOperator,
                        Whitespace,
                        IntegerLiteral ("42"),
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
    fn randomize_with_variable() {
        let source = r"
Sub Test()
    Dim seed As Long
    seed = 12345
    Randomize seed
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("seed"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("seed"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("12345"),
                        },
                        Newline,
                    },
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        Identifier ("seed"),
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
    fn randomize_in_if_statement() {
        let source = r"
Sub Test()
    If needsRandom Then
        Randomize
    End If
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("needsRandom"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            RandomizeStatement {
                                Whitespace,
                                RandomizeKeyword,
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
    fn randomize_inline_if() {
        let source = r"
Sub Test()
    If initialize Then Randomize
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("initialize"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        RandomizeStatement {
                            RandomizeKeyword,
                            Newline,
                        },
                        EndKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                },
            },
        ]);
    }

    #[test]
    fn randomize_with_comment() {
        let source = r"
Sub Test()
    Randomize ' Initialize RNG
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        EndOfLineComment,
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
    fn randomize_at_module_level() {
        let source = "Randomize\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(
            cst,
            [RandomizeStatement {
                RandomizeKeyword,
                Newline,
            },]
        );
    }

    #[test]
    fn randomize_preserves_whitespace() {
        let source = "    Randomize    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(
            cst,
            [
                Whitespace,
                RandomizeStatement {
                    RandomizeKeyword,
                    Whitespace,
                    Newline,
                },
            ]
        );
        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Randomize i
        x = Rnd
    Next i
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
                            RandomizeStatement {
                                Whitespace,
                                RandomizeKeyword,
                                Whitespace,
                                Identifier ("i"),
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
                                IdentifierExpression {
                                    Identifier ("Rnd"),
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
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
    fn multiple_randomize_statements() {
        let source = r"
Sub Test()
    Randomize
    x = Rnd
    Randomize 42
    y = Rnd
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
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
                        IdentifierExpression {
                            Identifier ("Rnd"),
                        },
                        Newline,
                    },
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        IntegerLiteral ("42"),
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
                        IdentifierExpression {
                            Identifier ("Rnd"),
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
    fn randomize_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Randomize
    If Err.Number <> 0 Then
        MsgBox "Error initializing RNG"
    End If
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
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        ResumeKeyword,
                        Whitespace,
                        NextKeyword,
                        Newline,
                    },
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Number"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"Error initializing RNG\""),
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
    fn randomize_with_negative_seed() {
        let source = r"
Sub Test()
    Randomize -1
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        SubtractionOperator,
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
    fn randomize_with_function_call() {
        let source = r"
Sub Test()
    Randomize GetSeed()
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        Identifier ("GetSeed"),
                        LeftParenthesis,
                        RightParenthesis,
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
    fn randomize_before_rnd() {
        let source = r"
Function GetRandomNumber() As Integer
    Randomize
    GetRandomNumber = Int((100 * Rnd) + 1)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetRandomNumber"),
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetRandomNumber"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Int"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        ParenthesizedExpression {
                                            LeftParenthesis,
                                            BinaryExpression {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("100"),
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("Rnd"),
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        AdditionOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
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
    fn randomize_with_select_case() {
        let source = r"
Sub Test()
    Select Case mode
        Case 1
            Randomize
        Case 2
            Randomize Timer
        Case Else
            Randomize 0
    End Select
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
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("mode"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                RandomizeStatement {
                                    Whitespace,
                                    RandomizeKeyword,
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
                                RandomizeStatement {
                                    Whitespace,
                                    RandomizeKeyword,
                                    Whitespace,
                                    Identifier ("Timer"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                RandomizeStatement {
                                    Whitespace,
                                    RandomizeKeyword,
                                    Whitespace,
                                    IntegerLiteral ("0"),
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
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn randomize_with_do_loop() {
        let source = r"
Sub Test()
    Do While True
        Randomize
        x = Rnd
        If x > 0.9 Then Exit Do
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
                        BooleanLiteralExpression {
                            TrueKeyword,
                        },
                        Newline,
                        StatementList {
                            RandomizeStatement {
                                Whitespace,
                                RandomizeKeyword,
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
                                IdentifierExpression {
                                    Identifier ("Rnd"),
                                },
                                Newline,
                            },
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
                                        SingleLiteral,
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
    fn randomize_multiline_with_continuation() {
        let source = r"
Sub Test()
    Randomize _
        Timer
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        Underscore,
                        Newline,
                        Whitespace,
                        Identifier ("Timer"),
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
    fn randomize_with_parentheses() {
        let source = r"
Sub Test()
    Randomize (seed)
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        LeftParenthesis,
                        Identifier ("seed"),
                        RightParenthesis,
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
    fn randomize_in_class_module() {
        let source = r"
Private Sub Class_Initialize()
    Randomize
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_with_decimal_seed() {
        let source = r"
Sub Test()
    Randomize 123.456
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
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        SingleLiteral,
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
    fn randomize_with_multiple_operations() {
        let source = r"
Sub GenerateRandomNumbers()
    Randomize Timer
    Dim nums(10) As Integer
    Dim i As Integer
    For i = 1 To 10
        nums(i) = Int((100 * Rnd) + 1)
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("GenerateRandomNumbers"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RandomizeStatement {
                        Whitespace,
                        RandomizeKeyword,
                        Whitespace,
                        Identifier ("Timer"),
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("nums"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
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
                            AssignmentStatement {
                                CallExpression {
                                    Identifier ("nums"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("i"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Int"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            BinaryExpression {
                                                ParenthesizedExpression {
                                                    LeftParenthesis,
                                                    BinaryExpression {
                                                        NumericLiteralExpression {
                                                            IntegerLiteral ("100"),
                                                        },
                                                        Whitespace,
                                                        MultiplicationOperator,
                                                        Whitespace,
                                                        IdentifierExpression {
                                                            Identifier ("Rnd"),
                                                        },
                                                    },
                                                    RightParenthesis,
                                                },
                                                Whitespace,
                                                AdditionOperator,
                                                Whitespace,
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("1"),
                                                },
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
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
}
