use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 LSet statement syntax:
    // - LSet stringvar = string
    // - LSet varname1 = varname2 (for user-defined types)
    //
    // Left-aligns a string within a string variable, or copies a variable of one user-defined type
    // to another variable of a different user-defined type.
    //
    // The LSet statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | stringvar     | Required. Name of string variable. |
    // | string        | Required. String expression to be left-aligned within stringvar. |
    // | varname1      | Required. Variable name of the user-defined type being copied to. |
    // | varname2      | Required. Variable name of the user-defined type being copied from. |
    //
    // Remarks:
    // - LSet left-aligns strings within string variables.
    // - If string is shorter than stringvar, LSet left-aligns the string in stringvar and pads
    //   remaining characters with spaces.
    // - If string is longer than stringvar, LSet places only the leftmost characters that fit into
    //   stringvar.
    // - Warning: Using LSet to copy variables of different user-defined types is not recommended.
    //   Copying variables of one user-defined type into variables of a different user-defined type
    //   can produce unpredictable results.
    // - When copying between variables of user-defined types, the memory assigned to one variable is
    //   copied byte-for-byte to the memory assigned to the other variable.
    // - LSet is commonly used with fixed-length strings.
    // - LSet can be used with variant variables that contain strings.
    //
    // Examples:
    // ```vb
    // LSet MyString = "Left"
    // LSet FixedString = userName
    // LSet myRecord = sourceRecord
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/lset-statement)
    pub(crate) fn parse_lset_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::LSetStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // LSet statement tests

    #[test]
    fn lset_simple() {
        let source = r#"
Sub Test()
    LSet MyString = "Left"
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
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("MyString"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"Left\""),
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
    fn lset_at_module_level() {
        let source = "LSet myVar = \"Test\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            LSetStatement {
                LSetKeyword,
                Whitespace,
                Identifier ("myVar"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"Test\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn lset_fixed_length_string() {
        let source = r"
Sub Test()
    LSet FixedString = userName
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
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("FixedString"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("userName"),
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
    fn lset_user_defined_type() {
        let source = r"
Sub Test()
    LSet myRecord = sourceRecord
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
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("myRecord"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("sourceRecord"),
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
    fn lset_with_expression() {
        let source = r"
Sub Test()
    LSet buffer = Left(data, 10)
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
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("buffer"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Left"),
                        LeftParenthesis,
                        Identifier ("data"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("10"),
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
    fn lset_preserves_whitespace() {
        let source = "    LSet    myStr    =    \"Text\"    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            LSetStatement {
                LSetKeyword,
                Whitespace,
                Identifier ("myStr"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"Text\""),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn lset_with_comment() {
        let source = r#"
Sub Test()
    LSet MyString = "Left" ' Left-align string
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
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("MyString"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"Left\""),
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
    fn lset_in_if_statement() {
        let source = r"
Sub Test()
    If needsPadding Then
        LSet outputStr = inputStr
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
                            Identifier ("needsPadding"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            LSetStatement {
                                Whitespace,
                                LSetKeyword,
                                Whitespace,
                                Identifier ("outputStr"),
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                Identifier ("inputStr"),
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
    fn lset_inline_if() {
        let source = r"
Sub Test()
    If leftAlign Then LSet myStr = value
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
                            Identifier ("leftAlign"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        LSetStatement {
                            LSetKeyword,
                            Whitespace,
                            Identifier ("myStr"),
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            Identifier ("value"),
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
    fn multiple_lset_statements() {
        let source = r#"
Sub Test()
    LSet field1 = "A"
    LSet field2 = "B"
    LSet field3 = "C"
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
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("field1"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"A\""),
                        Newline,
                    },
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("field2"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"B\""),
                        Newline,
                    },
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("field3"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"C\""),
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
    fn lset_padding_example() {
        let source = r#"
Sub Test()
    Dim MyString As String * 10
    LSet MyString = "Left"
    ' MyString now contains "Left      " (padded with 6 spaces)
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("MyString"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("10"),
                        Newline,
                    },
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("MyString"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"Left\""),
                        Newline,
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
    fn lset_vs_rset() {
        let source = r#"
Sub Test()
    LSet leftAligned = "L"
    RSet rightAligned = "R"
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
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("leftAligned"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"L\""),
                        Newline,
                    },
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("rightAligned"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"R\""),
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
    fn lset_with_concatenation() {
        let source = r#"
Sub Test()
    LSet myBuffer = firstName & " " & lastName
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
                    LSetStatement {
                        Whitespace,
                        LSetKeyword,
                        Whitespace,
                        Identifier ("myBuffer"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("firstName"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\" \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("lastName"),
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
