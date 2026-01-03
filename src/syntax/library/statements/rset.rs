use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

/// # `RSet` Statement
///
/// Right-aligns a string within a string variable or copies one user-defined variable to another.
///
/// ## Syntax
///
/// ```vb
/// RSet stringvar = string
/// RSet varname1 = varname2  ' For user-defined types
/// ```
///
/// ## Parts
///
/// - **stringvar**: Required. String variable or property name to be right-aligned.
/// - **string**: Required. String expression to be right-aligned within stringvar.
/// - **varname1**: Required. Variable of a user-defined type.
/// - **varname2**: Required. Variable of a different user-defined type.
///
/// ## Remarks
///
/// - **String Alignment**: When used with string variables, `RSet` right-aligns the string within
///   the variable. If the string is shorter than the variable, spaces are added on the left to
///   achieve right alignment.
/// - **Fixed-Length Strings**: `RSet` is particularly useful with fixed-length strings where you
///   need to right-justify text within a specific width.
/// - **User-Defined Types**: When used with user-defined types (UDTs), `RSet` copies data from one
///   variable to another on a byte-by-byte basis. This can be useful for converting between
///   different UDT structures that have the same size.
/// - **Shorter Strings**: If the source string is shorter than the target variable, spaces are
///   added on the left side to right-align the text.
/// - **Longer Strings**: If the source string is longer than the target variable, the string is
///   truncated on the left side, keeping only the rightmost characters that fit.
/// - **Comparison to `LSet`**: `RSet` is the opposite of `LSet`. While `LSet` left-aligns strings,
///   `RSet` right-aligns them.
///
/// ## Example
///
/// ```vb
/// Dim MyString As String * 10
/// MyString = String(10, "X")  ' Fill with X's
/// RSet MyString = "VB6"       ' Result: "       VB6"
/// ```
///
/// ## Example with User-Defined Types
///
/// ```vb
/// Type TypeA
///     Name As String * 20
///     Age As Integer
/// End Type
///
/// Type TypeB
///     Data As String * 22
/// End Type
///
/// Dim VarA As TypeA
/// Dim VarB As TypeB
///
/// VarA.Name = "John"
/// VarA.Age = 30
/// RSet VarB = VarA  ' Copy VarA to VarB byte-by-byte
/// ```
///
/// ## See Also
///
/// - `LSet` statement (left-align strings)
/// - `Mid` statement (replace characters in a string)
/// - Fixed-length string variables
///
/// ## References
///
/// - [RSet Statement (Visual Basic 6.0)](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266258(v=vs.60))
impl Parser<'_> {
    pub(crate) fn parse_rset_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::RSetStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // RSet statement tests

    #[test]
    fn rset_simple() {
        let source = r#"
Sub Test()
    RSet myString = "test"
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("myString"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"test\""),
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
    fn rset_at_module_level() {
        let source = "RSet fixedStr = \"VB6\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            RSetStatement {
                RSetKeyword,
                Whitespace,
                Identifier ("fixedStr"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"VB6\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn rset_fixed_length_string() {
        let source = r"
Sub Test()
    RSet FixedString = userName
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
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
    fn rset_user_defined_type() {
        let source = r"
Sub Test()
    RSet myRecord = sourceRecord
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
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
    fn rset_with_expression() {
        let source = r"
Sub Test()
    RSet buffer = Left$(inputStr, 5)
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("buffer"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Left$"),
                        LeftParenthesis,
                        Identifier ("inputStr"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
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
    fn rset_with_member_access() {
        let source = r"
Sub Test()
    RSet obj.Property = value
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        PeriodOperator,
                        PropertyKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("value"),
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
    fn rset_with_concatenation() {
        let source = r"
Sub Test()
    RSet result = prefix & suffix
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("result"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("prefix"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("suffix"),
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
    fn rset_inside_if_statement() {
        let source = r#"
If condition Then
    RSet output = "aligned"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                IdentifierExpression {
                    Identifier ("condition"),
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        OutputKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"aligned\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn rset_inside_loop() {
        let source = r"
For i = 1 To 10
    RSet buffer = data(i)
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ForStatement {
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("buffer"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("data"),
                        LeftParenthesis,
                        Identifier ("i"),
                        RightParenthesis,
                        Newline,
                    },
                },
                NextKeyword,
                Whitespace,
                Identifier ("i"),
                Newline,
            },
        ]);
    }

    #[test]
    fn rset_with_comment() {
        let source = r"
Sub Test()
    RSet aligned = text ' Right-align text
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("aligned"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        TextKeyword,
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
    fn rset_preserves_whitespace() {
        let source = "RSet   target   =   source\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            RSetStatement {
                RSetKeyword,
                Whitespace,
                Identifier ("target"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                Identifier ("source"),
                Newline,
            },
        ]);
    }

    #[test]
    fn rset_with_array_element() {
        let source = r"
Sub Test()
    RSet arr(index) = value
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("arr"),
                        LeftParenthesis,
                        Identifier ("index"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("value"),
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
    fn rset_with_multidimensional_array() {
        let source = r"
Sub Test()
    RSet matrix(row, col) = data
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("matrix"),
                        LeftParenthesis,
                        Identifier ("row"),
                        Comma,
                        Whitespace,
                        Identifier ("col"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("data"),
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
    fn rset_with_nested_property() {
        let source = r"
Sub Test()
    RSet obj.Field.Value = newValue
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("obj"),
                        PeriodOperator,
                        Identifier ("Field"),
                        PeriodOperator,
                        Identifier ("Value"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("newValue"),
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
    fn rset_with_str_function() {
        let source = r"
Sub Test()
    RSet buffer = Str$(number)
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("buffer"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Str$"),
                        LeftParenthesis,
                        Identifier ("number"),
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
    fn rset_with_trim() {
        let source = r"
Sub Test()
    RSet output = RTrim$(input)
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        OutputKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("RTrim$"),
                        LeftParenthesis,
                        InputKeyword,
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
    fn rset_multiple_on_same_line() {
        let source = "RSet a = x: RSet b = y\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            RSetStatement {
                RSetKeyword,
                Whitespace,
                Identifier ("a"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                Identifier ("x"),
                ColonOperator,
                Whitespace,
                RSetKeyword,
                Whitespace,
                Identifier ("b"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                Identifier ("y"),
                Newline,
            },
        ]);
    }

    #[test]
    fn rset_with_empty_string() {
        let source = r#"
Sub Test()
    RSet buffer = ""
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("buffer"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"\""),
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
    fn rset_with_space_function() {
        let source = r"
Sub Test()
    RSet padded = Space$(10) & text
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("padded"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Space$"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        TextKeyword,
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
    fn rset_with_iif() {
        let source = r#"
Sub Test()
    RSet display = IIf(flag, "Yes", "No")
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("display"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("IIf"),
                        LeftParenthesis,
                        Identifier ("flag"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Yes\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"No\""),
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
    fn rset_in_select_case() {
        let source = r#"
Select Case mode
    Case 1
        RSet output = "Left"
    Case 2
        RSet output = "Right"
End Select
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SelectCaseStatement {
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
                        RSetStatement {
                            Whitespace,
                            RSetKeyword,
                            Whitespace,
                            OutputKeyword,
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteral ("\"Left\""),
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
                        RSetStatement {
                            Whitespace,
                            RSetKeyword,
                            Whitespace,
                            OutputKeyword,
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteral ("\"Right\""),
                            Newline,
                        },
                    },
                },
                EndKeyword,
                Whitespace,
                SelectKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn rset_in_with_block() {
        let source = r"
With recordset
    RSet .Name = newName
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WithStatement {
                WithKeyword,
                Whitespace,
                Identifier ("recordset"),
                Newline,
                StatementList {
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        PeriodOperator,
                        NameKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("newName"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                WithKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn rset_in_sub() {
        let source = r"
Sub FormatOutput()
    RSet buffer = data
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("FormatOutput"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("buffer"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("data"),
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
    fn rset_in_function() {
        let source = r"
Function RightJustify(text As String) As String
    RSet RightJustify = text
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("RightJustify"),
                ParameterList {
                    LeftParenthesis,
                },
                TextKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("RightJustify"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        TextKeyword,
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
    fn rset_with_string_functions() {
        let source = r"
Sub Test()
    RSet formatted = Left$(s, 5) & Mid$(s, 6, 3) & Right$(s, 2)
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("formatted"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Left$"),
                        LeftParenthesis,
                        Identifier ("s"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Mid$"),
                        LeftParenthesis,
                        Identifier ("s"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("6"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("3"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Right$"),
                        LeftParenthesis,
                        Identifier ("s"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("2"),
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
    fn rset_in_class_module() {
        let source = r"
Private buffer As String * 20

Public Sub Align(text As String)
    RSet buffer = text
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
    }

    #[test]
    fn rset_with_format() {
        let source = r#"
Sub Test()
    RSet display = Format$(value, "000.00")
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("display"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Format$"),
                        LeftParenthesis,
                        Identifier ("value"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"000.00\""),
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
    fn rset_with_ucase() {
        let source = r"
Sub Test()
    RSet result = UCase$(input)
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("result"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("UCase$"),
                        LeftParenthesis,
                        InputKeyword,
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
    fn rset_with_replace() {
        let source = r#"
Sub Test()
    RSet clean = Replace(dirty, " ", "_")
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("clean"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Replace"),
                        LeftParenthesis,
                        Identifier ("dirty"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\" \""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"_\""),
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
    fn rset_with_line_continuation() {
        let source = r"
Sub Test()
    RSet longVar _
        = expression
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
                    RSetStatement {
                        Whitespace,
                        RSetKeyword,
                        Whitespace,
                        Identifier ("longVar"),
                        Whitespace,
                        Underscore,
                        Newline,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("expression"),
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
