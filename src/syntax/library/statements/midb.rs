//! # `MidB` Statement
//!
//! Replaces a specified number of bytes in a Variant (String) variable with bytes from another string.
//!
//! ## Syntax
//!
//! ```vb
//! MidB(stringvar, start[, length]) = string
//! ```
//!
//! - `stringvar`: Required. Name of string variable to modify
//! - `start`: Required. Byte position where replacement begins (1-based)
//! - `length`: Optional. Number of bytes to replace. If omitted, uses entire length of `string`
//! - `string`: Required. String expression used as replacement
//!
//! ## Remarks
//!
//! - `MidB` is used with byte data contained in a string
//! - Works with byte positions rather than character positions (important for double-byte character sets)
//! - The number of bytes replaced is always less than or equal to the number of bytes in `stringvar`
//! - If `start` is greater than the number of bytes in `stringvar`, `stringvar` is unchanged
//! - If `length` is omitted, all bytes from `start` to the end of the string are replaced
//! - `MidB` statement replaces bytes in-place; it does not change the byte length of the original string
//! - If replacement string is longer than `length`, only `length` bytes are used
//! - If replacement string is shorter than `length`, only available bytes are replaced
//! - Primarily used when working with double-byte character sets (DBCS) like Japanese, Chinese, or Korean
//!
//! ## Examples
//!
//! ```vb
//! Dim s As String
//! s = "ABCDEFGH"
//! MidB(s, 3, 2) = "12"       ' Replaces 2 bytes starting at byte 3
//!
//! ' For DBCS strings:
//! Dim dbcsStr As String
//! dbcsStr = "日本語"          ' Japanese characters
//! MidB(dbcsStr, 1, 2) = "XX" ' Replaces first 2 bytes
//! ```
//!
//! ## Reference
//!
//! [MidB Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/midb-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a `MidB` statement.
    pub(crate) fn parse_midb_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::MidBStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // MidB statement tests

    #[test]
    fn midb_simple() {
        let source = r#"
Sub Test()
    MidB(text, 5, 3) = "abc"
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
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        TextKeyword,
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("3"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"abc\""),
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
    fn midb_at_module_level() {
        let source = r#"MidB(globalStr, 1, 5) = "START""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            MidBStatement {
                MidBKeyword,
                LeftParenthesis,
                Identifier ("globalStr"),
                Comma,
                Whitespace,
                IntegerLiteral ("1"),
                Comma,
                Whitespace,
                IntegerLiteral ("5"),
                RightParenthesis,
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"START\""),
            },
        ]);
    }

    #[test]
    fn midb_without_length() {
        let source = r"
Sub Test()
    MidB(s, 10) = replacement
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
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("s"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("replacement"),
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
    fn midb_with_expressions() {
        let source = r"
Sub Test()
    MidB(arr(i), startPos + 1, LenB(newStr)) = newStr
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
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("arr"),
                        LeftParenthesis,
                        Identifier ("i"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("startPos"),
                        Whitespace,
                        AdditionOperator,
                        Whitespace,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("LenB"),
                        LeftParenthesis,
                        Identifier ("newStr"),
                        RightParenthesis,
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("newStr"),
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
    fn midb_preserves_whitespace() {
        let source = "    MidB  (  myString  ,  3  ,  2  )  =  \"XX\"    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            MidBStatement {
                MidBKeyword,
                Whitespace,
                LeftParenthesis,
                Whitespace,
                Identifier ("myString"),
                Whitespace,
                Comma,
                Whitespace,
                IntegerLiteral ("3"),
                Whitespace,
                Comma,
                Whitespace,
                IntegerLiteral ("2"),
                Whitespace,
                RightParenthesis,
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"XX\""),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn midb_with_comment() {
        let source = r"
Sub Test()
    MidB(buffer, pos, 10) = data ' Replace 10 bytes
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
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("buffer"),
                        Comma,
                        Whitespace,
                        Identifier ("pos"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("data"),
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
    fn midb_in_if_statement() {
        let source = r#"
Sub Test()
    If needsUpdate Then
        MidB(statusText, 1, 7) = "UPDATED"
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("needsUpdate"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            MidBStatement {
                                Whitespace,
                                MidBKeyword,
                                LeftParenthesis,
                                Identifier ("statusText"),
                                Comma,
                                Whitespace,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                IntegerLiteral ("7"),
                                RightParenthesis,
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                StringLiteral ("\"UPDATED\""),
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
    fn midb_inline_if() {
        let source = r#"
Sub Test()
    If valid Then MidB(s, 1, 1) = "A"
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("valid"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        MidBStatement {
                            MidBKeyword,
                            LeftParenthesis,
                            Identifier ("s"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("1"),
                            RightParenthesis,
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteral ("\"A\""),
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
    fn multiple_midb_statements() {
        let source = r#"
Sub ReplaceBytes()
    MidB(line1, 5) = "HELLO"
    MidB(line2, 1, 3) = "ABC"
    MidB(line3, 2, 4) = "TEST"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ReplaceBytes"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("line1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"HELLO\""),
                        Newline,
                    },
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("line2"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("3"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"ABC\""),
                        Newline,
                    },
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("line3"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("2"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("4"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"TEST\""),
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
    fn midb_dbcs_example() {
        let source = r#"
Sub Test()
    Dim dbcsStr As String
    MidB(dbcsStr, 1, 2) = "XX"
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
                        Identifier ("dbcsStr"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("dbcsStr"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("2"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"XX\""),
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
    fn midb_with_member_access() {
        let source = r"
Sub Test()
    MidB(obj.Data, 1, 10) = newData
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
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("obj"),
                        PeriodOperator,
                        Identifier ("Data"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("newData"),
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
    fn midb_with_concatenation() {
        let source = r"
Sub Test()
    MidB(fullText, pos, 5) = prefix & suffix
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
                    MidBStatement {
                        Whitespace,
                        MidBKeyword,
                        LeftParenthesis,
                        Identifier ("fullText"),
                        Comma,
                        Whitespace,
                        Identifier ("pos"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
                        RightParenthesis,
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
}
