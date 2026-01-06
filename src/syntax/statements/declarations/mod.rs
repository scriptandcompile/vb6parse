//! Array statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 array statements and variable declarations:
//! - Variable declarations (Dim, Private, Public, Const, Static)
//! - Private and Public variables with `WithEvents` keyword for event-capable objects
//! - `ReDim` - Reallocate storage space for dynamic array variables
//! - Erase - Reinitialize the elements of fixed-size arrays and deallocate dynamic arrays
//!
//! # Variables with `WithEvents`
//!
//! The `WithEvents` keyword is used with `Private`, `Public`, or `Dim` to declare object variables
//! that can respond to events raised by the object. This is commonly used in class modules
//! and form modules.
//!
//! ## Syntax
//! ```vb
//! Private WithEvents variablename As objecttype
//! Public WithEvents variablename As objecttype
//! Dim WithEvents variablename As objecttype
//! ```
//!
//! ## Examples
//! ```vb
//! Private WithEvents m_button As CommandButton
//! Public WithEvents g_conn As ADODB.Connection
//! Private WithEvents txtInput As TextBox
//! Public WithEvents AppEvents As Application
//! ```
//!
//! ## Remarks
//! - `WithEvents` can only be used with object variables
//! - `WithEvents` variables must be declared as a specific class type, not As Object
//! - Events are accessible through the object's event procedures (`objectname_eventname`)
//! - Public `WithEvents` variables are accessible from other modules
//! - Commonly used with form controls, `ActiveX` objects, and custom classes that raise events
//!
//! [WithEvents Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a `ReDim` statement.
    ///
    /// VB6 `ReDim` statement syntax:
    /// - `ReDim` [Preserve] varname(subscripts) [As type] [, varname(subscripts) [As type]] ...
    ///
    /// Used at procedure level to reallocate storage space for dynamic array variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    pub(crate) fn parse_redim_statement(&mut self) {
        // if we are now parsing a ReDim statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ReDimStatement.to_raw());

        // Consume "ReDim" keyword
        self.consume_token();
        self.consume_whitespace();

        // Optional Preserve
        if self.at_token(Token::PreserveKeyword) {
            self.consume_token();
            self.consume_whitespace();
        }

        loop {
            self.consume_whitespace();

            if self.at_token(Token::Newline)
                || self.at_token(Token::ColonOperator)
                || self.is_at_end()
            {
                break;
            }

            // Variable name
            if self.at_token(Token::Identifier) {
                self.consume_token();
            } else {
                // Error recovery
                while !self.is_at_end()
                    && !self.at_token(Token::Comma)
                    && !self.at_token(Token::Newline)
                {
                    self.consume_token();
                }
            }

            self.consume_whitespace();

            // Array bounds: (1 To 10)
            if self.at_token(Token::LeftParenthesis) {
                self.consume_token();
                // Parse bounds list
                loop {
                    self.consume_whitespace();
                    if self.at_token(Token::RightParenthesis) {
                        break;
                    }
                    self.parse_expression(); // lower or upper
                    self.consume_whitespace();
                    if self.at_token(Token::ToKeyword) {
                        self.consume_token();
                        self.consume_whitespace();
                        self.parse_expression(); // upper
                    }

                    if self.at_token(Token::Comma) {
                        self.consume_token();
                    } else {
                        break;
                    }
                }
                if self.at_token(Token::RightParenthesis) {
                    self.consume_token();
                }
            }

            self.consume_whitespace();

            // As Type
            if self.at_token(Token::AsKeyword) {
                self.consume_token();
                self.consume_whitespace();
                // Type name
                self.consume_token();
                while self.at_token(Token::PeriodOperator) {
                    self.consume_token();
                    self.consume_token();
                }
            }

            self.consume_whitespace();

            if self.at_token(Token::Comma) {
                self.consume_token();
            } else {
                break;
            }
        }

        // Consume everything until newline (Preserve, variable declarations, etc.)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // ReDimStatement
    }

    /// Parse a Dim statement: Dim/Private/Public/Const/Static x As Type
    ///
    /// VB6 variable declaration statement syntax:
    /// - Dim varname [As type]
    /// - Private varname [As type]
    /// - Private `WithEvents` varname As objecttype
    /// - Public varname [As type]
    /// - Public `WithEvents` varname As objecttype
    /// - Const constname = expression
    /// - Static varname [As type]
    ///
    /// Used to declare variables and allocate storage space.
    ///
    /// The `WithEvents` keyword can be used with `Private`, `Public`, or `Dim` to declare
    /// object variables that can respond to events raised by the object.
    ///
    /// Examples:
    /// ```vb
    /// Dim x As Integer
    /// Private m_value As Long
    /// Private WithEvents m_button As CommandButton
    /// Public g_config As String
    /// Public WithEvents g_app As Application
    /// Const MAX_SIZE = 100
    /// Static counter As Long
    /// ```
    ///
    /// [Dim Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement)
    /// [WithEvents Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    pub(crate) fn parse_dim(&mut self) {
        // if we are now parsing a dim statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::DimStatement.to_raw());

        // Consume the keyword (Dim, Private, Public, Const, Static, etc.)
        self.consume_token();

        loop {
            self.consume_whitespace();

            if self.at_token(Token::Newline)
                || self.at_token(Token::ColonOperator)
                || self.is_at_end()
            {
                break;
            }

            // WithEvents
            if self.at_token(Token::WithEventsKeyword) {
                self.consume_token();
                self.consume_whitespace();
            }

            // Variable name
            if self.at_token(Token::Identifier) {
                self.consume_token();
            } else {
                // Error recovery: consume until comma or newline
                while !self.is_at_end()
                    && !self.at_token(Token::Comma)
                    && !self.at_token(Token::Newline)
                {
                    self.consume_token();
                }
            }

            self.consume_whitespace();

            // Array bounds: (1 To 10)
            if self.at_token(Token::LeftParenthesis) {
                self.consume_token();
                // Parse bounds list
                loop {
                    self.consume_whitespace();
                    if self.at_token(Token::RightParenthesis) {
                        break;
                    }
                    self.parse_expression(); // lower or upper
                    self.consume_whitespace();
                    if self.at_token(Token::ToKeyword) {
                        self.consume_token();
                        self.consume_whitespace();
                        self.parse_expression(); // upper
                    }

                    if self.at_token(Token::Comma) {
                        self.consume_token();
                    } else {
                        break;
                    }
                }
                if self.at_token(Token::RightParenthesis) {
                    self.consume_token();
                }
            }

            self.consume_whitespace();

            // As Type
            if self.at_token(Token::AsKeyword) {
                self.consume_token();
                self.consume_whitespace();
                if self.at_token(Token::NewKeyword) {
                    self.consume_token();
                    self.consume_whitespace();
                }
                // Type name (identifier or keyword)
                self.consume_token();
                // Handle complex types like ADODB.Connection
                while self.at_token(Token::PeriodOperator) {
                    self.consume_token();
                    self.consume_token();
                }
            }

            self.consume_whitespace();

            // Initializer (for Const or optional initialization)
            if self.at_token(Token::EqualityOperator) {
                self.consume_token();
                self.consume_whitespace();
                self.parse_expression();
            }

            self.consume_whitespace();

            if self.at_token(Token::Comma) {
                self.consume_token();
            } else {
                break;
            }
        }

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // DimStatement
    }

    /// Parse an Erase statement: Erase array1 [, array2] ...
    ///
    /// VB6 Erase statement syntax:
    /// - Erase arraylist
    ///
    /// The Erase statement is used to reinitialize the elements of fixed-size arrays
    /// and to release storage space used by dynamic arrays.
    ///
    /// The arraylist argument is a list of one or more comma-delimited array variable names.
    ///
    /// Behavior:
    /// - For fixed-size arrays: Reinitializes the elements to their default values
    ///   (0 for numeric types, "" for strings, Nothing for objects)
    /// - For dynamic arrays: Deallocates the memory used by the array
    ///
    /// Examples:
    /// ```vb
    /// Erase myArray
    /// Erase array1, array2, array3
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/erase-statement)
    pub(crate) fn parse_erase_statement(&mut self) {
        // if we are now parsing an erase statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::EraseStatement.to_raw());

        // Consume "Erase" keyword
        self.consume_token();

        // Consume everything until newline (array names, commas, etc.)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // EraseStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn redim_simple_array() {
        let source = r"
Sub Test()
    ReDim myArray(10)
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("myArray"),
                        LeftParenthesis,
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
    fn redim_with_preserve() {
        let source = r"
Sub Test()
    ReDim Preserve argv(argc - 1&)
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        PreserveKeyword,
                        Whitespace,
                        Identifier ("argv"),
                        LeftParenthesis,
                        Identifier ("argc"),
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        LongLiteral,
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
    fn redim_with_as_type() {
        let source = r"
Sub Test()
    ReDim ICI(1 To num) As ImageCodecInfo
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("ICI"),
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        Identifier ("num"),
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("ImageCodecInfo"),
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
    fn redim_preserve_with_as_type() {
        let source = r"
Sub Test()
    ReDim Preserve fileNameArray(rdIconMaximum) As String
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        PreserveKeyword,
                        Whitespace,
                        Identifier ("fileNameArray"),
                        LeftParenthesis,
                        Identifier ("rdIconMaximum"),
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
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
    fn redim_zero_based() {
        let source = r"
Sub Test()
    ReDim argv(0&)
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("argv"),
                        LeftParenthesis,
                        LongLiteral,
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
    fn redim_with_to_clause() {
        let source = r"
Sub Test()
    ReDim hIcon(lIconIndex To lIconIndex + nIcons * 2 - 1)
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("hIcon"),
                        LeftParenthesis,
                        Identifier ("lIconIndex"),
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        Identifier ("lIconIndex"),
                        Whitespace,
                        AdditionOperator,
                        Whitespace,
                        Identifier ("nIcons"),
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("2"),
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        IntegerLiteral ("1"),
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
    fn redim_multiple_arrays() {
        let source = r"
Sub Test()
    ReDim arr1(10), arr2(20), arr3(30)
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("arr1"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("arr2"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("20"),
                        },
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("arr3"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("30"),
                        },
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
    fn redim_in_if_statement() {
        let source = r"
Sub Test()
    If needResize Then ReDim myArray(newSize)
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
                            Identifier ("needResize"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        ReDimStatement {
                            ReDimKeyword,
                            Whitespace,
                            Identifier ("myArray"),
                            LeftParenthesis,
                            IdentifierExpression {
                                Identifier ("newSize"),
                            },
                            RightParenthesis,
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
    fn redim_with_comment() {
        let source = r"
Sub Test()
    ReDim Preserve fileNameArray(rdIconMaximum) As String ' the file location of the original icons
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        PreserveKeyword,
                        Whitespace,
                        Identifier ("fileNameArray"),
                        LeftParenthesis,
                        Identifier ("rdIconMaximum"),
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
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
    fn redim_multiple_in_sequence() {
        let source = r"
Sub Test()
    ReDim Preserve fileNameArray(rdIconMaximum) As String
    ReDim Preserve dictionaryLocationArray(rdIconMaximum) As String
    ReDim Preserve namesListArray(rdIconMaximum) As String
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        PreserveKeyword,
                        Whitespace,
                        Identifier ("fileNameArray"),
                        LeftParenthesis,
                        Identifier ("rdIconMaximum"),
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        PreserveKeyword,
                        Whitespace,
                        Identifier ("dictionaryLocationArray"),
                        LeftParenthesis,
                        Identifier ("rdIconMaximum"),
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        PreserveKeyword,
                        Whitespace,
                        Identifier ("namesListArray"),
                        LeftParenthesis,
                        Identifier ("rdIconMaximum"),
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
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
    fn redim_in_multiline_if() {
        let source = r"
Sub Test()
    If arraysNeedResize Then
        ReDim Preserve myArray(newSize)
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
                            Identifier ("arraysNeedResize"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            ReDimStatement {
                                Whitespace,
                                ReDimKeyword,
                                Whitespace,
                                PreserveKeyword,
                                Whitespace,
                                Identifier ("myArray"),
                                LeftParenthesis,
                                Identifier ("newSize"),
                                RightParenthesis,
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
    fn redim_with_expression_bounds() {
        let source = r"
Sub Test()
    ReDim Buffer(1 To Size) As Byte
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("Buffer"),
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        Identifier ("Size"),
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        ByteKeyword,
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
    fn redim_at_module_level() {
        let source = r"
ReDim globalArray(100)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ReDimStatement {
                ReDimKeyword,
                Whitespace,
                Identifier ("globalArray"),
                LeftParenthesis,
                NumericLiteralExpression {
                    IntegerLiteral ("100"),
                },
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn redim_multidimensional() {
        let source = r"
Sub Test()
    ReDim matrix(10, 20)
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("matrix"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("20"),
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

    // Dim statement tests

    #[test]
    fn dim_simple_declaration() {
        let source = "Dim x As Integer\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
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
        ]);
    }

    #[test]
    fn dim_private_declaration() {
        let source = "Private m_value As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_value"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn dim_public_declaration() {
        let source = "Public g_config As String\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                Identifier ("g_config"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn dim_multiple_variables() {
        let source = "Dim x, y, z As Integer\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("x"),
                Comma,
                Whitespace,
                Identifier ("y"),
                Comma,
                Whitespace,
                Identifier ("z"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn dim_const_declaration() {
        let source = "Const MAX_SIZE = 100\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                ConstKeyword,
                Whitespace,
                Identifier ("MAX_SIZE"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("100"),
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn dim_private_const() {
        let source = "Private Const MODULE_NAME = \"MyModule\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                ConstKeyword,
                Whitespace,
                Identifier ("MODULE_NAME"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"MyModule\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn dim_static_declaration() {
        let source = "Static counter As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                StaticKeyword,
                Whitespace,
                Identifier ("counter"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    // Erase statement tests

    #[test]
    fn erase_simple_array() {
        let source = r"
Sub Test()
    Erase myArray
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
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("myArray"),
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
    fn erase_multiple_arrays() {
        let source = r"
Sub Test()
    Erase array1, array2, array3
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
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("array1"),
                        Comma,
                        Whitespace,
                        Identifier ("array2"),
                        Comma,
                        Whitespace,
                        Identifier ("array3"),
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
    fn erase_at_module_level() {
        let source = "Erase globalArray\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EraseStatement {
                EraseKeyword,
                Whitespace,
                Identifier ("globalArray"),
                Newline,
            },
        ]);
    }

    #[test]
    fn erase_preserves_whitespace() {
        let source = "    Erase    myArray    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            EraseStatement {
                EraseKeyword,
                Whitespace,
                Identifier ("myArray"),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn erase_with_comment() {
        let source = r"
Sub Test()
    Erase tempArray ' Free up memory
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
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("tempArray"),
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
    fn erase_in_if_statement() {
        let source = r"
Sub Cleanup()
    If shouldClear Then
        Erase dataArray
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
                Identifier ("Cleanup"),
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
                            Identifier ("shouldClear"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            EraseStatement {
                                Whitespace,
                                EraseKeyword,
                                Whitespace,
                                Identifier ("dataArray"),
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
    fn erase_inline_if() {
        let source = r"
Sub Test()
    If resetFlag Then Erase buffer
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
                            Identifier ("resetFlag"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        EraseStatement {
                            EraseKeyword,
                            Whitespace,
                            Identifier ("buffer"),
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
    fn erase_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Erase tempArrays(i)
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
                            EraseStatement {
                                Whitespace,
                                EraseKeyword,
                                Whitespace,
                                Identifier ("tempArrays"),
                                LeftParenthesis,
                                Identifier ("i"),
                                RightParenthesis,
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
    fn erase_with_parentheses() {
        let source = r"
Sub Test()
    Erase myArray()
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
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("myArray"),
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
    fn multiple_erase_statements() {
        let source = r"
Sub Test()
    Erase array1
    DoSomething
    Erase array2
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
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("array1"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("DoSomething"),
                        Newline,
                    },
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("array2"),
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
    fn erase_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Erase dynamicArray
    If Err.Number <> 0 Then
        MsgBox "Error erasing array"
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
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("dynamicArray"),
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
                                StringLiteral ("\"Error erasing array\""),
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
    fn erase_after_redim() {
        let source = r"
Sub Test()
    ReDim myArray(100)
    ' Use the array
    Erase myArray
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
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("myArray"),
                        LeftParenthesis,
                        IntegerLiteral ("100"),
                        RightParenthesis,
                        Newline,
                    },
                    Whitespace,
                    EndOfLineComment,
                    Newline,
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("myArray"),
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
    fn erase_complex_array_list() {
        let source = r"
Sub Test()
    Erase buffer1, buffer2, cache(), tempData
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
                    EraseStatement {
                        Whitespace,
                        EraseKeyword,
                        Whitespace,
                        Identifier ("buffer1"),
                        Comma,
                        Whitespace,
                        Identifier ("buffer2"),
                        Comma,
                        Whitespace,
                        Identifier ("cache"),
                        LeftParenthesis,
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("tempData"),
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

    // Private variable declaration tests

    #[test]
    fn private_variable_simple() {
        let source = "Private m_name As String\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_name"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_integer() {
        let source = "Private m_count As Integer\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_count"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_long() {
        let source = "Private m_id As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_id"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_variant() {
        let source = "Private m_data As Variant\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_data"),
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_object() {
        let source = "Private m_connection As ADODB.Connection\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_connection"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("ADODB"),
                PeriodOperator,
                Identifier ("Connection"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_array() {
        let source = "Private m_items() As String\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_items"),
                LeftParenthesis,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_array_with_bounds() {
        let source = "Private m_matrix(1 To 10, 1 To 10) As Double\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_matrix"),
                LeftParenthesis,
                NumericLiteralExpression {
                    IntegerLiteral ("1"),
                },
                Whitespace,
                ToKeyword,
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("10"),
                },
                Comma,
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
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_multiple_declarations() {
        let source = "Private m_x, m_y, m_z As Integer\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_x"),
                Comma,
                Whitespace,
                Identifier ("m_y"),
                Comma,
                Whitespace,
                Identifier ("m_z"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_new_keyword() {
        let source = "Private m_collection As New Collection\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_collection"),
                Whitespace,
                AsKeyword,
                Whitespace,
                NewKeyword,
                Whitespace,
                Identifier ("Collection"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_custom_type() {
        let source = "Private m_person As PersonType\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_person"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("PersonType"),
                Newline,
            },
        ]);
    }

    // WithEvents tests

    #[test]
    fn private_withevents_simple() {
        let source = "Private WithEvents m_button As Button\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_button"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Button"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_form() {
        let source = "Private WithEvents m_form As Form\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_form"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Form"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_adodb_connection() {
        let source = "Private WithEvents m_conn As ADODB.Connection\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_conn"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("ADODB"),
                PeriodOperator,
                Identifier ("Connection"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_custom_class() {
        let source = "Private WithEvents m_worker As WorkerClass\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_worker"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("WorkerClass"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_preserves_whitespace() {
        let source = "    Private    WithEvents    m_obj    As    MyClass    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_obj"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("MyClass"),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_multiple_declarations() {
        let source = "Private WithEvents m_btn1 As Button\nPrivate WithEvents m_btn2 As Button\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_btn1"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Button"),
                Newline,
            },
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_btn2"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Button"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_in_class_module() {
        let source = r"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Private WithEvents m_timer As Timer

Private Sub m_timer_Tick()
    ' Handle timer event
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            VersionStatement {
                VersionKeyword,
                Whitespace,
                SingleLiteral,
                Whitespace,
                ClassKeyword,
                Newline,
            },
            PropertiesBlock {
                BeginKeyword,
                Newline,
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("MultiUse"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        SubtractionOperator,
                        IntegerLiteral ("1"),
                        Whitespace,
                        EndOfLineComment,
                    },
                    Newline,
                },
                EndKeyword,
                Newline,
            },
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_timer"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Timer"),
                Newline,
            },
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("m_timer_Tick"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
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
    fn private_withevents_mixed_with_regular() {
        let source = "Private m_value As Long\nPrivate WithEvents m_control As Control\nPrivate m_name As String\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_value"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_control"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Control"),
                Newline,
            },
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_name"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_excel_application() {
        let source = "Private WithEvents m_excelApp As Excel.Application\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_excelApp"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Excel"),
                PeriodOperator,
                Identifier ("Application"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_textbox() {
        let source = "Private WithEvents txtInput As TextBox\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("txtInput"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("TextBox"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_combobox() {
        let source = "Private WithEvents cboList As ComboBox\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("cboList"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("ComboBox"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_winsock() {
        let source = "Private WithEvents m_socket As Winsock\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_socket"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Winsock"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_mscomm() {
        let source = "Private WithEvents m_comm As MSComm\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_comm"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("MSComm"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_withevents_at_module_level() {
        let source = "Private WithEvents m_db As Database\n\nSub Test()\n    Set m_db = OpenDatabase(\"test.mdb\")\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_db"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DatabaseKeyword,
                Newline,
            },
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
                        Identifier ("m_db"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("OpenDatabase"),
                        LeftParenthesis,
                        StringLiteral ("\"test.mdb\""),
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
    fn private_variable_no_type() {
        let source = "Private m_temp\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_temp"),
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_boolean() {
        let source = "Private m_isValid As Boolean\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_isValid"),
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_byte() {
        let source = "Private m_flags As Byte\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_flags"),
                Whitespace,
                AsKeyword,
                Whitespace,
                ByteKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_currency() {
        let source = "Private m_price As Currency\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_price"),
                Whitespace,
                AsKeyword,
                Whitespace,
                CurrencyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_date() {
        let source = "Private m_startDate As Date\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_startDate"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DateKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_single() {
        let source = "Private m_ratio As Single\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_ratio"),
                Whitespace,
                AsKeyword,
                Whitespace,
                SingleKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn private_variable_double() {
        let source = "Private m_pi As Double\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_pi"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
            },
        ]);
    }

    // Public WithEvents tests

    #[test]
    fn public_withevents_simple() {
        let source = "Public WithEvents g_app As Application\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("g_app"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Application"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_form() {
        let source = "Public WithEvents MainForm As Form\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("MainForm"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Form"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_excel_application() {
        let source = "Public WithEvents xlApp As Excel.Application\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("xlApp"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Excel"),
                PeriodOperator,
                Identifier ("Application"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_adodb_connection() {
        let source = "Public WithEvents dbConn As ADODB.Connection\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("dbConn"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("ADODB"),
                PeriodOperator,
                Identifier ("Connection"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_custom_class() {
        let source = "Public WithEvents TaskManager As TaskProcessor\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("TaskManager"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("TaskProcessor"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_preserves_whitespace() {
        let source = "    Public    WithEvents    g_obj    As    CustomClass    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("g_obj"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("CustomClass"),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_multiple_declarations() {
        let source = "Public WithEvents g_ctrl1 As Control\nPublic WithEvents g_ctrl2 As Control\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("g_ctrl1"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Control"),
                Newline,
            },
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("g_ctrl2"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Control"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_in_class_module() {
        let source = r"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Public WithEvents g_worker As BackgroundWorker

Private Sub g_worker_Complete()
    ' Handle completion event
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.cls", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            VersionStatement {
                VersionKeyword,
                Whitespace,
                SingleLiteral,
                Whitespace,
                ClassKeyword,
                Newline,
            },
            PropertiesBlock {
                BeginKeyword,
                Newline,
                Whitespace,
                Property {
                    PropertyKey {
                        Identifier ("MultiUse"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    PropertyValue {
                        SubtractionOperator,
                        IntegerLiteral ("1"),
                        Whitespace,
                        EndOfLineComment,
                    },
                    Newline,
                },
                EndKeyword,
                Newline,
            },
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("g_worker"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("BackgroundWorker"),
                Newline,
            },
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("g_worker_Complete"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
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
    fn public_withevents_mixed_with_private() {
        let source = "Private WithEvents m_local As Control\nPublic WithEvents g_shared As Control\nPrivate m_data As String\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PrivateKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("m_local"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Control"),
                Newline,
            },
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("g_shared"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Control"),
                Newline,
            },
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_data"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_word_application() {
        let source = "Public WithEvents wdApp As Word.Application\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("wdApp"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Word"),
                PeriodOperator,
                Identifier ("Application"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_outlook_application() {
        let source = "Public WithEvents olApp As Outlook.Application\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("olApp"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Outlook"),
                PeriodOperator,
                Identifier ("Application"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_chart() {
        let source = "Public WithEvents ChartObject As Chart\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("ChartObject"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Chart"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_worksheet() {
        let source = "Public WithEvents ws As Worksheet\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("ws"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Worksheet"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_recordset() {
        let source = "Public WithEvents rs As ADODB.Recordset\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("rs"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("ADODB"),
                PeriodOperator,
                Identifier ("Recordset"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_at_module_level() {
        let source = "Public WithEvents ServerSocket As Winsock\n\nSub Initialize()\n    Set ServerSocket = New Winsock\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("ServerSocket"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Winsock"),
                Newline,
            },
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Initialize"),
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
                        Identifier ("ServerSocket"),
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("Winsock"),
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
    fn public_withevents_commandbutton() {
        let source = "Public WithEvents cmdSubmit As CommandButton\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("cmdSubmit"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("CommandButton"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_listbox() {
        let source = "Public WithEvents lstItems As ListBox\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("lstItems"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("ListBox"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_timer() {
        let source = "Public WithEvents tmrMain As Timer\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("tmrMain"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Timer"),
                Newline,
            },
        ]);
    }

    #[test]
    fn public_withevents_class_factory() {
        let source = "Public WithEvents Factory As ClassFactory\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                PublicKeyword,
                Whitespace,
                WithEventsKeyword,
                Whitespace,
                Identifier ("Factory"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("ClassFactory"),
                Newline,
            },
        ]);
    }
}
