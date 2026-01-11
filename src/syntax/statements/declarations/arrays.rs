//! ReDim statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 ReDim (reallocate dimension) statements:
//! - `ReDim` - Reallocate storage space for dynamic array variables
//! - `ReDim Preserve` - Reallocate while preserving existing data
//!
//! # ReDim Statement
//!
//! The `ReDim` statement is used at procedure level to reallocate storage space
//! for dynamic array variables. The optional `Preserve` keyword preserves the data
//! in the existing array when you change the size of the last dimension.
//!
//! ## Syntax
//! ```vb
//! ReDim [Preserve] varname(subscripts) [As type] [, varname(subscripts) [As type]] ...
//! ```
//!
//! ## Examples
//! ```vb
//! ReDim myArray(10)
//! ReDim Preserve argv(argc - 1)
//! ReDim ICI(1 To num) As ImageCodecInfo
//! ReDim Buffer(1 To Size) As Byte
//! ReDim arr1(10), arr2(20), arr3(30)
//! ```
//!
//! ## Remarks
//! - Can be used only at procedure level
//! - Can change the number of dimensions, size of each dimension, and data type
//! - Preserve keyword keeps existing data but only allows resizing the last dimension
//! - Can reallocate multiple arrays in a single statement
//!
//! [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))

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
}

#[cfg(test)]
mod tests {
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn redim_at_module_level() {
        let source = r"
ReDim globalArray(100)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
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

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/declarations/arrays");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
