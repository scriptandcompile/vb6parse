//! # Mid Statement
//!
//! Replaces a specified number of characters in a Variant (String) variable with characters from another string.
//!
//! ## Syntax
//!
//! ```vb
//! Mid(stringvar, start[, length]) = string
//! ```
//!
//! - `stringvar`: Required. Name of string variable to modify
//! - `start`: Required. Character position where replacement begins (1-based)
//! - `length`: Optional. Number of characters to replace. If omitted, uses entire length of `string`
//! - `string`: Required. String expression used as replacement
//!
//! ## Remarks
//!
//! - The number of characters replaced is always less than or equal to the number of characters in `stringvar`
//! - If `start` is greater than the length of `stringvar`, `stringvar` is unchanged
//! - If `length` is omitted, all characters from `start` to the end of the string are replaced
//! - `Mid` statement replaces characters in-place; it does not change the length of the original string
//! - If replacement string is longer than `length`, only `length` characters are used
//! - If replacement string is shorter than `length`, only available characters are replaced
//!
//! ## Examples
//!
//! ```vb
//! Dim s As String
//! s = "Hello World"
//! Mid(s, 7, 5) = "VB6!!"     ' s becomes "Hello VB6!!"
//!
//! s = "ABCDEFGH"
//! Mid(s, 3) = "123"          ' s becomes "AB123FGH"
//!
//! s = "Test"
//! Mid(s, 2, 2) = "XX"        ' s becomes "TXXt"
//! ```
//!
//! ## Reference
//!
//! [Mid Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mid-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a Mid statement.
    pub(crate) fn parse_mid_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::MidStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn mid_simple() {
        let source = r#"
Sub Test()
    Mid(text, 5, 3) = "abc"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_at_module_level() {
        let source = r#"Mid(globalStr, 1, 5) = "START""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_without_length() {
        let source = r"
Sub Test()
    Mid(s, 10) = replacement
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_with_expressions() {
        let source = r"
Sub Test()
    Mid(arr(i), startPos + 1, Len(newStr)) = newStr
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_preserves_whitespace() {
        let source = "    Mid  (  myString  ,  3  ,  2  )  =  \"XX\"    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_with_comment() {
        let source = r"
Sub Test()
    Mid(buffer, pos, 10) = data ' Replace 10 characters
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_in_if_statement() {
        let source = r#"
Sub Test()
    If needsUpdate Then
        Mid(statusText, 1, 7) = "UPDATED"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_inline_if() {
        let source = r#"
Sub Test()
    If valid Then Mid(s, 1, 1) = "A"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_mid_statements() {
        let source = r#"
Sub ReplaceChars()
    Mid(line1, 5) = "HELLO"
    Mid(line2, 1, 3) = "ABC"
    Mid(line3, 2, 4) = "TEST"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_replace_example() {
        let source = r#"
Sub Test()
    Dim s As String
    s = "Hello World"
    Mid(s, 7, 5) = "VB6!!"
    ' s now contains "Hello VB6!!"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_with_member_access() {
        let source = r"
Sub Test()
    Mid(obj.Name, 1, 10) = newName
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn mid_with_concatenation() {
        let source = r"
Sub Test()
    Mid(fullText, pos, 5) = prefix & suffix
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../../snapshots/syntax/library/statements/mid");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
