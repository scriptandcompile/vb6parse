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
mod test {
    use crate::*;

    // MidB statement tests
    #[test]
    fn midb_simple() {
        let source = r#"
Sub Test()
    MidB(text, 5, 3) = "abc"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
        assert!(debug.contains("MidBKeyword"));
        assert!(debug.contains("text"));
    }

    #[test]
    fn midb_at_module_level() {
        let source = r#"MidB(globalStr, 1, 5) = "START""#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
    }

    #[test]
    fn midb_without_length() {
        let source = r#"
Sub Test()
    MidB(s, 10) = replacement
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
        assert!(debug.contains("replacement"));
    }

    #[test]
    fn midb_with_expressions() {
        let source = r#"
Sub Test()
    MidB(arr(i), startPos + 1, LenB(newStr)) = newStr
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
        assert!(debug.contains("startPos"));
    }

    #[test]
    fn midb_preserves_whitespace() {
        let source = "    MidB  (  myString  ,  3  ,  2  )  =  \"XX\"    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(
            cst.text(),
            "    MidB  (  myString  ,  3  ,  2  )  =  \"XX\"    \n"
        );

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
    }

    #[test]
    fn midb_with_comment() {
        let source = r#"
Sub Test()
    MidB(buffer, pos, 10) = data ' Replace 10 bytes
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
        assert!(debug.contains("Comment"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
    }

    #[test]
    fn midb_inline_if() {
        let source = r#"
Sub Test()
    If valid Then MidB(s, 1, 1) = "A"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("MidBStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn midb_dbcs_example() {
        let source = r#"
Sub Test()
    Dim dbcsStr As String
    MidB(dbcsStr, 1, 2) = "XX"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
    }

    #[test]
    fn midb_with_member_access() {
        let source = r#"
Sub Test()
    MidB(obj.Data, 1, 10) = newData
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
        assert!(debug.contains("Data"));
    }

    #[test]
    fn midb_with_concatenation() {
        let source = r#"
Sub Test()
    MidB(fullText, pos, 5) = prefix & suffix
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("MidBStatement"));
        assert!(debug.contains("prefix"));
        assert!(debug.contains("suffix"));
    }
}
