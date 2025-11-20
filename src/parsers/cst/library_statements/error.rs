use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    // VB6 Error statement syntax:
    // - Error errornumber
    //
    // Generates a run-time error; can be used instead of the Err.Raise method.
    //
    // The Error statement syntax has this part:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | errornumber   | Required. Any valid error number. |
    //
    // Remarks:
    // - The Error statement is supported for backward compatibility.
    // - In new code, use the Err object's Raise method to generate run-time errors.
    // - If errornumber is defined, the Error statement calls the error handler after the properties
    //   of the Err object are assigned the following default values:
    //   * Err.Number: The value specified as the argument to the Error statement
    //   * Err.Source: The name of the current Visual Basic project
    //   * Err.Description: String expression corresponding to the return value of the Error function
    //     for the specified Number, if this string exists
    //
    // Examples:
    // ```vb
    // Error 11  ' Generate "Division by zero" error
    // Error 53  ' Generate "File not found" error
    // Error vbObjectError + 1000  ' Generate custom error
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/error-statement)
    pub(crate) fn parse_error_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::ErrorStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // Error statement tests
    #[test]
    fn error_simple() {
        let source = r#"
Sub Test()
    Error 11
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
        assert!(debug.contains("ErrorKeyword"));
    }

    #[test]
    fn error_at_module_level() {
        let source = "Error 53\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_with_literal() {
        let source = r#"
Sub Test()
    Error 9
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_with_expression() {
        let source = r#"
Sub Test()
    Error vbObjectError + 1000
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
        assert!(debug.contains("vbObjectError"));
    }

    #[test]
    fn error_with_variable() {
        let source = r#"
Sub Test()
    Error errorNumber
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
        assert!(debug.contains("errorNumber"));
    }

    #[test]
    fn error_preserves_whitespace() {
        let source = "    Error    11    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Error    11    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_with_comment() {
        let source = r#"
Sub Test()
    Error 11 ' Division by zero
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn error_in_if_statement() {
        let source = r#"
Sub Test()
    If shouldFail Then
        Error 5
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_inline_if() {
        let source = r#"
Sub Test()
    If invalidData Then Error 13
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_in_select_case() {
        let source = r#"
Sub Test()
    Select Case errorType
        Case 1
            Error 11
        Case 2
            Error 13
    End Select
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let error_count = debug.matches("ErrorStatement").count();
        assert_eq!(error_count, 2);
    }

    #[test]
    fn error_with_error_handler() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    DoSomething
    Exit Sub
ErrorHandler:
    Error 1000
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_custom_number() {
        let source = r#"
Sub Test()
    Error 32000
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn multiple_error_statements() {
        let source = r#"
Sub Test()
    Error 1
    DoSomething
    Error 2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let error_count = debug.matches("ErrorStatement").count();
        assert_eq!(error_count, 2);
    }
}
