use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    // VB6 Date statement syntax:
    // - Date = dateexpression
    //
    // Sets the current system date.
    //
    // dateexpression: Required. Any expression that can represent a date.
    //
    // Note: The Date statement is used to set the date. To retrieve the current date,
    // use the Date function.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/date-statement)
    pub(crate) fn parse_date_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::DateStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn date_simple() {
        let source = r#"
Sub Test()
    Date = #1/1/2024#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_with_variable() {
        let source = r#"
Sub Test()
    Date = newDate
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_with_function_call() {
        let source = r#"
Sub Test()
    Date = DateSerial(2024, 1, 1)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_with_string_expression() {
        let source = r#"
Sub Test()
    Date = "January 1, 2024"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_with_expression() {
        let source = r#"
Sub Test()
    Date = Now() + 7
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_preserves_whitespace() {
        let source = r#"
Sub Test()
    Date   =   #1/1/2024#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn multiple_date_statements() {
        let source = r#"
Sub Test()
    Date = #1/1/2024#
    Date = #2/1/2024#
    Date = #3/1/2024#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("DateStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn date_in_if_statement() {
        let source = r#"
Sub Test()
    If resetDate Then
        Date = #1/1/2024#
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn date_inline_if() {
        let source = r#"
Sub Test()
    If resetDate Then Date = #1/1/2024#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
    }

    #[test]
    fn date_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Date = userDate
    If Err Then MsgBox "Invalid date"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
    }

    #[test]

    fn date_at_module_level() {
        let source = r#"
Date = #1/1/2024#
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
    }

    #[test]
    fn date_with_dateadd() {
        let source = r#"
Sub Test()
    Date = DateAdd("d", 30, Date)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
    }
}
