//! End statement parsing for VB6.

use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a standalone `End` statement.
    ///
    /// The `End` statement terminates program execution immediately.
    /// It closes all files opened using the `Open` statement and clears all variables.
    ///
    /// Syntax:
    ///   `End`
    ///
    /// Note: This is distinct from compound `End` keywords like `End If`, `End Sub`,
    /// `End Function`, etc., which are block terminators handled by their respective parsers.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/end-statement)
    pub(crate) fn parse_end_statement(&mut self) {
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::EndStatement.to_raw());
        self.consume_whitespace();

        // Consume "End" keyword
        self.consume_token();

        self.builder.finish_node(); // EndStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn end_standalone_in_sub() {
        let source = r"
Sub Test()
    End
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/end");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn end_standalone_in_if_block() {
        let source = r#"
Sub Test()
    If x > 0 Then
        End
    Else
        MsgBox "negative"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/end");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn end_standalone_at_module_level() {
        let source = "End\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/control_flow/end");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
