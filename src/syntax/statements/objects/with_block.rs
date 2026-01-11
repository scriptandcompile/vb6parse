//! With statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 With blocks for simplified object property access:
//! - `With...End With` - Reference an object multiple times without repeating its name
//!
//! # With Statement
//!
//! The With statement allows you to perform a series of statements on a specified object
//! without requalifying the name of the object. You can use `.PropertyName` or `.MethodName`
//! within the With block to access members of the object.
//!
//! ## Syntax
//! ```vb
//! With object
//!     .Property1 = value1
//!     .Property2 = value2
//!     .Method arg1, arg2
//! End With
//! ```
//!
//! ## Examples
//! ```vb
//! With myObject
//!     .Property = "value"
//!     .NestedObject.Property = 123
//!     .Method "arg"
//! End With
//!
//! ' Nested With blocks
//! With obj1
//!     With .NestedObj
//!         .Value = 10
//!     End With
//! End With
//!
//! ' With New keyword
//! With New MyClass
//!     .Initialize
//! End With
//! ```
//!
//! ## Remarks
//! - Statements in the With block access members using the dot prefix (`.`)
//! - With blocks can be nested
//! - The object reference is evaluated once at the beginning
//! - Can be used with the New keyword to create and initialize objects
//! - Improves code readability and can improve performance
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/with-statement)

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a With statement.
    ///
    /// VB6 With statement syntax:
    ///
    /// ```vb
    /// With object
    ///     .Property1 = value1
    ///     .Property2 = value2
    /// End With
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/with-statement)
    pub(crate) fn parse_with_statement(&mut self) {
        // if we are now parsing a with statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::WithStatement.to_raw());
        self.consume_whitespace();

        // Consume "With" keyword
        self.consume_token();

        // Consume everything until newline (the object expression)
        self.consume_until_after(Token::Newline);

        // Parse the body until "End With"
        self.parse_statement_list(|parser| {
            parser.at_token(Token::EndKeyword)
                && parser.peek_next_keyword() == Some(Token::WithKeyword)
        });

        // Consume "End With" and trailing tokens
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "With"
            self.consume_whitespace();

            // Consume "With"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // WithStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn with_statement_simple() {
        let source = r"
Sub Test()
    With myObject
        .Property = 123
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_nested_property() {
        let source = r"
Sub Test()
    With myObject
        .NestedObject.Property = 123
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_method_call() {
        let source = r#"
Sub Test()
    With myObject
        .Method "argument"
    End With
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_nested() {
        let source = r"
Sub Test()
    With obj1
        .Property1 = 1
        With .NestedObj
            .Property2 = 2
        End With
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_multiple_properties() {
        let source = r"
Sub Test()
    With myObject
        .Property1 = 1
        .Property2 = 2
        .Property3 = 3
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_with_if() {
        let source = r"
Sub Test()
    With myObject
        If .Property > 0 Then
            .Method
        End If
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_with_loop() {
        let source = r"
Sub Test()
    With myObject
        For i = 1 To 10
            .Process i
        Next i
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_array_access() {
        let source = r"
Sub Test()
    With myObject(index)
        .Property = 123
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_function_result() {
        let source = r"
Sub Test()
    With GetObject()
        .Property = 123
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_empty() {
        let source = r"
Sub Test()
    With myObject
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_sequential() {
        let source = r"
Sub Test()
    With obj1
        .Property = 1
    End With
    With obj2
        .Property = 2
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_preserves_whitespace() {
        let source = r"
Sub Test()
    With   myObject
        .Property   =   123
    End   With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_new_object() {
        let source = r"
Sub Test()
    With New MyClass
        .Property = 123
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn with_statement_at_module_level() {
        let source = r"
With globalObject
    .Property = 123
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/with_block");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
