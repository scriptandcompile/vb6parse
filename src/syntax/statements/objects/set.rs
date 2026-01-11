//! Set statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Set statements for object reference assignment:
//! - `Set` - Assign an object reference to a variable
//!
//! # Set Statement
//!
//! The Set statement assigns an object reference to a variable or property.
//! In VB6, object variables must be assigned using Set (unlike value types which use `=` alone).
//!
//! ## Syntax
//! ```vb
//! Set objectVar = [New] objectExpression
//! Set objectVar = Nothing
//! ```
//!
//! ## Examples
//! ```vb
//! Set obj = myObject
//! Set obj = New MyClass
//! Set obj = Nothing
//! Set myObj.Property = otherObj
//! Set result = GetObject("WinMgmts:")
//! Set item = collection.Item(1)
//! ```
//!
//! ## Remarks
//! - Required for assigning object references (not for value types)
//! - Can use `New` keyword to create a new instance
//! - Use `Nothing` to release an object reference
//! - Assignment operator `=` follows the Set keyword
//! - Works with properties and collection items
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement)

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a Set statement.
    ///
    /// VB6 Set statement syntax:
    /// - Set objectVar = [New] objectExpression
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement)
    pub(crate) fn parse_set_statement(&mut self) {
        // if we are now parsing a set statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::SetStatement.to_raw());
        self.consume_whitespace();

        // Consume "Set" keyword
        self.consume_token();

        // Consume everything until newline
        // This includes: variable, "=", [New], object expression
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // SetStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn set_statement_simple() {
        let source = r"
Sub Test()
    Set obj = myObject
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn set_statement_with_new() {
        let source = r"
Sub Test()
    Set obj = New MyClass
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn set_statement_to_nothing() {
        let source = r"
Sub Test()
    Set obj = Nothing
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn set_statement_with_property_access() {
        let source = r"
Sub Test()
    Set myObj.Property = otherObj
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn set_statement_with_function_call() {
        let source = r#"
Sub Test()
    Set result = GetObject("WinMgmts:")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn set_statement_with_collection_access() {
        let source = r"
Sub Test()
    Set item = collection.Item(1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_set_statements() {
        let source = r"
Sub Test()
    Set obj1 = New Class1
    Set obj2 = New Class2
    Set obj3 = Nothing
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn set_statement_preserves_whitespace() {
        let source = r"
Sub Test()
    Set   obj   =   New   MyClass
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn set_statement_in_function() {
        let source = r"
Function GetObject() As Object
    Set GetObject = New MyClass
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn set_statement_at_module_level() {
        let source = r"
Set globalObj = New MyClass
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/set");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
