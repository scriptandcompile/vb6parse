//! Option statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Option statements:
//! - Option Explicit - Require explicit variable declarations
//! - Option Base - Set default lower bound for array subscripts
//! - Option Compare - Set string comparison method
//! - Option Private - Set module visibility

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse an Option statement: Option Explicit On/Off or Option Base 0/1
    pub(super) fn parse_option_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::OptionStatement.to_raw());

        // Consume "Option" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(VB6Token::Newline);

        self.builder.finish_node(); // OptionStatement
    }

    /// Parse an Option Base statement.
    ///
    /// Sets the default lower bound for array subscripts.
    ///
    /// # Syntax
    ///
    /// | Clause | Description |
    /// |--------|-------------|
    /// | `Option Base 0` | Sets the default lower bound to 0 (default) |
    /// | `Option Base 1` | Sets the default lower bound to 1 |
    ///
    /// # Remarks
    ///
    /// The `Option Base` statement is used to set the default lower bound for array subscripts
    /// in a module. By default, VB6 uses 0 as the lower bound. Using `Option Base 1` changes
    /// this to 1 for all arrays that don't explicitly specify bounds.
    ///
    /// - Must be used at module level (before any procedures)
    /// - Only values 0 and 1 are allowed
    /// - Affects only arrays declared without explicit lower bounds
    /// - Does not affect arrays declared with explicit bounds (e.g., `Dim arr(5 To 10)`)
    ///
    /// # Examples
    ///
    /// ```vb6
    /// Option Base 1
    ///
    /// Sub Example()
    ///     Dim arr(10)  ' Lower bound is 1, upper bound is 10
    ///     arr(1) = "First element"
    /// End Sub
    /// ```
    ///
    /// # References
    ///
    /// [Microsoft Documentation](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-base-statement)
    pub(super) fn parse_option_base_statement(&mut self) {
        self.parse_option_statement();
    }

    /// Parse an Option Compare statement.
    ///
    /// Sets the string comparison method for the module.
    ///
    /// # Syntax
    ///
    /// | Clause | Description |
    /// |--------|-------------|
    /// | `Option Compare Binary` | Case-sensitive string comparison based on binary representation |
    /// | `Option Compare Text` | Case-insensitive string comparison |
    /// | `Option Compare Database` | String comparison based on database locale (Access only) |
    ///
    /// # Remarks
    ///
    /// The `Option Compare` statement is used to set the default string comparison method
    /// for a module. This affects how VB6 compares strings in operations like `=`, `<`, `>`,
    /// and in string functions.
    ///
    /// - Must be used at module level (before any procedures)
    /// - If not specified, the default is `Binary`
    /// - **Binary**: Case-sensitive comparison based on internal binary representation of characters
    /// - **Text**: Case-insensitive comparison (A = a, B = b, etc.)
    /// - **Database**: Uses database sort order (Microsoft Access only)
    ///
    /// Binary comparison is faster but case-sensitive. Text comparison is case-insensitive
    /// but may be slower. The comparison method affects:
    /// - String comparisons in If statements
    /// - `InStr` function
    /// - `StrComp` function (unless comparison argument is specified)
    /// - Select Case with string expressions
    ///
    /// # Examples
    ///
    /// ```vb6
    /// Option Compare Text
    ///
    /// Sub Example()
    ///     If "ABC" = "abc" Then  ' True with Text, False with Binary
    ///         Debug.Print "Strings are equal"
    ///     End If
    /// End Sub
    /// ```
    ///
    /// ```vb6
    /// Option Compare Binary
    ///
    /// Sub Example()
    ///     If "ABC" = "abc" Then  ' False - case sensitive
    ///         Debug.Print "This won't print"
    ///     End If
    /// End Sub
    /// ```
    ///
    /// # References
    ///
    /// [Microsoft Documentation](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-compare-statement)
    pub(super) fn parse_option_compare_statement(&mut self) {
        self.parse_option_statement();
    }

    /// Parse an Option Private statement.
    ///
    /// Controls the visibility of module-level entities (classes, functions, etc.).
    ///
    /// # Syntax
    ///
    /// | Clause | Description |
    /// |--------|-------------|
    /// | `Option Private Module` | Makes entities in the module private to the project |
    ///
    /// # Remarks
    ///
    /// The `Option Private Module` statement is used to indicate that the entire module
    /// is private to the project in which it resides. This means that the module and its
    /// public members are not available to other projects or type libraries.
    ///
    /// - Must be used at module level (at the very top of the module)
    /// - Only valid in standard modules (.bas files) and class modules (.cls files)
    /// - Does not affect the visibility of members within the same project
    /// - When used in a class module, the class cannot be created from outside the project
    /// - Has no effect in form modules (.frm files)
    ///
    /// This is particularly useful for creating helper modules or classes that should only
    /// be used internally within a project and not exposed to external projects that might
    /// reference this one.
    ///
    /// # Examples
    ///
    /// ```vb6
    /// Option Private Module
    ///
    /// ' This module's public functions are only accessible within this project
    /// Public Function InternalHelper() As String
    ///     InternalHelper = "This is private to the project"
    /// End Function
    /// ```
    ///
    /// # References
    ///
    /// [Microsoft Documentation](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-private-statement)
    pub(super) fn parse_option_private_statement(&mut self) {
        self.parse_option_statement();
    }
}

#[cfg(test)]
mod test {
    use crate::parsers::{ConcreteSyntaxTree, SyntaxKind};

    #[test]
    fn parse_option_explicit_on() {
        let source = "Option Explicit On\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit On\n");
    }

    #[test]
    fn parse_option_explicit_off() {
        let source = "Option Explicit Off\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit Off\n");
    }

    #[test]
    fn parse_option_explicit() {
        let source = "Option Explicit\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit\n");
    }

    #[test]
    fn parse_option_base_0() {
        let source = "Option Base 0\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("OptionKeyword"));
        assert!(debug.contains("BaseKeyword"));
        assert_eq!(cst.text(), "Option Base 0\n");
    }

    #[test]
    fn parse_option_base_1() {
        let source = "Option Base 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("OptionKeyword"));
        assert!(debug.contains("BaseKeyword"));
        assert_eq!(cst.text(), "Option Base 1\n");
    }

    #[test]
    fn option_base_at_module_level() {
        let source = "Option Base 1\n\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
        assert!(debug.contains("SubStatement"));
    }

    #[test]
    fn option_base_with_whitespace() {
        let source = "Option  Base  1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
        assert_eq!(cst.text(), "Option  Base  1\n");
    }

    #[test]
    fn option_base_with_comment() {
        let source = "Option Base 0 ' Set default array base\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn option_base_preserves_whitespace() {
        let source = "Option Base 1  \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "Option Base 1  \n");
    }

    #[test]
    fn multiple_option_statements() {
        let source = "Option Explicit\nOption Base 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert_eq!(cst.child_count(), 2);
        assert!(debug.contains("ExplicitKeyword"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_base_case_insensitive() {
        let source = "OPTION BASE 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_base_with_line_continuation() {
        let source = "Option _\nBase 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_base_before_declarations() {
        let source = "Option Base 1\nDim arr(10) As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn option_base_in_module() {
        let source = r#"Attribute VB_Name = "Module1"
Option Base 1
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AttributeStatement"));
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_base_0_default() {
        let source = "Option Base 0\nDim x(5) As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn parse_option_compare_binary() {
        let source = "Option Compare Binary\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("OptionKeyword"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("BinaryKeyword"));
        assert_eq!(cst.text(), "Option Compare Binary\n");
    }

    #[test]
    fn parse_option_compare_text() {
        let source = "Option Compare Text\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("OptionKeyword"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("TextKeyword"));
        assert_eq!(cst.text(), "Option Compare Text\n");
    }

    #[test]
    fn parse_option_compare_database() {
        let source = "Option Compare Database\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("OptionKeyword"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("DatabaseKeyword"));
        assert_eq!(cst.text(), "Option Compare Database\n");
    }

    #[test]
    fn option_compare_at_module_level() {
        let source = "Option Compare Text\n\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("SubStatement"));
    }

    #[test]
    fn option_compare_with_whitespace() {
        let source = "Option  Compare  Binary\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("CompareKeyword"));
        assert_eq!(cst.text(), "Option  Compare  Binary\n");
    }

    #[test]
    fn option_compare_with_comment() {
        let source = "Option Compare Text ' Case-insensitive\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn option_compare_preserves_whitespace() {
        let source = "Option Compare Binary  \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "Option Compare Binary  \n");
    }

    #[test]
    fn multiple_option_statements_with_compare() {
        let source = "Option Explicit\nOption Compare Text\nOption Base 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert_eq!(cst.child_count(), 3);
        assert!(debug.contains("ExplicitKeyword"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_compare_case_insensitive() {
        let source = "OPTION COMPARE BINARY\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("CompareKeyword"));
    }

    #[test]
    fn option_compare_with_line_continuation() {
        let source = "Option _\nCompare Text\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("CompareKeyword"));
    }

    #[test]
    fn option_compare_before_declarations() {
        let source = "Option Compare Binary\nDim str As String\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn option_compare_in_module() {
        let source = r#"Attribute VB_Name = "Module1"
Option Compare Text
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AttributeStatement"));
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("CompareKeyword"));
    }

    #[test]
    fn option_compare_text_case_insensitive_behavior() {
        let source = "Option Compare Text\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("TextKeyword"));
    }

    #[test]
    fn option_compare_binary_default() {
        let source = "Option Compare Binary\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BinaryKeyword"));
    }

    #[test]
    fn option_compare_database_access_only() {
        let source = "Option Compare Database\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("DatabaseKeyword"));
    }

    #[test]
    fn all_three_option_statements() {
        let source = "Option Explicit\nOption Compare Binary\nOption Base 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert_eq!(cst.child_count(), 3);
        assert!(debug.contains("ExplicitKeyword"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("BinaryKeyword"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn parse_option_private_module() {
        let source = "Option Private Module\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("OptionKeyword"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("ModuleKeyword"));
        assert_eq!(cst.text(), "Option Private Module\n");
    }

    #[test]
    fn option_private_at_module_level() {
        let source = "Option Private Module\n\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("SubStatement"));
    }

    #[test]
    fn option_private_with_whitespace() {
        let source = "Option  Private  Module\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert_eq!(cst.text(), "Option  Private  Module\n");
    }

    #[test]
    fn option_private_with_comment() {
        let source = "Option Private Module ' Make this module private\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn option_private_preserves_whitespace() {
        let source = "Option Private Module  \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "Option Private Module  \n");
    }

    #[test]
    fn multiple_options_with_private() {
        let source = "Option Explicit\nOption Private Module\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert_eq!(cst.child_count(), 2);
        assert!(debug.contains("ExplicitKeyword"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn option_private_case_insensitive() {
        let source = "OPTION PRIVATE MODULE\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn option_private_with_line_continuation() {
        let source = "Option _\nPrivate Module\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn option_private_before_declarations() {
        let source = "Option Private Module\nDim x As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn option_private_in_class_module() {
        let source = r#"VERSION 1.0 CLASS
Option Private Module

Public Function Test() As String
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn option_private_in_standard_module() {
        let source = r#"Attribute VB_Name = "Module1"
Option Private Module
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AttributeStatement"));
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn all_four_option_statements() {
        let source =
            "Option Explicit\nOption Compare Binary\nOption Base 1\nOption Private Module\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert_eq!(cst.child_count(), 4);
        assert!(debug.contains("ExplicitKeyword"));
        assert!(debug.contains("CompareKeyword"));
        assert!(debug.contains("BaseKeyword"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn option_private_typical_usage() {
        let source = r#"Option Private Module

Public Function InternalHelper() As String
    InternalHelper = "Internal use only"
End Function
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("FunctionStatement"));
    }
}
