//! Array statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 array statements:
//! - Variable declarations (Dim, Private, Public, Const, Static)
//! - ReDim - Reallocate storage space for dynamic array variables
//! - Erase - Reinitialize the elements of fixed-size arrays and deallocate dynamic arrays

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse a ReDim statement.
    ///
    /// VB6 ReDim statement syntax:
    /// - ReDim [Preserve] varname(subscripts) [As type] [, varname(subscripts) [As type]] ...
    ///
    /// Used at procedure level to reallocate storage space for dynamic array variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    pub(super) fn parse_redim_statement(&mut self) {
        // if we are now parsing a ReDim statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ReDimStatement.to_raw());

        // Consume "ReDim" keyword
        self.consume_token();

        // Consume everything until newline (Preserve, variable declarations, etc.)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ReDimStatement
    }

    /// Parse a Dim statement: Dim/Private/Public/Const/Static x As Type
    ///
    /// VB6 variable declaration statement syntax:
    /// - Dim varname [As type]
    /// - Private varname [As type]
    /// - Public varname [As type]
    /// - Const constname = expression
    /// - Static varname [As type]
    ///
    /// Used to declare variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dim-statement)
    pub(super) fn parse_dim(&mut self) {
        // if we are now parsing a dim statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::DimStatement.to_raw());

        // Consume the keyword (Dim, Private, Public, Const, Static, etc.)
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

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
    pub(super) fn parse_erase_statement(&mut self) {
        // if we are now parsing an erase statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::EraseStatement.to_raw());

        // Consume "Erase" keyword
        self.consume_token();

        // Consume everything until newline (array names, commas, etc.)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // EraseStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn redim_simple_array() {
        let source = r#"
Sub Test()
    ReDim myArray(10)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
    }

    #[test]
    fn redim_with_preserve() {
        let source = r#"
Sub Test()
    ReDim Preserve argv(argc - 1&)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
        assert!(debug.contains("PreserveKeyword"));
    }

    #[test]
    fn redim_with_as_type() {
        let source = r#"
Sub Test()
    ReDim ICI(1 To num) As ImageCodecInfo
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
        assert!(debug.contains("AsKeyword"));
    }

    #[test]
    fn redim_preserve_with_as_type() {
        let source = r#"
Sub Test()
    ReDim Preserve fileNameArray(rdIconMaximum) As String
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
        assert!(debug.contains("PreserveKeyword"));
        assert!(debug.contains("AsKeyword"));
    }

    #[test]
    fn redim_zero_based() {
        let source = r#"
Sub Test()
    ReDim argv(0&)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
    }

    #[test]
    fn redim_with_to_clause() {
        let source = r#"
Sub Test()
    ReDim hIcon(lIconIndex To lIconIndex + nIcons * 2 - 1)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
        assert!(debug.contains("ToKeyword"));
    }

    #[test]
    fn redim_multiple_arrays() {
        let source = r#"
Sub Test()
    ReDim arr1(10), arr2(20), arr3(30)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
    }

    #[test]
    fn redim_in_if_statement() {
        let source = r#"
Sub Test()
    If needResize Then ReDim myArray(newSize)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
    }

    #[test]
    fn redim_with_comment() {
        let source = r#"
Sub Test()
    ReDim Preserve fileNameArray(rdIconMaximum) As String ' the file location of the original icons
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
        assert!(debug.contains("EndOfLineComment"));
    }

    #[test]
    fn redim_multiple_in_sequence() {
        let source = r#"
Sub Test()
    ReDim Preserve fileNameArray(rdIconMaximum) As String
    ReDim Preserve dictionaryLocationArray(rdIconMaximum) As String
    ReDim Preserve namesListArray(rdIconMaximum) As String
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let redim_count = debug.matches("ReDimStatement").count();
        assert_eq!(redim_count, 3, "Expected 3 ReDim statements");
    }

    #[test]
    fn redim_in_multiline_if() {
        let source = r#"
Sub Test()
    If arraysNeedResize Then
        ReDim Preserve myArray(newSize)
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
    }

    #[test]
    fn redim_with_expression_bounds() {
        let source = r#"
Sub Test()
    ReDim Buffer(1 To Size) As Byte
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
    }

    #[test]
    fn redim_at_module_level() {
        let source = r#"
ReDim globalArray(100)
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
        assert!(debug.contains("ReDimKeyword"));
    }

    #[test]
    fn redim_multidimensional() {
        let source = r#"
Sub Test()
    ReDim matrix(10, 20)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ReDimStatement"));
    }

    // Dim statement tests
    #[test]
    fn dim_simple_declaration() {
        let source = "Dim x As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Dim x As Integer\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn dim_private_declaration() {
        let source = "Private m_value As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Private m_value As Long\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn dim_public_declaration() {
        let source = "Public g_config As String\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Public g_config As String\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("PublicKeyword"));
    }

    #[test]
    fn dim_multiple_variables() {
        let source = "Dim x, y, z As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Dim x, y, z As Integer\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn dim_const_declaration() {
        let source = "Const MAX_SIZE = 100\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Const MAX_SIZE = 100\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("ConstKeyword"));
    }

    #[test]
    fn dim_private_const() {
        let source = "Private Const MODULE_NAME = \"MyModule\"\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Private Const MODULE_NAME = \"MyModule\"\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("ConstKeyword"));
    }

    #[test]
    fn dim_static_declaration() {
        let source = "Static counter As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Static counter As Long\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("StaticKeyword"));
    }

    // Erase statement tests
    #[test]
    fn erase_simple_array() {
        let source = r#"
Sub Test()
    Erase myArray
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
        assert!(debug.contains("EraseKeyword"));
    }

    #[test]
    fn erase_multiple_arrays() {
        let source = r#"
Sub Test()
    Erase array1, array2, array3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
        assert!(debug.contains("array1"));
        assert!(debug.contains("array2"));
        assert!(debug.contains("array3"));
    }

    #[test]
    fn erase_at_module_level() {
        let source = "Erase globalArray\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
    }

    #[test]
    fn erase_preserves_whitespace() {
        let source = "    Erase    myArray    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Erase    myArray    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
    }

    #[test]
    fn erase_with_comment() {
        let source = r#"
Sub Test()
    Erase tempArray ' Free up memory
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn erase_in_if_statement() {
        let source = r#"
Sub Cleanup()
    If shouldClear Then
        Erase dataArray
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
    }

    #[test]
    fn erase_inline_if() {
        let source = r#"
Sub Test()
    If resetFlag Then Erase buffer
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
    }

    #[test]
    fn erase_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        Erase tempArrays(i)
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
    }

    #[test]
    fn erase_with_parentheses() {
        let source = r#"
Sub Test()
    Erase myArray()
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
    }

    #[test]
    fn multiple_erase_statements() {
        let source = r#"
Sub Test()
    Erase array1
    DoSomething
    Erase array2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let erase_count = debug.matches("EraseStatement").count();
        assert_eq!(erase_count, 2);
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
    }

    #[test]
    fn erase_after_redim() {
        let source = r#"
Sub Test()
    ReDim myArray(100)
    ' Use the array
    Erase myArray
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
        assert!(debug.contains("ReDimStatement"));
    }

    #[test]
    fn erase_complex_array_list() {
        let source = r#"
Sub Test()
    Erase buffer1, buffer2, cache(), tempData
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("EraseStatement"));
        assert!(debug.contains("buffer1"));
        assert!(debug.contains("cache"));
    }
}
