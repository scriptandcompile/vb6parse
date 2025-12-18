use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a VB6 Randomize statement.
    ///
    /// # Syntax
    ///
    /// ```vb
    /// Randomize [number]
    /// ```
    ///
    /// # Arguments
    ///
    /// | Part | Optional / Required | Description |
    /// |------|---------------------|-------------|
    /// | number | Optional | A Variant or any valid numeric expression that is used as the new seed value to initialize the random number generator. |
    ///
    /// # Remarks
    ///
    /// - The Randomize statement initializes the random-number generator, giving it a new seed value.
    /// - If you omit number, the value returned by the system timer is used as the new seed value.
    /// - If Randomize is not used, the Rnd function (with no arguments) uses the same number as a seed the first time it is called, and thereafter uses the last generated number as a seed value.
    /// - To repeat sequences of random numbers, call Rnd with a negative argument immediately before using Randomize with a numeric argument.
    /// - Using Randomize with the same value for number does not repeat the previous sequence.
    ///
    /// # Examples
    ///
    /// ```vb
    /// ' Initialize random number generator
    /// Randomize
    /// x = Int((100 * Rnd) + 1)
    ///
    /// ' Initialize with specific seed
    /// Randomize 42
    /// x = Rnd
    ///
    /// ' Use timer as seed
    /// Randomize Timer
    /// ```
    ///
    /// # References
    ///
    /// [Microsoft VBA Language Reference - Randomize Statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/randomize-statement)
    pub(crate) fn parse_randomize_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::RandomizeStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn randomize_simple() {
        let source = r"
Sub Test()
    Randomize
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("RandomizeKeyword"));
    }

    #[test]
    fn randomize_with_seed() {
        let source = r"
Sub Test()
    Randomize 42
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("RandomizeKeyword"));
        assert!(debug.contains("42"));
    }

    #[test]
    fn randomize_with_timer() {
        let source = r"
Sub Test()
    Randomize Timer
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("Timer"));
    }

    #[test]
    fn randomize_with_expression() {
        let source = r"
Sub Test()
    Randomize x * 100 + 42
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_with_variable() {
        let source = r"
Sub Test()
    Dim seed As Long
    seed = 12345
    Randomize seed
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("seed"));
    }

    #[test]
    fn randomize_in_if_statement() {
        let source = r"
Sub Test()
    If needsRandom Then
        Randomize
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn randomize_inline_if() {
        let source = r"
Sub Test()
    If initialize Then Randomize
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_with_comment() {
        let source = r"
Sub Test()
    Randomize ' Initialize RNG
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn randomize_at_module_level() {
        let source = "Randomize\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_preserves_whitespace() {
        let source = "    Randomize    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Randomize    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Randomize i
        x = Rnd
    Next i
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn multiple_randomize_statements() {
        let source = r"
Sub Test()
    Randomize
    x = Rnd
    Randomize 42
    y = Rnd
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("RandomizeStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn randomize_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Randomize
    If Err.Number <> 0 Then
        MsgBox "Error initializing RNG"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn randomize_with_negative_seed() {
        let source = r"
Sub Test()
    Randomize -1
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_with_function_call() {
        let source = r"
Sub Test()
    Randomize GetSeed()
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("GetSeed"));
    }

    #[test]
    fn randomize_before_rnd() {
        let source = r"
Function GetRandomNumber() As Integer
    Randomize
    GetRandomNumber = Int((100 * Rnd) + 1)
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("FunctionStatement"));
    }

    #[test]
    fn randomize_with_select_case() {
        let source = r"
Sub Test()
    Select Case mode
        Case 1
            Randomize
        Case 2
            Randomize Timer
        Case Else
            Randomize 0
    End Select
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("RandomizeStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn randomize_with_do_loop() {
        let source = r"
Sub Test()
    Do While True
        Randomize
        x = Rnd
        If x > 0.9 Then Exit Do
    Loop
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
        assert!(debug.contains("DoStatement"));
    }

    #[test]
    fn randomize_multiline_with_continuation() {
        let source = r"
Sub Test()
    Randomize _
        Timer
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_with_parentheses() {
        let source = r"
Sub Test()
    Randomize (seed)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_in_class_module() {
        let source = r"
Private Sub Class_Initialize()
    Randomize
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_with_decimal_seed() {
        let source = r"
Sub Test()
    Randomize 123.456
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }

    #[test]
    fn randomize_with_multiple_operations() {
        let source = r"
Sub GenerateRandomNumbers()
    Randomize Timer
    Dim nums(10) As Integer
    Dim i As Integer
    For i = 1 To 10
        nums(i) = Int((100 * Rnd) + 1)
    Next i
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RandomizeStatement"));
    }
}
