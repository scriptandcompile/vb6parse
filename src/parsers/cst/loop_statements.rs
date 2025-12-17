//! Do/Loop statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 loop statements:
//! - Do While...Loop
//! - Do Until...Loop
//! - Do...Loop While
//! - Do...Loop Until
//! - Do...Loop (infinite loop)
//! - While...Wend

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Do...Loop statement.
    ///
    /// VB6 supports several forms of Do loops:
    /// - Do While condition...Loop
    /// - Do Until condition...Loop
    /// - Do...Loop While condition
    /// - Do...Loop Until condition
    /// - Do...Loop (infinite loop, requires Exit Do)
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doloop-statement)
    pub(super) fn parse_do_statement(&mut self) {
        // if we are now parsing a do statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::DoStatement.to_raw());

        // Consume "Do" keyword
        self.consume_token();

        // Consume whitespace after Do
        self.consume_whitespace();

        // Check if we have While or Until after Do
        let has_top_condition =
            self.at_token(Token::WhileKeyword) || self.at_token(Token::UntilKeyword);

        if has_top_condition {
            // Consume While or Until
            self.consume_token();

            // Consume any whitespace after While or Until
            self.consume_whitespace();

            // Parse condition - consume everything until newline
            self.parse_expression();
        }

        // Consume newline after Do line
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Loop"
        self.parse_code_block(|parser| parser.at_token(Token::LoopKeyword));

        // Consume "Loop" keyword
        if self.at_token(Token::LoopKeyword) {
            self.consume_token();

            // Consume whitespace after Loop
            self.consume_whitespace();

            // Check if we have While or Until after Loop
            if self.at_token(Token::WhileKeyword) || self.at_token(Token::UntilKeyword) {
                // Consume While or Until
                self.consume_token();

                // Consume any whitespace after While or Until
                self.consume_whitespace();

                // Parse condition - consume everything until newline
                self.parse_expression();
            }

            // Consume newline after Loop
            if self.at_token(Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // DoStatement
    }

    /// Parse a While...Wend statement.
    ///
    /// VB6 While...Wend loop syntax:
    /// - While condition
    ///   ...statements...
    ///   Wend
    ///
    /// While...Wend statement syntax:
    ///
    /// | Part      | Description |
    /// |-----------|-------------|
    /// | condition | Required. Numeric or String expression that evaluates to True or False. If condition is Null, condition is treated as False. |
    /// | statements| Optional. One or more statements executed while condition is True. |
    ///
    /// Remarks:
    /// - If condition is True, all statements are executed until the Wend statement is encountered.
    /// - Control then returns to the While statement and condition is again checked.
    /// - If condition is still True, the process is repeated. If it's not True, execution resumes with the statement following the Wend statement.
    /// - While...Wend loops can be nested to any level. Each Wend matches the most recent While.
    /// - Note: The Do...Loop statement provides a more structured and flexible way to perform looping.
    /// - Tip: While...Wend is provided for compatibility with earlier versions of Visual Basic. Consider using Do...Loop instead for new code.
    ///
    /// Examples:
    /// ```vb
    /// Dim counter As Integer
    /// counter = 0
    /// While counter < 20
    ///     counter = counter + 1
    ///     Debug.Print counter
    /// Wend
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/whilewend-statement)
    /// Parse a While...Wend statement.
    ///
    /// While...Wend is a legacy VB6 loop construct that executes a block of
    /// statements while a condition is true. It has been superseded by Do While...Loop
    /// but is still supported for backward compatibility.
    ///
    /// Syntax:
    /// ```vb6
    /// While condition
    ///     statements
    /// Wend
    /// ```
    ///
    /// Example:
    /// ```vb6
    /// While x < 10
    ///     x = x + 1
    /// Wend
    /// ```
    ///
    /// The condition is evaluated before each iteration. If the condition is
    /// initially false, the loop body will not execute at all.
    pub(super) fn parse_while_statement(&mut self) {
        // if we are now parsing a while statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::WhileStatement.to_raw());

        // Consume "While" keyword
        self.consume_token();

        // Consume whitespace after While
        self.consume_whitespace();

        // Parse condition - consume everything until newline
        self.parse_expression();

        // Consume newline after While line
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Wend"
        self.parse_code_block(|parser| parser.at_token(Token::WendKeyword));

        // Consume "Wend" keyword
        if self.at_token(Token::WendKeyword) {
            self.consume_token();

            // Consume newline after Wend
            if self.at_token(Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // WhileStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    // While...Wend statement tests
    #[test]
    fn while_simple() {
        let source = r#"
Sub Test()
    While x < 10
        x = x + 1
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
        assert!(debug.contains("WhileKeyword"));
        assert!(debug.contains("WendKeyword"));
    }

    #[test]
    fn while_at_module_level() {
        let source = r#"
While x < 5
    x = x + 1
Wend
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_string_condition() {
        let source = r#"
Sub Test()
    While inputText <> ""
        ProcessInput
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_not_eof() {
        let source = r#"
Sub Test()
    While Not EOF(1)
        Line Input #1, textLine
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_nested() {
        let source = r#"
Sub Test()
    While i < 10
        While j < 5
            j = j + 1
        Wend
        i = i + 1
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WhileStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn while_with_exit() {
        let source = r#"
Sub Test()
    While True
        If x > 100 Then Exit Do
        x = x + 1
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_comment() {
        let source = r#"
Sub Test()
    While count < limit ' Loop until limit
        count = count + 1
    Wend ' End of loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn while_with_complex_condition() {
        let source = r#"
Sub Test()
    While (x < 10 And y > 0) Or z = 5
        Process
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_function_call() {
        let source = r#"
Sub Test()
    While IsValid(data)
        ProcessData data
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_property_access() {
        let source = r#"
Sub Test()
    While rs.EOF = False
        rs.MoveNext
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_empty_body() {
        let source = r#"
Sub Test()
    While condition
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_doevents() {
        let source = r#"
Sub Test()
    While processing
        DoEvents
        CheckStatus
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_string_length_check() {
        let source = r#"
Sub Test()
    While Len(text) > 0
        text = Mid(text, 2)
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_array_access() {
        let source = r#"
Sub Test()
    While arr(index) <> 0
        index = index + 1
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_if_statement() {
        let source = r#"
Sub Test()
    While active
        If condition Then
            Process
        End If
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn while_with_select_case() {
        let source = r#"
Sub Test()
    While running
        Select Case action
            Case 1
                DoAction1
            Case 2
                DoAction2
        End Select
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_for_loop() {
        let source = r#"
Sub Test()
    While outerCondition
        For i = 1 To 10
            Process i
        Next i
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn while_triple_nested() {
        let source = r#"
Sub Test()
    While a < 10
        While b < 5
            While c < 3
                c = c + 1
            Wend
            b = b + 1
        Wend
        a = a + 1
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("WhileStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn while_with_comparison_operators() {
        let source = r#"
Sub Test()
    While value <= maxValue
        value = value * 2
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_boolean_literal() {
        let source = r#"
Sub Test()
    While True
        If userQuit Then Exit Sub
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_multiple_statements() {
        let source = r#"
Sub Test()
    While counter < 100
        counter = counter + 1
        sum = sum + counter
        average = sum / counter
        Display average
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_reading_file() {
        let source = r#"
Sub ReadFile()
    Open "data.txt" For Input As #1
    While Not EOF(1)
        Line Input #1, dataLine
        ProcessLine dataLine
    Wend
    Close #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_parenthesized_condition() {
        let source = r#"
Sub Test()
    While (counter < limit)
        counter = counter + 1
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_recordset_iteration() {
        let source = r#"
Sub Test()
    Set rs = db.OpenRecordset("Table1")
    While Not rs.EOF
        Debug.Print rs!FieldName
        rs.MoveNext
    Wend
    rs.Close
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_timer_based() {
        let source = r#"
Sub Test()
    startTime = Timer
    While Timer - startTime < 5
        DoEvents
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_with_msgbox() {
        let source = r#"
Sub Test()
    While confirm = vbYes
        Process
        confirm = MsgBox("Continue?", vbYesNo)
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_collection_count() {
        let source = r#"
Sub Test()
    While col.Count > 0
        col.Remove 1
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_dir_function() {
        let source = r#"
Sub Test()
    fileName = Dir("*.txt")
    While fileName <> ""
        ProcessFile fileName
        fileName = Dir
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_instr_search() {
        let source = r#"
Sub Test()
    position = InStr(text, searchTerm)
    While position > 0
        FoundAt position
        position = InStr(position + 1, text, searchTerm)
    Wend
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn while_preserves_whitespace() {
        let source = "    While    x <    10    \n        x = x + 1\n    Wend    \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert!(cst.text().contains("While    x <    10"));
        let debug = cst.debug_tree();
        assert!(debug.contains("WhileStatement"));
    }

    // Do...Loop statement tests
    #[test]
    fn do_while_loop() {
        let source = r#"
Sub Test()
    Do While x < 10
        x = x + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("WhileKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_until_loop() {
        let source = r#"
Sub Test()
    Do Until x >= 10
        x = x + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("UntilKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_loop_while() {
        let source = r#"
Sub Test()
    Do
        x = x + 1
    Loop While x < 10
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("WhileKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_loop_until() {
        let source = r#"
Sub Test()
    Do
        x = x + 1
    Loop Until x >= 10
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("UntilKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_loop_infinite() {
        let source = r#"
Sub Test()
    Do
        If x > 10 Then Exit Do
        x = x + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn nested_do_loops() {
        let source = r#"
Sub Test()
    Do While i < 10
        Do While j < 5
            j = j + 1
        Loop
        i = i + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        // Should have two DoStatements
        let do_count = debug.matches("DoStatement").count();
        assert_eq!(do_count, 2);
    }

    #[test]
    fn do_while_with_complex_condition() {
        let source = r#"
Sub Test()
    Do While x < 10 And y > 0
        x = x + 1
        y = y - 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("WhileKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_loop_preserves_whitespace() {
        let source = r#"
Sub Test()
    Do  While  x < 10
        x = x + 1
    Loop  While  y > 0
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        // Check that whitespace is preserved
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn function_with_do_loop_ending_at_end_function() {
        let source = r#"Function Test()
Do
Loop
End Function
"#;

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Check if "End" appears as Unknown in the tree
        let tree_str = cst.debug_tree();
        assert!(
            !tree_str.contains("Unknown"),
            "Should not have any Unknown tokens\n{}",
            tree_str
        );
    }

    #[test]
    fn function_with_do_until_loop() {
        let source = r#"Function Test()
Do Until x = ""
  y = z
Loop
End Function
"#;

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Check if "End" appears as Unknown in the tree
        let tree_str = cst.debug_tree();
        assert!(
            !tree_str.contains("Unknown"),
            "Should not have any Unknown tokens\n{}",
            tree_str
        );
    }
}
