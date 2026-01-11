//! Call statement and procedure call parsing for VB6 CST.
//!
//! This module handles parsing of VB6 procedure invocation:
//! - `Call` - Explicit Call statement with Call keyword
//! - Procedure calls - Implicit procedure calls without Call keyword
//!
//! # Call Statement
//!
//! The Call statement explicitly invokes a Sub or Function procedure.
//! The Call keyword is optional in VB6; procedures can be called without it.
//!
//! ## Syntax
//! ```vb
//! Call procedurename [(argumentlist)]
//! procedurename [argumentlist]
//! ```
//!
//! ## Examples
//! ```vb
//! Call MySubroutine()
//! Call ProcessData(x, y, z)
//! MySubroutine              ' Without Call keyword
//! ProcessData x, y, z       ' Without Call keyword, no parentheses
//! DoSomething()             ' Without Call keyword, with parentheses
//! ```
//!
//! ## Remarks
//! - The Call keyword is optional
//! - When using Call, arguments must be enclosed in parentheses
//! - Without Call, parentheses are optional for Sub procedures
//! - For Functions, use without Call when you want the return value
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/call-statement)

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

impl Parser<'_> {
    /// Parse a Call statement:
    ///
    /// \[ Call \] name \[ argumentlist \]
    ///
    /// The Call statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Call        | Optional            | Indicates that a procedure is being called. The Call keyword is optional; if omitted, the procedure name is used directly. |
    /// | name        | Required            | Name of the procedure to be called; follows standard variable naming conventions. |
    /// | argumentlist| Optional            | List of arguments to be passed to the procedure. Arguments are enclosed in parentheses and separated by commas. |
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/call-statement)
    pub(crate) fn parse_call_statement(&mut self) {
        // if we are now parsing a call statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::CallStatement.to_raw());
        self.consume_whitespace();

        // Consume "Call" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // CallStatement
    }

    /// Parse a procedure call without the Call keyword.
    /// In VB6, you can call a Sub procedure without using the Call keyword:
    /// - `MySub arg1, arg2` instead of `Call MySub(arg1, arg2)`
    /// - `MySub` (no arguments)
    pub(crate) fn parse_procedure_call(&mut self) {
        // if we are now parsing a procedure call, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::CallStatement.to_raw());
        self.consume_whitespace();

        // Consume everything until newline (procedure name and arguments)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // CallStatement
    }

    /// Check if the current position is at a procedure call (without Call keyword).
    /// This is true if we have an identifier that's not followed by an assignment operator.
    /// In VB6, procedure calls can appear as:
    /// - `MySub` (no arguments)
    /// - `MySub arg1, arg2` (arguments without parentheses)
    /// - `MySub(arg1, arg2)` (arguments with parentheses)
    pub(crate) fn is_at_procedure_call(&self) -> bool {
        // Must start with an identifier or keyword used as identifier
        // BUT exclude keywords that have structural meaning and can't be procedure names
        if self.at_token(Token::Identifier) {
            // Identifiers are OK
        } else if self.at_keyword() {
            // Some keywords should never be treated as procedure calls
            // These are structural keywords that have special parsing rules
            if let Some(
                Token::EndKeyword
                | Token::ExitKeyword
                | Token::LoopKeyword
                | Token::NextKeyword
                | Token::WendKeyword
                | Token::ElseKeyword
                | Token::ElseIfKeyword
                | Token::CaseKeyword
                | Token::IfKeyword
                | Token::ThenKeyword
                | Token::SelectKeyword
                | Token::DoKeyword
                | Token::WhileKeyword
                | Token::UntilKeyword
                | Token::ForKeyword
                | Token::ToKeyword
                | Token::StepKeyword
                | Token::SubKeyword
                | Token::FunctionKeyword
                | Token::PropertyKeyword
                | Token::WithKeyword
                | Token::ReturnKeyword
                | Token::ResumeKeyword,
            ) = self.current_token()
            {
                return false;
            }
        } else {
            return false;
        }

        // Look ahead to see if there's an assignment operator
        // If there's an =, it's an assignment, not a procedure call
        for (_text, token) in self.tokens.iter().skip(self.pos) {
            match token {
                Token::Newline | Token::EndOfLineComment | Token::RemComment => {
                    // Reached end of line without finding assignment - this is a procedure call
                    return true;
                }
                Token::EqualityOperator => {
                    // Found = operator - this is an assignment, not a procedure call
                    return false;
                }
                // Procedure calls can have various tokens before newline
                Token::Identifier
                | Token::LeftParenthesis
                | Token::RightParenthesis
                | Token::Comma
                | Token::PeriodOperator
                | Token::StringLiteral
                | Token::IntegerLiteral
                | Token::LongLiteral
                | Token::SingleLiteral
                | Token::DoubleLiteral => {
                    // These can all appear in procedure calls, continue looking
                }
                // If it's a keyword, it could be an argument
                _ if token.is_keyword() => {
                    // Keywords can be used as arguments (e.g., True, False, Nothing)
                }
                // Whitespace or Anything else could indicate it's not a simple procedure call
                _ => {}
            }
        }

        false
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn call_statement_simple() {
        let source = "Call MySubroutine()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn call_statement_with_arguments() {
        let source = "Call ProcessData(x, y, z)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn call_statement_preserves_whitespace() {
        let source = "Call  MyFunction (  arg1 ,  arg2  )\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn call_statement_in_sub() {
        let source = "Sub Main()\nCall DoSomething()\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn call_statement_no_parentheses() {
        let source = "Call MySubroutine\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_call_statements() {
        let source = "Call First()\nCall Second()\nCall Third()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn call_statement_with_string_arguments() {
        let source = "Call ShowMessage(\"Hello, World!\")\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn call_statement_with_complex_expressions() {
        let source = "Call Calculate(x + y, z * 2, (a - b) / c)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    // Procedure call tests (without Call keyword)

    #[test]
    fn procedure_call_no_arguments() {
        let source = "InitializeRandomDNA\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn procedure_call_with_parentheses() {
        let source = "DoSomething()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn procedure_call_with_arguments_no_parentheses() {
        let source = "MsgBox \"Hello\", vbInformation, \"Title\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn procedure_call_with_arguments_with_parentheses() {
        let source = "ProcessData(x, y, z)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn multiple_procedure_calls_in_sub() {
        let source = "Sub Test()\nInitializeRandomDNA\nGetInitialSize\nGetInitialSpeed\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn procedure_call_preserves_whitespace() {
        let source = "MySub  arg1 ,  arg2\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn procedure_call_vs_assignment() {
        // This should be an assignment, not a procedure call
        let source = "x = 5\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
