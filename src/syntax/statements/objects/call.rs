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

use crate::errors::ModuleError;
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
        self.consume_whitespace();

        // Parse the callee (procedure name, which may include member access)
        self.parse_call_target();

        // With the Call keyword, arguments must be in parentheses
        self.consume_whitespace();
        if self.at_token(Token::LeftParenthesis) {
            self.parse_parenthesized_arguments();
        }

        // Consume until newline
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

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

        // Parse the callee (procedure name, which may include member access or dot-prefix)
        let is_print_like_call = self.parse_call_target();

        // Parse arguments (with or without parentheses)
        // Check if there's whitespace before the parenthesis - this is important for VB6's graphics methods.
        // In VB6:
        // - `MySub()` - zero argument call with immediate parentheses
        // - `MySub(x, y)` - call with parentheses enclosing arguments
        // - `MySub x, y` - call without parentheses
        // - `MySub (x, y)` - call with SPACE before paren means first arg is parenthesized expression
        // - `Picture1.Line (x1, y1)-(x2, y2)` - graphics method with special coordinate syntax

        // Check for whitespace BEFORE consuming it
        let has_whitespace_before_paren = self.at_token(Token::Whitespace);
        self.consume_whitespace();

        if self.at_token(Token::LeftParenthesis) {
            if has_whitespace_before_paren {
                // Space before parenthesis means the parentheses are part of argument expressions
                // This handles graphics methods like Line: Picture1.Line (x, y)-(x2, y2)
                self.parse_unparenthesized_arguments(is_print_like_call);
            } else {
                // No space means parentheses delimit the argument list
                self.parse_parenthesized_arguments();
            }
        } else if !self.at_token(Token::Newline) && !self.is_at_end() {
            // Arguments without parentheses (VB6 allows this for Sub calls)
            self.parse_unparenthesized_arguments(is_print_like_call);
        }

        // Consume until newline
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // CallStatement
    }

    /// Parse the call target (procedure name with optional member access).
    /// This handles patterns like:
    /// - `MySub`
    /// - `obj.Method`
    /// - `.Method` (in With blocks)
    fn parse_call_target(&mut self) -> bool {
        let mut last_name_is_print = false;

        // Check if this starts with a period (With block member access)
        if self.at_token(Token::PeriodOperator) {
            self.consume_token();
            self.consume_whitespace();
        }

        // Consume the identifier or keyword (VB6 allows keywords as method names)
        if self.is_identifier() || self.at_keyword() {
            last_name_is_print = self.current_token_is_print_name();
            self.consume_token();
        }

        // Handle member access chains (obj.prop.method)
        // Note: We peek ahead to see if there's a period, but we DON'T consume trailing whitespace
        // at the end of the call target. This allows parse_procedure_call to check for whitespace
        // before arguments to distinguish between `MySub()` and `MySub ()`
        loop {
            // Check if next non-whitespace token is a period
            let mut lookahead = 0;
            let mut found_whitespace = false;
            while let Some((_, token)) = self.tokens.get(self.pos + lookahead) {
                if *token == Token::Whitespace {
                    lookahead += 1;
                    found_whitespace = true;
                } else {
                    break;
                }
            }

            // If we found a period after optional whitespace, consume the whitespace and period
            if let Some((_, token)) = self.tokens.get(self.pos + lookahead) {
                if *token == Token::PeriodOperator {
                    // Consume any whitespace before the period
                    if found_whitespace {
                        for _ in 0..lookahead {
                            self.consume_token();
                        }
                    }

                    // Consume the period
                    self.consume_token();

                    // Consume whitespace after the period
                    self.consume_whitespace();

                    // Consume the member name
                    if self.is_identifier() || self.at_keyword() {
                        last_name_is_print = self.current_token_is_print_name();
                        self.consume_token();
                    } else {
                        break;
                    }
                } else {
                    // Not a period, stop here without consuming trailing whitespace
                    break;
                }
            } else {
                break;
            }
        }

        last_name_is_print
    }

    /// Returns true when the current identifier/keyword token spells `Print`.
    fn current_token_is_print_name(&self) -> bool {
        if let Some((text, token)) = self.tokens.get(self.pos) {
            *token == Token::PrintKeyword || text.eq_ignore_ascii_case("print")
        } else {
            false
        }
    }

    /// Parse arguments enclosed in parentheses.
    /// Creates an `ArgumentList` node with `Argument` children.
    fn parse_parenthesized_arguments(&mut self) {
        self.consume_token(); // (

        self.builder.start_node(SyntaxKind::ArgumentList.to_raw());
        self.consume_whitespace();

        // Parse arguments until we hit the closing parenthesis
        while !self.is_at_end() && !self.at_token(Token::RightParenthesis) {
            self.builder.start_node(SyntaxKind::Argument.to_raw());

            // Check if this is an empty argument (comma or closing paren immediately following)
            // Empty arguments are valid in VB6: Err.Raise 1, , "error message"
            if !self.at_token(Token::Comma) && !self.at_token(Token::RightParenthesis) {
                // Check for ByVal/ByRef keyword (VB6 allows overriding passing mode at call site)
                if self.at_token(Token::ByValKeyword) || self.at_token(Token::ByRefKeyword) {
                    self.consume_token();
                    self.consume_whitespace();
                }
                // Consume optional named-argument prefix: `Identifier :=`
                let _ = self.try_consume_named_argument_prefix_for_call_argument();
                // Parse the argument expression
                self.parse_expression();
            }

            self.builder.finish_node(); // Argument

            self.consume_whitespace();

            // Check for comma (more arguments)
            if self.at_token(Token::Comma) {
                self.consume_token();
                self.consume_whitespace();
            } else {
                break;
            }
        }

        self.builder.finish_node(); // ArgumentList

        // Consume closing parenthesis
        if self.at_token(Token::RightParenthesis) {
            self.consume_token();
        }
    }

    /// Parse arguments without parentheses (VB6 Sub call syntax).
    /// Creates an `ArgumentList` node with `Argument` children.
    fn parse_unparenthesized_arguments(&mut self, allow_semicolon_separator: bool) {
        self.builder.start_node(SyntaxKind::ArgumentList.to_raw());

        // Parse arguments separated by commas until newline.
        // For print-like calls (Debug.Print / Printer.Print), semicolon is also a separator.
        loop {
            if self.at_token(Token::Newline) || self.is_at_end() {
                break;
            }

            self.builder.start_node(SyntaxKind::Argument.to_raw());

            // Check if this is an empty argument (separator or newline immediately following)
            // Empty arguments are valid in VB6: Err.Raise 1, , "error message"
            if !(self.at_token(Token::Comma)
                || self.at_token(Token::Newline)
                || allow_semicolon_separator && self.at_token(Token::Semicolon))
            {
                // Check for ByVal/ByRef keyword (VB6 allows overriding passing mode at call site)
                if self.at_token(Token::ByValKeyword) || self.at_token(Token::ByRefKeyword) {
                    self.consume_token();
                    self.consume_whitespace();
                }
                // Consume optional named-argument prefix: `Identifier :=`
                let _ = self.try_consume_named_argument_prefix_for_call_argument();
                // Parse the argument expression
                self.parse_expression();
            }

            self.builder.finish_node(); // Argument

            self.consume_whitespace();

            // Check for separator (more arguments). Print-like calls can use ';'.
            if self.at_token(Token::Comma)
                || (allow_semicolon_separator && self.at_token(Token::Semicolon))
            {
                self.consume_token();
                self.consume_whitespace();
            } else {
                if !allow_semicolon_separator && self.at_token(Token::Semicolon) {
                    self.report_error(ModuleError::InvalidSemicolonSeparatorInProcedureCall);
                }
                break;
            }
        }

        self.builder.finish_node(); // ArgumentList
    }

    /// Try to consume a VB6 named-argument prefix (`name :=`) at the start
    /// of a call-statement argument. Returns true if consumed.
    fn try_consume_named_argument_prefix_for_call_argument(&mut self) -> bool {
        let mut idx = self.pos;

        let Some((_, first_token)) = self.tokens.get(idx) else {
            return false;
        };

        if !(*first_token == Token::Identifier || first_token.is_keyword()) {
            return false;
        }

        idx += 1;
        while let Some((_, Token::Whitespace)) = self.tokens.get(idx) {
            idx += 1;
        }

        if self.tokens.get(idx).map(|(_, token)| *token) != Some(Token::ColonOperator) {
            return false;
        }

        idx += 1;
        while let Some((_, Token::Whitespace)) = self.tokens.get(idx) {
            idx += 1;
        }

        if self.tokens.get(idx).map(|(_, token)| *token) != Some(Token::EqualityOperator) {
            return false;
        }

        while self.pos <= idx {
            self.consume_token();
        }

        self.consume_whitespace();
        true
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
        // If there's an = at depth 0 (not inside parentheses) AND we haven't seen other operators,
        // it's an assignment, not a procedure call
        let mut paren_depth: i32 = 0;
        let mut seen_other_operator = false;
        let mut previous_non_whitespace: Option<Token> = None;
        for (_text, token) in self.tokens.iter().skip(self.pos) {
            match token {
                Token::Newline | Token::EndOfLineComment | Token::RemComment => {
                    // Reached end of line without finding assignment - this is a procedure call
                    return true;
                }
                Token::Whitespace => {
                    // Skip whitespace in lookahead analysis.
                    continue;
                }
                Token::LeftParenthesis => {
                    paren_depth += 1;
                }
                Token::RightParenthesis => {
                    paren_depth = paren_depth.saturating_sub(1);
                }
                Token::EqualityOperator if paren_depth == 0 => {
                    // Named arguments use `:=` and must not be treated as assignment.
                    if previous_non_whitespace == Some(Token::ColonOperator) {
                        previous_non_whitespace = Some(*token);
                        continue;
                    }
                    // Found = operator at depth 0
                    // If we've seen other operators (like >=, And, Or), this is part of an expression in a procedure call
                    // Otherwise, it's an assignment
                    return seen_other_operator;
                }
                // Track operators that indicate we're in an expression context
                Token::AndKeyword
                | Token::OrKeyword
                | Token::XorKeyword
                | Token::EqvKeyword
                | Token::ImpKeyword
                | Token::ModKeyword
                | Token::NotKeyword
                | Token::LessThanOperator
                | Token::GreaterThanOperator
                | Token::LessThanOrEqualOperator
                | Token::GreaterThanOrEqualOperator
                | Token::InequalityOperator
                | Token::AdditionOperator
                | Token::SubtractionOperator
                | Token::MultiplicationOperator
                | Token::DivisionOperator
                | Token::BackwardSlashOperator
                | Token::ExponentiationOperator
                | Token::Ampersand => {
                    seen_other_operator = true;
                }
                // All other tokens can appear in procedure calls, continue looking
                _ => {}
            }

            previous_non_whitespace = Some(*token);
        }

        // Reached end of input without finding assignment or newline - this is a procedure call
        true
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
    fn procedure_call_with_byval_arguments() {
        let source = "RtlMoveMemory vtbl, ByVal pEnumerator, 4\n";
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
    fn call_statement_with_byval_arguments() {
        let source = "Call RtlMoveMemory(vtbl, ByVal pEnumerator, 4)\n";
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

    #[test]
    fn procedure_call_debug_print_trailing_semicolon() {
        let source = "Sub Test()\n    Debug.Print Hex(i);\nEnd Sub\n";
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
    fn procedure_call_debug_print_semicolon_separated_arguments() {
        let source = "Sub Test()\n    Debug.Print \"A\"; \"B\"\nEnd Sub\n";
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
    fn procedure_call_output_object_print_trailing_semicolon() {
        let source = "Sub Test()\n    Printer.Print \"A\";\nEnd Sub\n";
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
    fn procedure_call_non_print_semicolon_not_separator() {
        let source = "Sub Test()\n    Foo 1; 2\nEnd Sub\n";
        let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let failure = &failures[0];
        assert!(
            matches!(
                failure.kind.as_ref(),
                crate::errors::ErrorKind::Module(
                    crate::errors::ModuleError::InvalidSemicolonSeparatorInProcedureCall
                )
            ),
            "Expected parser failure for ';' in non-print procedure call"
        );

        // "Sub Test()\n" is 11 bytes and ';' is the 10th byte on line 2 (0-based index 9).
        assert_eq!(failure.error_offset, 20);
        assert_eq!(failure.line_start, 2);
        assert_eq!(failure.line_end, 2);

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../snapshots/syntax/statements/objects/call");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn procedure_call_with_named_argument() {
        let source = "Sub Test()\n    pvRecvBody baBuffer, Flush:=True\nEnd Sub\n";
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
