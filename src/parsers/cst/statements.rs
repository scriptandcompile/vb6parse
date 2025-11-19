//! Statement parsing for VB6 CST.
//!
//! This module handles parsing of all VB6 statement types:
//! - Control flow statements (If, Do, For, While, Select Case)
//! - Object manipulation (Call, Set, With)
//! - Program flow (GoTo, Exit, Label)
//! - Array operations (ReDim)
//! - System commands (AppActivate, Beep, ChDir, ChDrive)

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse an If statement: If condition Then ... End If
    /// Handles both single-line and multi-line If statements
    ///
    /// IfStatement
    /// ├─ If keyword
    /// ├─ condition tokens
    /// ├─ Then keyword
    /// ├─ body tokens
    /// ├─ ElseIfClause (if present)
    /// │  ├─ ElseIf keyword
    /// │  ├─ condition tokens
    /// │  ├─ Then keyword
    /// │  └─ body tokens
    /// ├─ ElseClause (if present)
    /// │  ├─ Else keyword
    /// │  └─ body tokens
    /// ├─ End keyword
    /// └─ If keyword
    ///
    pub(super) fn parse_if_statement(&mut self) {
        self.builder.start_node(SyntaxKind::IfStatement.to_raw());

        // Consume "If" keyword
        self.consume_token();

        // Parse the conditional expression
        self.parse_conditional();

        // Consume "Then" if present
        if self.at_token(VB6Token::ThenKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Then
        self.consume_whitespace();

        // Check if this is a single-line If statement (has code on the same line after Then)
        let is_single_line = !self.at_token(VB6Token::Newline) && !self.is_at_end();

        if is_single_line {
            // Single-line If: parse the inline statement(s)
            // We parse until we hit a newline or reach a colon (which could indicate Else on same line)
            while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
                // Check for inline Else (: Else or just Else on same line)
                if self.at_token(VB6Token::ElseKeyword) {
                    break;
                }
                
                // Try to parse using centralized statement dispatcher
                if self.is_statement_keyword() {
                    self.parse_statement();
                    continue;
                }

                // Handle other inline constructs
                match self.current_token() {
                    Some(VB6Token::Whitespace) | Some(VB6Token::EndOfLineComment) | Some(VB6Token::RemComment) => {
                        self.consume_token();
                    }
                    Some(VB6Token::ColonOperator) => {
                        // Colon can separate statements or precede Else
                        self.consume_token();
                    }
                    _ => {
                        // Check if this looks like an assignment
                        if self.is_at_assignment() {
                            self.parse_assignment_statement();
                        } else {
                            // Consume as unknown
                            self.consume_token();
                        }
                    }
                }
            }

            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        } else {
            // Multi-line If: consume newline after Then
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }

            // Parse body until "End If", "Else", or "ElseIf"
            self.parse_code_block(|parser| {
                (parser.at_token(VB6Token::EndKeyword)
                    && parser.peek_next_keyword() == Some(VB6Token::IfKeyword))
                    || parser.at_token(VB6Token::ElseIfKeyword)
                    || parser.at_token(VB6Token::ElseKeyword)
            });

            // Handle ElseIf and Else clauses
            while !self.is_at_end() {
                if self.at_token(VB6Token::ElseIfKeyword) {
                    // Parse ElseIf clause
                    self.parse_elseif_clause();
                } else if self.at_token(VB6Token::ElseKeyword) {
                    // Parse Else clause
                    self.parse_else_clause();
                } else {
                    break;
                }
            }

            // Consume "End If" and trailing tokens
            if self.at_token(VB6Token::EndKeyword) {
                // Consume "End"
                self.consume_token();

                // Consume any whitespace between "End" and "If"
                self.consume_whitespace();

                // Consume "If"
                self.consume_token();

                // Consume until newline (including it)
                self.consume_until(VB6Token::Newline);

                // Consume the newline
                if self.at_token(VB6Token::Newline) {
                    self.consume_token();
                }
            }
        }

        self.builder.finish_node(); // IfStatement
    }

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
    pub(super) fn parse_call_statement(&mut self) {

        // if we are now parsing a call statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::CallStatement.to_raw());
        
        // Consume "Call" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // CallStatement
    }

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
        let has_top_condition = self.at_token(VB6Token::WhileKeyword) || self.at_token(VB6Token::UntilKeyword);

        if has_top_condition {
            // Consume While or Until
            self.consume_token();

            // Parse condition - consume everything until newline
            self.parse_conditional();
        }

        // Consume newline after Do line
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Loop"
        self.parse_code_block(|parser| parser.at_token(VB6Token::LoopKeyword));

        // Consume "Loop" keyword
        if self.at_token(VB6Token::LoopKeyword) {
            self.consume_token();

            // Consume whitespace after Loop
            self.consume_whitespace();

            // Check if we have While or Until after Loop
            if self.at_token(VB6Token::WhileKeyword) || self.at_token(VB6Token::UntilKeyword) {
                // Consume While or Until
                self.consume_token();

                // Parse condition - consume everything until newline
                self.parse_conditional();
            }

            // Consume newline after Loop
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // DoStatement
    }

    /// Parse a For...Next statement.
    ///
    /// VB6 For...Next loop syntax:
    /// - For counter = start To end [Step step]...Next [counter]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fornext-statement)
    pub(super) fn parse_for_statement(&mut self) {

        // if we are now parsing a for statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ForStatement.to_raw());

        // Consume "For" keyword
        self.consume_token();

        // Consume everything until "To" or newline
        // This includes: counter variable, "=", start value
        while !self.is_at_end() 
            && !self.at_token(VB6Token::ToKeyword) 
            && !self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Consume "To" keyword if present
        if self.at_token(VB6Token::ToKeyword) {
            self.consume_token();

            // Consume everything until "Step" or newline (the end value)
            while !self.is_at_end() 
                && !self.at_token(VB6Token::StepKeyword) 
                && !self.at_token(VB6Token::Newline) {
                self.consume_token();
            }

            // Consume "Step" keyword if present
            if self.at_token(VB6Token::StepKeyword) {
                self.consume_token();

                // Consume everything until newline (the step value)
                self.consume_until(VB6Token::Newline);
            }
        }

        // Consume newline after For line
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Next"
        self.parse_code_block(|parser| parser.at_token(VB6Token::NextKeyword));

        // Consume "Next" keyword
        if self.at_token(VB6Token::NextKeyword) {
            self.consume_token();

            // Consume everything until newline (optional counter variable)
            self.consume_until(VB6Token::Newline);

            // Consume newline after Next
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // ForStatement
    }

    /// Parse a For Each...Next statement.
    ///
    /// VB6 For Each...Next loop syntax:
    /// - For Each element In collection...Next [element]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/for-eachnext-statement)
    pub(super) fn parse_for_each_statement(&mut self) {

        // if we are now parsing a for each statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ForEachStatement.to_raw());

        // Consume "For" keyword
        self.consume_token();

        // Consume whitespace
        self.consume_whitespace();

        // Consume "Each" keyword
        if self.at_token(VB6Token::EachKeyword) {
            self.consume_token();
        }

        // Consume everything until "In" or newline
        // This includes: element variable name and whitespace
        while !self.is_at_end() 
            && !self.at_token(VB6Token::InKeyword) 
            && !self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Consume "In" keyword if present
        if self.at_token(VB6Token::InKeyword) {
            self.consume_token();

            // Consume everything until newline (the collection)
            self.consume_until(VB6Token::Newline);
        }

        // Consume newline after For Each line
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Next"
        self.parse_code_block(|parser| parser.at_token(VB6Token::NextKeyword));

        // Consume "Next" keyword
        if self.at_token(VB6Token::NextKeyword) {
            self.consume_token();

            // Consume everything until newline (optional element variable)
            self.consume_until(VB6Token::Newline);

            // Consume newline after Next
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // ForEachStatement
    }

    /// Parse a Set statement.
    ///
    /// VB6 Set statement syntax:
    /// - Set objectVar = [New] objectExpression
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement)
    pub(super) fn parse_set_statement(&mut self) {

        // if we are now parsing a set statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::SetStatement.to_raw());

        // Consume "Set" keyword
        self.consume_token();

        // Consume everything until newline
        // This includes: variable, "=", [New], object expression
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // SetStatement
    }

    /// Parse a With statement.
    ///
    /// VB6 With statement syntax:
    /// - With object
    ///     .Property1 = value1
    ///     .Property2 = value2
    ///   End With
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/with-statement)
    pub(super) fn parse_with_statement(&mut self) {

        // if we are now parsing a with statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::WithStatement.to_raw());

        // Consume "With" keyword
        self.consume_token();

        // Consume everything until newline (the object expression)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the body until "End With"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::WithKeyword)
        });

        // Consume "End With" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "With"
            self.consume_whitespace();

            // Consume "With"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // WithStatement
    }

    /// Parse a Select Case statement.
    ///
    /// Syntax:
    ///   Select Case testexpression
    ///     Case expression1
    ///       statements1
    ///     Case expression2
    ///       statements2
    ///     Case Else
    ///       statementsElse
    ///   End Select
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/select-case-statement)
    pub(super) fn parse_select_case_statement(&mut self) {
        
        // if we are now parsing a select case statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::SelectCaseStatement.to_raw());

        // Consume "Select" keyword
        self.consume_token();

        // Consume any whitespace between "Select" and "Case"
        self.consume_whitespace();

        // Consume "Case" keyword
        if self.at_token(VB6Token::CaseKeyword) {
            self.consume_token();
        }

        // Consume everything until newline (the test expression)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse Case clauses until "End Select"
        while !self.is_at_end() {
            // Check for "End Select"
            if self.at_token(VB6Token::EndKeyword)
                && self.peek_next_keyword() == Some(VB6Token::SelectKeyword)
            {
                break;
            }

            // Check for "Case" keyword
            if self.at_token(VB6Token::CaseKeyword) {
                // Check if this is "Case Else"
                let is_case_else = self.peek_next_keyword() == Some(VB6Token::ElseKeyword);

                if is_case_else {
                    // Parse Case Else clause
                    self.builder.start_node(SyntaxKind::CaseElseClause.to_raw());

                    // Consume "Case"
                    self.consume_token();

                    // Consume any whitespace between "Case" and "Else"
                    self.consume_whitespace();

                    // Consume "Else"
                    if self.at_token(VB6Token::ElseKeyword) {
                        self.consume_token();
                    }

                    // Consume until newline
                    self.consume_until(VB6Token::Newline);
                    if self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }

                    // Parse statements in Case Else until next Case or End Select
                    self.parse_code_block(|parser| {
                        (parser.at_token(VB6Token::CaseKeyword))
                            || (parser.at_token(VB6Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(VB6Token::SelectKeyword))
                    });

                    self.builder.finish_node(); // CaseElseClause
                } else {
                    // Parse regular Case clause
                    self.builder.start_node(SyntaxKind::CaseClause.to_raw());

                    // Consume "Case"
                    self.consume_token();

                    // Consume the case expression(s) until newline
                    self.consume_until(VB6Token::Newline);
                    if self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }

                    // Parse statements in Case until next Case or End Select
                    self.parse_code_block(|parser| {
                        (parser.at_token(VB6Token::CaseKeyword))
                            || (parser.at_token(VB6Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(VB6Token::SelectKeyword))
                    });

                    self.builder.finish_node(); // CaseClause
                }
            } else {
                // Consume whitespace, newlines, and comments
                self.consume_token();
            }
        }

        // Consume "End Select" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Select"
            self.consume_whitespace();

            // Consume "Select"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // SelectCaseStatement
    }

    /// Parse a GoTo statement.
    ///
    /// Syntax:
    ///   GoTo label
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/goto-statement)
    pub(super) fn parse_goto_statement(&mut self) {
        
        // if we are now parsing a goto statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::GotoStatement.to_raw());

        // Consume "GoTo" keyword
        self.consume_token();

        // Consume everything until newline (the label name)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // GotoStatement
    }

    /// Parse a label statement.
    ///
    /// VB6 label syntax:
    /// - LabelName:
    ///
    /// Labels are used as targets for GoTo and GoSub statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/goto-statement)
    pub(super) fn parse_label_statement(&mut self) {

        // if we are now parsing a label statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::LabelStatement.to_raw());

        // Consume the label identifier
        self.consume_token();

        // Consume optional whitespace
        self.consume_whitespace();

        // Consume the colon
        if self.at_token(VB6Token::ColonOperator) {
            self.consume_token();
        }

        // Consume the newline if present
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // LabelStatement
    }

    /// Check if the current position is at a label.
    /// A label is an identifier or number followed by a colon.
    pub(super) fn is_at_label(&self) -> bool {

        let next_token_is_colon = matches!(self.peek_next_token(), Some(VB6Token::ColonOperator));

        if next_token_is_colon == false {
            return false;
        }

        // If we are not parsing the header, then some keywords are valid identifiers (like "Begin")
        // TODO: Consider adding a list of keywords that can be used as labels. 
        // TODO: Also consider modifying tokenizer to recognize when inside header to more easily identify Identifiers vs header only keywords.
        if !self.parsing_header && matches!(self.current_token(), Some(VB6Token::BeginKeyword)) {
            return true;
        }

        self.is_identifier() || self.is_number()
    }

    /// Parse an AppActivate statement.
    ///
    /// VB6 AppActivate statement syntax:
    /// - AppActivate title[, wait]
    ///
    /// Activates an application window.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/appactivate-statement)
    pub(super) fn parse_appactivate_statement(&mut self) {
        // if we are now parsing an AppActivate statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::AppActivateStatement.to_raw());

        // Consume "AppActivate" keyword
        self.consume_token();

        // Consume everything until newline (title and optional wait parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // AppActivateStatement
    }

    /// Parse a Beep statement.
    ///
    /// VB6 Beep statement syntax:
    /// - Beep
    ///
    /// Sounds a tone through the computer's speaker.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/beep-statement)
    pub(super) fn parse_beep_statement(&mut self) {
        // if we are now parsing a beep statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::BeepStatement.to_raw());

        // Consume "Beep" keyword
        self.consume_token();

        // Consume any whitespace and comments until newline
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // BeepStatement
    }

    /// Parse a ChDir statement.
    ///
    /// VB6 ChDir statement syntax:
    /// - ChDir path
    ///
    /// Changes the current directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdir-statement)
    pub(super) fn parse_chdir_statement(&mut self) {
        // if we are now parsing a ChDir statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ChDirStatement.to_raw());

        // Consume "ChDir" keyword
        self.consume_token();

        // Consume everything until newline (the path parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ChDirStatement
    }

    /// Parse a ChDrive statement.
    ///
    /// VB6 ChDrive statement syntax:
    /// - ChDrive drive
    ///
    /// Changes the current drive.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdrive-statement)
    pub(super) fn parse_chdrive_statement(&mut self) {
        // if we are now parsing a ChDrive statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ChDriveStatement.to_raw());

        // Consume "ChDrive" keyword
        self.consume_token();

        // Consume everything until newline (the drive parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ChDriveStatement
    }

    /// Parse a ReDim statement.
    ///
    /// VB6 ReDim statement syntax:
    /// - ReDim [Preserve] varname(subscripts) [As type] [, varname(subscripts) [As type]] ...
    ///
    /// Used at procedure level to reallocate storage space for dynamic array variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/redim-statement)
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

    /// Parse an Exit statement.
    ///
    /// Syntax:
    ///   Exit Do
    ///   Exit For
    ///   Exit Function
    ///   Exit Property
    ///   Exit Sub
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/exit-statement)
    pub(super) fn parse_exit_statement(&mut self) {
        
        // if we are now parsing an exit statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ExitStatement.to_raw());

        // Consume "Exit" keyword
        self.consume_token();

        // Consume whitespace after Exit
        self.consume_whitespace();

        // Consume the exit type (Do, For, Function, Property, Sub)
        if self.at_token(VB6Token::DoKeyword)
            || self.at_token(VB6Token::ForKeyword)
            || self.at_token(VB6Token::FunctionKeyword)
            || self.at_token(VB6Token::PropertyKeyword)
            || self.at_token(VB6Token::SubKeyword)
        {
            self.consume_token();
        }

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ExitStatement
    }

    /// Check if the current token is a statement keyword that parse_statement can handle.
    pub(super) fn is_statement_keyword(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::IfKeyword)
                | Some(VB6Token::CallKeyword)
                | Some(VB6Token::SetKeyword)
                | Some(VB6Token::WithKeyword)
                | Some(VB6Token::SelectKeyword)
                | Some(VB6Token::GotoKeyword)
                | Some(VB6Token::ExitKeyword)
                | Some(VB6Token::AppActivateKeyword)
                | Some(VB6Token::BeepKeyword)
                | Some(VB6Token::ChDirKeyword)
                | Some(VB6Token::ChDriveKeyword)
                | Some(VB6Token::ReDimKeyword)
                | Some(VB6Token::DoKeyword)
                | Some(VB6Token::ForKeyword)
        )
    }

    /// This is a centralized statement dispatcher that handles all VB6 statement types.
    pub(super) fn parse_statement(&mut self) {
        match self.current_token() {
            Some(VB6Token::IfKeyword) => {
                self.parse_if_statement();
            }
            Some(VB6Token::CallKeyword) => {
                self.parse_call_statement();
            }
            Some(VB6Token::SetKeyword) => {
                self.parse_set_statement();
            }
            Some(VB6Token::WithKeyword) => {
                self.parse_with_statement();
            }
            Some(VB6Token::SelectKeyword) => {
                self.parse_select_case_statement();
            }
            Some(VB6Token::GotoKeyword) => {
                self.parse_goto_statement();
            }
            Some(VB6Token::ExitKeyword) => {
                self.parse_exit_statement();
            }
            Some(VB6Token::AppActivateKeyword) => {
                self.parse_appactivate_statement();
            }
            Some(VB6Token::BeepKeyword) => {
                self.parse_beep_statement();
            }
            Some(VB6Token::ChDirKeyword) => {
                self.parse_chdir_statement();
            }
            Some(VB6Token::ChDriveKeyword) => {
                self.parse_chdrive_statement();
            }
            Some(VB6Token::ReDimKeyword) => {
                self.parse_redim_statement();
            }
            Some(VB6Token::DoKeyword) => {
                self.parse_do_statement();
            }
            Some(VB6Token::ForKeyword) => {
                // Peek ahead to see if next keyword is "Each"
                if let Some(VB6Token::EachKeyword) = self.peek_next_keyword() {
                    self.parse_for_each_statement();
                } else {
                    self.parse_for_statement();
                }
            }
            _ => {},
        }
    }
}
