//! Built-in VB6 statement parsing.
//!
//! This module handles parsing of VB6 built-in system statements
//! each in their own sub-module.
//!
//! The built-in statements handled here are:
//! - AppActivate: Activate an application window
//! - Beep: Sound a tone through the computer's speaker
//! - ChDir: Change the current directory or folder
//! - ChDrive: Change the current drive
//! - Close: Close files opened with the Open statement
//! - Date: Set the current system date
//! - DeleteSetting: Delete a section or key setting from the Windows registry
//! - Error: Generate a run-time error
//! - FileCopy: Copy a file
//! - Get: Read data from an open disk file into a variable
//! - Input: Read data from an open sequential file
//! - Line Input: Read an entire line from a sequential file
//! - Kill: Delete a file from disk
//! - Load: Load a form or control into memory
//! - Lock: Control access to all or part of an open file
//! - Unlock: Remove access restrictions from an open file
//! - LSet: Left-align a string or copy a user-defined type
//! - Mid: Replace characters in a string variable
//! - MidB: Replace bytes in a string variable
//! - Reset: Close all disk files opened using the Open statement
//!

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

mod app_activate;
mod beep;
mod ch_dir;
mod ch_drive;
mod close;
mod date;
mod delete_setting;
mod error;
mod file_copy;
mod get;
mod input;
mod kill;
mod line_input;
mod load;
mod lock;
mod lset;
mod mid;
mod midb;
mod reset;
mod unlock;

impl<'a> Parser<'a> {
    /// Check if the current token is a library statement keyword.
    pub(super) fn is_library_statement_keyword(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::AppActivateKeyword)
                | Some(VB6Token::BeepKeyword)
                | Some(VB6Token::ChDirKeyword)
                | Some(VB6Token::ChDriveKeyword)
                | Some(VB6Token::CloseKeyword)
                | Some(VB6Token::DateKeyword)
                | Some(VB6Token::DeleteSettingKeyword)
                | Some(VB6Token::ErrorKeyword)
                | Some(VB6Token::FileCopyKeyword)
                | Some(VB6Token::GetKeyword)
                | Some(VB6Token::InputKeyword)
                | Some(VB6Token::KillKeyword)
                | Some(VB6Token::LineKeyword)
                | Some(VB6Token::LoadKeyword)
                | Some(VB6Token::LockKeyword)
                | Some(VB6Token::UnlockKeyword)
                | Some(VB6Token::LSetKeyword)
                | Some(VB6Token::MidKeyword)
                | Some(VB6Token::MidBKeyword)
                | Some(VB6Token::ResetKeyword)
        )
    }

    /// Dispatch library statement parsing to the appropriate parser.
    pub(super) fn parse_library_statement(&mut self) {
        match self.current_token() {
            Some(VB6Token::AppActivateKeyword) => {
                self.parse_app_activate_statement();
            }
            Some(VB6Token::BeepKeyword) => {
                self.parse_beep_statement();
            }
            Some(VB6Token::ChDirKeyword) => {
                self.parse_ch_dir_statement();
            }
            Some(VB6Token::ChDriveKeyword) => {
                self.parse_ch_drive_statement();
            }
            Some(VB6Token::CloseKeyword) => {
                self.parse_close_statement();
            }
            Some(VB6Token::DateKeyword) => {
                self.parse_date_statement();
            }
            Some(VB6Token::DeleteSettingKeyword) => {
                self.parse_delete_setting_statement();
            }
            Some(VB6Token::ErrorKeyword) => {
                self.parse_error_statement();
            }
            Some(VB6Token::FileCopyKeyword) => {
                self.parse_file_copy_statement();
            }
            Some(VB6Token::GetKeyword) => {
                self.parse_get_statement();
            }
            Some(VB6Token::InputKeyword) => {
                self.parse_input_statement();
            }
            Some(VB6Token::KillKeyword) => {
                self.parse_kill_statement();
            }
            Some(VB6Token::LineKeyword) => {
                self.parse_line_input_statement();
            }
            Some(VB6Token::LoadKeyword) => {
                self.parse_load_statement();
            }
            Some(VB6Token::LockKeyword) => {
                self.parse_lock_statement();
            }
            Some(VB6Token::UnlockKeyword) => {
                self.parse_unlock_statement();
            }
            Some(VB6Token::LSetKeyword) => {
                self.parse_lset_statement();
            }
            Some(VB6Token::MidKeyword) => {
                self.parse_mid_statement();
            }
            Some(VB6Token::MidBKeyword) => {
                self.parse_midb_statement();
            }
            Some(VB6Token::ResetKeyword) => {
                self.parse_reset_statement();
            }
            _ => {}
        }
    }

    /// Generic parser for built-in statements that follow the pattern:
    /// - Keyword [arguments]
    ///
    /// All built-in statements in this module share the same structure:
    /// 1. Set parsing_header to false
    /// 2. Start a syntax node of the given kind
    /// 3. Consume the keyword token
    /// 4. Consume everything until newline (arguments/parameters)
    /// 5. Consume the newline
    /// 6. Finish the syntax node
    pub(super) fn parse_simple_builtin_statement(&mut self, kind: SyntaxKind) {
        // if we are now parsing a built-in statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(kind.to_raw());

        // Consume the keyword
        self.consume_token();

        // Consume everything until newline (arguments/parameters)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node();
    }
}
