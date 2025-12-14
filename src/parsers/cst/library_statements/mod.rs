//! Built-in VB6 statement parsing.
//!
//! This module handles parsing of VB6 built-in system statements
//! each in their own sub-module.
//!
//! The built-in statements handled here are:
//! - `AppActivate`: Activate an application window
//! - Beep: Sound a tone through the computer's speaker
//! - `ChDir`: Change the current directory or folder
//! - `ChDrive`: Change the current drive
//! - Close: Close files opened with the Open statement
//! - Date: Set the current system date
//! - `DeleteSetting`: Delete a section or key setting from the Windows registry
//! - Error: Generate a run-time error
//! - `FileCopy`: Copy a file
//! - Get: Read data from an open disk file into a variable
//! - Put: Write data from a variable to a disk file
//! - Input: Read data from an open sequential file
//! - Line Input: Read an entire line from a sequential file
//! - Kill: Delete a file from disk
//! - Load: Load a form or control into memory
//! - Unload: Remove a form or control from memory
//! - Lock: Control access to all or part of an open file
//! - Unlock: Remove access restrictions from an open file
//! - `LSet`: Left-align a string or copy a user-defined type
//! - `RSet`: Right-align a string within a string variable
//! - Mid: Replace characters in a string variable
//! - `MidB`: Replace bytes in a string variable
//! - `MkDir`: Create a new directory or folder
//! - `RmDir`: Remove an empty directory or folder
//! - Name: Rename a disk file, directory, or folder
//! - Open: Open a file for input/output operations
//! - Print: Write display-formatted data to a sequential file
//! - Reset: Close all disk files opened using the Open statement
//! - `SavePicture`: Save a graphical image to a file
//! - `SaveSetting`: Save or create an application entry in the Windows registry
//! - Seek: Set the position for the next read/write operation in a file
//! - `SendKeys`: Send keystrokes to the active window
//! - `SetAttr`: Set attribute information for a file
//! - Stop: Suspend execution
//! - Time: Set the current system time
//! - Randomize: Initialize the random number generator
//! - Width: Assign an output line width to a file
//! - Write: Write data to a sequential file
//!

use crate::language::Token;
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
mod mkdir;
mod name;
mod open;
mod print;
mod put;
mod randomize;
mod reset;
mod rmdir;
mod rset;
mod savepicture;
mod savesetting;
mod seek;
mod sendkeys;
mod setattr;
mod stop;
mod time;
mod unload;
mod unlock;
mod width;
mod write;

impl Parser<'_> {
    /// Check if the current token is a library statement keyword.
    ///
    /// Special handling:
    /// - `ErrorKeyword` followed by `DollarSign` is NOT a statement (it's the Error$ function)
    /// - `MidKeyword` followed by `DollarSign` is NOT a statement (it's the Mid$ function)
    ///    So we exclude those patterns.
    pub(super) fn is_library_statement_keyword(&self) -> bool {
        // Special case: keyword/identifier + DollarSign is a function, not a statement
        if self.at_keyword_dollar() {
            return false;
        }

        matches!(
            self.current_token(),
            Some(
                Token::AppActivateKeyword
                    | Token::BeepKeyword
                    | Token::ChDirKeyword
                    | Token::ChDriveKeyword
                    | Token::CloseKeyword
                    | Token::DateKeyword
                    | Token::DeleteSettingKeyword
                    | Token::ErrorKeyword
                    | Token::FileCopyKeyword
                    | Token::GetKeyword
                    | Token::PutKeyword
                    | Token::InputKeyword
                    | Token::KillKeyword
                    | Token::LineKeyword
                    | Token::LoadKeyword
                    | Token::UnloadKeyword
                    | Token::LockKeyword
                    | Token::UnlockKeyword
                    | Token::LSetKeyword
                    | Token::MidKeyword
                    | Token::MidBKeyword
                    | Token::MkDirKeyword
                    | Token::NameKeyword
                    | Token::OpenKeyword
                    | Token::PrintKeyword
                    | Token::RandomizeKeyword
                    | Token::ResetKeyword
                    | Token::RmDirKeyword
                    | Token::RSetKeyword
                    | Token::SavePictureKeyword
                    | Token::SaveSettingKeyword
                    | Token::SeekKeyword
                    | Token::SendKeysKeyword
                    | Token::SetAttrKeyword
                    | Token::StopKeyword
                    | Token::TimeKeyword
                    | Token::WidthKeyword
                    | Token::WriteKeyword
            )
        )
    }

    /// Dispatch library statement parsing to the appropriate parser.
    pub(super) fn parse_library_statement(&mut self) {
        match self.current_token() {
            Some(Token::AppActivateKeyword) => {
                self.parse_app_activate_statement();
            }
            Some(Token::BeepKeyword) => {
                self.parse_beep_statement();
            }
            Some(Token::ChDirKeyword) => {
                self.parse_ch_dir_statement();
            }
            Some(Token::ChDriveKeyword) => {
                self.parse_ch_drive_statement();
            }
            Some(Token::CloseKeyword) => {
                self.parse_close_statement();
            }
            Some(Token::DateKeyword) => {
                self.parse_date_statement();
            }
            Some(Token::DeleteSettingKeyword) => {
                self.parse_delete_setting_statement();
            }
            Some(Token::ErrorKeyword) => {
                self.parse_error_statement();
            }
            Some(Token::FileCopyKeyword) => {
                self.parse_file_copy_statement();
            }
            Some(Token::GetKeyword) => {
                self.parse_get_statement();
            }
            Some(Token::PutKeyword) => {
                self.parse_put_statement();
            }
            Some(Token::InputKeyword) => {
                self.parse_input_statement();
            }
            Some(Token::KillKeyword) => {
                self.parse_kill_statement();
            }
            Some(Token::LineKeyword) => {
                self.parse_line_input_statement();
            }
            Some(Token::LoadKeyword) => {
                self.parse_load_statement();
            }
            Some(Token::UnloadKeyword) => {
                self.parse_unload_statement();
            }
            Some(Token::LockKeyword) => {
                self.parse_lock_statement();
            }
            Some(Token::UnlockKeyword) => {
                self.parse_unlock_statement();
            }
            Some(Token::LSetKeyword) => {
                self.parse_lset_statement();
            }
            Some(Token::MidKeyword) => {
                self.parse_mid_statement();
            }
            Some(Token::MidBKeyword) => {
                self.parse_midb_statement();
            }
            Some(Token::MkDirKeyword) => {
                self.parse_mkdir_statement();
            }
            Some(Token::NameKeyword) => {
                self.parse_name_statement();
            }
            Some(Token::OpenKeyword) => {
                self.parse_open_statement();
            }
            Some(Token::PrintKeyword) => {
                self.parse_print_statement();
            }
            Some(Token::RandomizeKeyword) => {
                self.parse_randomize_statement();
            }
            Some(Token::ResetKeyword) => {
                self.parse_reset_statement();
            }
            Some(Token::RmDirKeyword) => {
                self.parse_rmdir_statement();
            }
            Some(Token::RSetKeyword) => {
                self.parse_rset_statement();
            }
            Some(Token::SavePictureKeyword) => {
                self.parse_savepicture_statement();
            }
            Some(Token::SaveSettingKeyword) => {
                self.parse_savesetting_statement();
            }
            Some(Token::SeekKeyword) => {
                self.parse_seek_statement();
            }
            Some(Token::SendKeysKeyword) => {
                self.parse_sendkeys_statement();
            }
            Some(Token::SetAttrKeyword) => {
                self.parse_setattr_statement();
            }
            Some(Token::StopKeyword) => {
                self.parse_stop_statement();
            }
            Some(Token::TimeKeyword) => {
                self.parse_time_statement();
            }
            Some(Token::WidthKeyword) => {
                self.parse_width_statement();
            }
            Some(Token::WriteKeyword) => {
                self.parse_write_statement();
            }
            _ => {}
        }
    }

    /// Generic parser for built-in statements that follow the pattern:
    /// - Keyword [arguments]
    ///
    /// All built-in statements in this module share the same structure:
    /// 1. Set `parsing_header` to false
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
        self.consume_until_after(Token::Newline);

        self.builder.finish_node();
    }
}
