//! Built-in VB6 statement parsing.
//!
//! This module handles parsing of VB6 built-in system statements organized by category.
//!
//! # Module Organization
//!
//! Library statements are organized into focused subdirectories:
//!
//! ## File Operations ([`file_operations`])
//! Statements for file I/O and manipulation:
//! - **Binary I/O**: `Get`, `Put`
//! - **Sequential I/O**: `Input`, `Line Input`, `Print`, `Write`
//! - **File Management**: `Open`, `Close`, `Reset`
//! - **File Control**: `Lock`, `Unlock`, `Seek`
//! - **File Manipulation**: `FileCopy`, `Kill`, `Name`
//! - **Formatting**: `Width`
//!
//! ## Filesystem ([`filesystem`])
//! Statements for directory and filesystem operations:
//! - **Navigation**: `ChDir`, `ChDrive`
//! - **Management**: `MkDir`, `RmDir`
//! - **Attributes**: `SetAttr`
//!
//! ## System Interaction ([`system_interaction`])
//! Statements for system and user interaction:
//! - **Application Control**: `AppActivate`, `Stop`
//! - **User Feedback**: `Beep`
//! - **UI Management**: `Load`, `Unload`
//! - **Registry**: `DeleteSetting`, `SaveSetting`
//! - **Graphics**: `SavePicture`
//! - **Input Simulation**: `SendKeys`
//!
//! ## String Manipulation ([`string_manipulation`])
//! Statements for string operations:
//! - **Alignment**: `LSet`, `RSet`
//! - **Replacement**: `Mid`, `MidB`
//!
//! ## Runtime State ([`runtime_state`])
//! Statements for runtime state management:
//! - **System Time**: `Date`, `Time`
//! - **Error Handling**: `Error`
//! - **Random Numbers**: `Randomize`

use crate::language::Token;
use crate::parsers::cst::Parser;
use crate::parsers::SyntaxKind;

pub(crate) mod file_operations;
pub(crate) mod filesystem;
pub(crate) mod runtime_state;
pub(crate) mod string_manipulation;
pub(crate) mod system_interaction;

impl Parser<'_> {
    /// Check if the current token is a library statement keyword.
    ///
    /// Special handling:
    /// - `ErrorKeyword` followed by `DollarSign` is NOT a statement (it's the Error$ function)
    /// - `MidKeyword` followed by `DollarSign` is NOT a statement (it's the Mid$ function) so we exclude those patterns.
    ///
    /// Checks both current position and next non-whitespace token.
    pub(crate) fn is_library_statement_keyword(&self) -> bool {
        // Special case: keyword/identifier + DollarSign is a function, not a statement
        if self.at_keyword_dollar() {
            return false;
        }

        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        matches!(
            token,
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
    pub(crate) fn parse_library_statement(&mut self) {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        match token {
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

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume the keyword
        self.consume_token();

        // Consume everything until newline (arguments/parameters)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node();
    }
}
