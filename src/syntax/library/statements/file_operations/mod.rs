//! File I/O operation statements.
//!
//! This module contains parsers for VB6 statements that perform file input/output operations:
//! - Binary file operations (Get, Put)
//! - Sequential file operations (Input, Line Input, Print, Write)
//! - File management (Open, Close, Reset)
//! - File positioning (Seek)
//! - File access control (Lock, Unlock)
//! - File manipulation (FileCopy, Kill, Name)
//! - Output formatting (Width)

pub(crate) mod close;
pub(crate) mod file_copy;
pub(crate) mod get;
pub(crate) mod input;
pub(crate) mod kill;
pub(crate) mod line_input;
pub(crate) mod lock;
pub(crate) mod name;
pub(crate) mod open;
pub(crate) mod print;
pub(crate) mod put;
pub(crate) mod reset;
pub(crate) mod seek;
pub(crate) mod unlock;
pub(crate) mod width;
pub(crate) mod write;
