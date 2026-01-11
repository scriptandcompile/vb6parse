//! Runtime control and state statements.
//!
//! This module contains parsers for VB6 statements that control runtime state:
//! - System time (Date, Time)
//! - Error handling (Error)
//! - Random number generation (Randomize)

pub(crate) mod date;
pub(crate) mod error;
pub(crate) mod randomize;
pub(crate) mod time;
