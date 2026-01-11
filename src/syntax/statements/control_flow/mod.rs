//! Control flow statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 control flow statements:
//! - Jump statements (`GoTo`, `GoSub`, `Return`, `Label`) - see [`jump`] module
//! - Exit and Resume statements - see [`exit_resume`] module
//! - On-prefixed statements (`On Error`, `On GoTo`, `On GoSub`) - see [`on_statements`] module
//!
//! Note: `If`/`Then`/`Else`/`ElseIf` statements are in the `if_statements` module.
//! Note: `Select Case` statements are in the `select_statements` module.
//! Note: `For`/`Next` and `For Each`/`Next` statements are in the `for_statements` module.
//! Note: `Do`/`Loop` statements are in the `loop_statements` module.

pub mod exit_resume;
pub mod jump;
pub mod on_statements;
