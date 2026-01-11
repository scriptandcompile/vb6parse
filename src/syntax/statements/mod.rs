//! VB6 statement parsing.
//!
//! This module contains parsers for all VB6 statements, organized by category:
//! - Control flow (`If`, `Select`, `For`, `Do`, `GoTo`, etc.)
//! - Declarations (`Dim`, `Const`, `ReDim`, `Type`, `Enum`, etc.)
//! - Object operations (`Set`, `With`, `New`)

pub mod control_flow;
pub mod declarations;
pub mod objects;
