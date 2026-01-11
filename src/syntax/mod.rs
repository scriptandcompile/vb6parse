//! VB6 Syntax Layer - Concrete Syntax Tree (CST) Parsing.
//!
//! This module provides the syntax parsing layer that constructs a Concrete Syntax Tree (CST)
//! from a token stream. The CST preserves all source information including whitespace and comments.
//!
//! The syntax layer is organized into:
//! - **statements** - Statement parsing (control flow, declarations, object operations)
//! - **expressions** - Expression parsing (literals, operators, function calls, etc.)
//! - **library** - Built-in VB6 functions and statements
//!
//! For the CST data structure itself, see the `parsers::cst` module.

pub mod expressions;
pub mod library;
pub mod statements;

// Re-export commonly used types from the CST module for convenience
pub use crate::parsers::{parse, ConcreteSyntaxTree, CstNode, SyntaxKind};
