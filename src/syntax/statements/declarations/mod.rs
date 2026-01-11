//! Declaration statement parsing for VB6 CST.
//!
//! This module provides parsers for VB6 declaration statements including:
//! - Variable declarations (Dim, Private, Public, Const, Static)
//! - Array operations (ReDim with Preserve support)
//! - Array cleanup (Erase)
//!
//! The implementation is organized into focused submodules:
//! - `arrays` - ReDim statement parsing for dynamic array reallocation
//! - `variables` - Variable and constant declarations with WithEvents support
//! - `erase` - Erase statement parsing for array cleanup

pub(crate) mod arrays;
pub(crate) mod erase;
pub(crate) mod variables;
