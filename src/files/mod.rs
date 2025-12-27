//! VB6 File Format Parsers (Object Layer)
//!
//! This module provides high-level parsers for VB6 file formats:
//! - `.vbp` - Project files ([`ProjectFile`])
//! - `.cls` - Class modules ([`ClassFile`])
//! - `.bas` - Code modules ([`ModuleFile`])
//! - `.frm` - Forms ([`FormFile`])
//! - `.frx` - Form resources ([`FormResourceFile`])
//!
//! These parsers operate on the "object layer" - they parse complete files into
//! structured objects. They build on top of the lower-level syntax (CST) parsers.
//!
//! # Architecture
//!
//! The files module is distinct from the syntax/CST layer:
//! - **Syntax layer** ([`crate::parsers::cst`]): Parses VB6 code into syntax trees
//! - **Files layer** (this module): Parses complete file formats, including headers,
//!   metadata, and code content
//!
//! # Common Utilities
//!
//! The [`common`] module provides shared utilities for parsing file headers,
//! properties, and attributes that are common across multiple file types.
//!
//! # Example
//!
//! ```rust
//! use vb6parse::io::SourceFile;
//! use vb6parse::files::ProjectFile;
//!
//! let input = r#"Type=Exe
//! Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#...\stdole2.tlb#OLE Automation
//! Form=Form1.frm
//! Module=Module1; Module1.bas
//! "#;
//!
//! let source = SourceFile::from_string("Project1.vbp", input);
//! let result = ProjectFile::parse(&source);
//!
//! let (project, failures) = result.unpack();
//! assert!(project.is_some());
//! ```

pub mod class;
pub mod common;
pub mod form;
pub mod module;
pub mod project;
pub mod resource;

// Re-export main file types for convenience
pub use class::ClassFile;
pub use form::FormFile;
pub use module::ModuleFile;
pub use project::ProjectFile;
pub use resource::FormResourceFile;
