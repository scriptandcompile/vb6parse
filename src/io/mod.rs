//! Input/Output layer for `VB6Parse`
//!
//! This module handles the lowest level of parsing infrastructure:
//! - File reading and Windows-1252 encoding/decoding
//! - Character stream management with position tracking
//! - Line/column offset tracking
//!
//! # Key Components
//!
//! - [`SourceFile`] - Handles Windows-1252 decoding and file content management
//! - [`SourceStream`] - Low-level character stream with offset tracking
//!
//! # Example
//!
//! ```no_run
//! use vb6parse::io::SourceFile;
//!
//! let bytes = std::fs::read("MyProject.vbp").expect("Failed to read project file");
//! let source = SourceFile::decode_with_replacement("MyProject.vbp", &bytes)
//!     .expect("Failed to decode source file");
//! let stream = source.source_stream();
//! // Use stream for parsing...
//! ```

pub mod source_file;
pub mod source_stream;

pub use source_file::SourceFile;
pub use source_stream::{Comparator, SourceStream};
