//! This module contains the syntax parsers for the Visual Basic 6 language.
//!
//! The major component of this module is the Concrete Syntax Tree (CST) parser
//! which parses VB6 code into a syntax tree representation.
//!
//! For file format parsers (Project, Class, Form, Module, Resource files),
//! see the [`crate::files`] module.
//!

pub mod cst;
pub mod parseresults;
pub mod syntaxkind;

pub use crate::io::{Comparator, SourceFile, SourceStream};
pub use crate::parsers::cst::{parse, ConcreteSyntaxTree, CstNode, SerializableTree};
pub use crate::parsers::syntaxkind::SyntaxKind;
pub use parseresults::{Diagnostics, ParseResult};
