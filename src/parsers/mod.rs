//! This module contains the file parsers for the Visual Basic 6 language.
//!
//! The major components of this module are:
//!
//! `VB6Stream` - The main parsing structure for VB6 files. It holds the
//! bytes for a file and provides basic method for iterating over the bytes as
//! well as tracking the line, column, and character position of the current byte.
//!
//! `VB6Project` - a structure for holding the parsed information from a VB6
//! project file.
//!
//! `VB6ClassFile` - a structure for holding the parsed information from a VB6
//! class file.
//!
//! `VB6FormFile` - a structure for holding the parsed information from a VB6
//! form file.
//!
//! `VB6ModuleFile` - a structure for holding the parsed information from a VB6
//! module file.
//!

mod header;

pub mod class;
pub mod cst;
pub mod form;
pub mod module;
pub mod objectreference;
pub mod parseresults;
pub mod project;
pub mod properties;
pub mod resource;
pub mod syntaxkind;

pub use class::*;
pub use form::FormFile;
pub use header::{FileAttributes, FileFormatVersion};
pub use module::ModuleFile;
pub use objectreference::ObjectReference;
pub use properties::Properties;
pub use resource::{list_resolver, resource_file_resolver};

pub use project::{
    compilesettings::*, properties::*, Project, ProjectClassReference, ProjectModuleReference,
    ProjectReference,
};

pub use crate::parsers::cst::{parse, ConcreteSyntaxTree, CstNode, SerializableTree};
pub use crate::parsers::syntaxkind::SyntaxKind;
pub use crate::sourcestream::*;
pub use crate::SourceFile;
pub use parseresults::ParseResult;
pub use uuid::Uuid;
