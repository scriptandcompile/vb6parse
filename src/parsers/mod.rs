//! This module contains the file parsers for the Visual Basic 6 language.
//!
//! The major components of this module are:
//!
//! `ProjectFile` - a structure for holding the parsed information from a VB6
//! project file.
//!
//! `ClassFile` - a structure for holding the parsed information from a VB6
//! class file.
//!
//! `FormFile` - a structure for holding the parsed information from a VB6
//! form file.
//!
//! `ModuleFile` - a structure for holding the parsed information from a VB6
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
pub use resource::FormResourceFile;

pub use project::{
    compilesettings::*, properties::*, ProjectClassReference, ProjectFile, ProjectModuleReference,
    ProjectReference,
};

pub use crate::io::{Comparator, SourceFile, SourceStream};
pub use crate::parsers::cst::{parse, ConcreteSyntaxTree, CstNode, SerializableTree};
pub use crate::parsers::syntaxkind::SyntaxKind;
pub use parseresults::ParseResult;
pub use uuid::Uuid;
