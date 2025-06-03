//! This module contains the file parsers for the Visual Basic 6 language.
//!
//! The major components of this module are:
//!
//! `VB6Stream` - The main parsing structure for VB6 files. It holds the
//! bytes for a file and provides basic methoda for iterating over the bytes as
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
mod vb6stream;

pub mod class;
pub mod compilesettings;
pub mod form;
pub mod module;
pub mod objectreference;
pub mod parserresults;
pub mod project;
pub mod properties;
pub mod sourcestream;
pub mod vb6;

pub use class::{
    FileUsage, MtsStatus, Persistance, VB6ClassFile, VB6ClassHeader, VB6ClassProperties,
    VB6ClassVersion,
};

pub use compilesettings::{
    BoundsCheck, CompilationType, FloatingPointErrorCheck, OverflowCheck, PentiumFDivBugCheck,
    UnroundedFloatingPoint,
};
pub use form::{resource_file_resolver, VB6FormFile};
pub use header::{VB6FileAttributes, VB6FileFormatVersion};
pub use module::VB6ModuleFile;

pub use properties::Properties;

pub use objectreference::VB6ObjectReference;

pub use project::{
    CompileTargetType, VB6Project, VB6ProjectClass, VB6ProjectModule, VB6ProjectReference,
};

pub use vb6::{is_english_code, vb6_parse};

pub use vb6stream::VB6Stream;
