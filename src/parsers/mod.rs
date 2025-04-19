mod header;
mod vb6stream;

pub mod class;
pub mod compilesettings;
pub mod form;
pub mod module;
pub mod objectreference;
pub mod project;
pub mod properties;
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
pub use header::VB6FileAttributes;
pub use module::VB6ModuleFile;

pub use properties::Properties;

pub use objectreference::VB6ObjectReference;

pub use project::{
    CompileTargetType, VB6Project, VB6ProjectClass, VB6ProjectModule, VB6ProjectReference,
};

pub use vb6::{is_english_code, vb6_parse};

pub use vb6stream::VB6Stream;
