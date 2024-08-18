pub mod errors;
pub mod language;
pub mod parsers;

pub use crate::language::VB6Color;
pub use crate::language::VB6Control;
pub use crate::language::VB6ControlCommonInformation;
pub use crate::language::VB6ControlKind;
pub use crate::language::VB6Token;

// mod parsers::class;
// mod parsers::form::VB6FormFile;
// mod parsers::header::VB6Header;
// mod parsers::module::VB6ModuleFile;

pub use crate::parsers::vb6;
pub use crate::parsers::VB6Project;
