pub mod class;
pub mod errors;
pub mod form;
pub mod module;
pub mod project;
pub mod vb6;
pub mod vb6stream;

/// Represents a VB6 file format version.
/// A VB6 file format version contains a major version number and a minor version number.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6FileFormatVersion {
    pub major: u8,
    pub minor: u8,
}
