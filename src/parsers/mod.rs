mod header;
mod vb6stream;

pub mod class;
pub mod form;
pub mod module;
pub mod project;
pub mod vb6;

use bstr::BStr;
use serde::Serialize;
use uuid::Uuid;

pub use class::{
    FileUsage, MtsStatus, Persistance, VB6ClassAttributes, VB6ClassFile, VB6ClassHeader,
    VB6ClassProperties, VB6ClassVersion,
};

pub use form::VB6FormFile;
pub use module::VB6ModuleFile;

pub use project::{
    CompileTargetType, VB6Project, VB6ProjectClass, VB6ProjectModule, VB6ProjectReference,
};

pub use vb6::vb6_parse;

pub use vb6stream::VB6Stream;

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ObjectReference<'a> {
    pub uuid: Uuid,
    pub version: &'a BStr,
    pub unknown1: &'a BStr,
    pub file_name: &'a BStr,
}

impl Serialize for VB6ObjectReference<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("VB6Object", 4)?;

        state.serialize_field("uuid", &self.uuid.to_string())?;
        state.serialize_field("version", &self.version)?;
        state.serialize_field("unknown1", &self.unknown1)?;
        state.serialize_field("file_name", &self.file_name)?;

        state.end()
    }
}
