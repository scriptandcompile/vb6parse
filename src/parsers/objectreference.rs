use bstr::BStr;
use serde::Serialize;
use uuid::Uuid;

#[derive(Debug, PartialEq, Eq, Copy, Clone)]
pub enum VB6ObjectReference<'a> {
    Compiled {
        uuid: Uuid,
        version: &'a BStr,
        unknown1: &'a BStr,
        file_name: &'a BStr,
    },
    Project {
        path: &'a BStr,
    },
}

impl Serialize for VB6ObjectReference<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        match self {
            VB6ObjectReference::Compiled {
                uuid,
                version,
                unknown1,
                file_name,
            } => {
                let mut state = serializer.serialize_struct("CompiledReference", 4)?;

                state.serialize_field("uuid", &uuid.to_string())?;
                state.serialize_field("version", version)?;
                state.serialize_field("unknown1", unknown1)?;
                state.serialize_field("file_name", file_name)?;

                state.end()
            }
            VB6ObjectReference::Project { path } => {
                let mut state = serializer.serialize_struct("ProjectReference", 1)?;

                state.serialize_field("path", path)?;

                state.end()
            }
        }
    }
}
