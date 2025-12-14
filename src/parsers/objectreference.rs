use serde::Serialize;
use uuid::Uuid;

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum ObjectReference {
    Compiled {
        uuid: Uuid,
        version: String,
        unknown1: String,
        file_name: String,
    },
    Project {
        path: String,
    },
}

impl Serialize for ObjectReference {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        match self {
            ObjectReference::Compiled {
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
            ObjectReference::Project { path } => {
                let mut state = serializer.serialize_struct("ProjectReference", 1)?;

                state.serialize_field("path", path)?;

                state.end()
            }
        }
    }
}
