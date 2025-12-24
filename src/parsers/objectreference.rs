//! Defines the `ObjectReference` enum representing references to compiled objects or project files.
//!

use serde::Serialize;
use uuid::Uuid;

/// Represents a reference to either a compiled object or a project file.
/// # Examples
/// ```rust
/// # use uuid::Error;
///
/// # fn main() -> Result<(), uuid::Error> {
///     use vb6parse::parsers::objectreference::ObjectReference;
///     use uuid::Uuid;
///
///     let compiled_ref = ObjectReference::Compiled {
///         uuid: Uuid::parse_str("123e4567-e89b-12d3-a456-426614174000")?,
///         version: "1.0".to_string(),
///         unknown1: "SomeValue".to_string(),
///         file_name: "MyLibrary.dll".to_string(),
///     };
///     let project_ref = ObjectReference::Project {
///         path: "MyProject.vbp".to_string(),
///     };
///     # Ok(())
/// # }
/// ```
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum ObjectReference {
    /// A reference to a compiled object.
    Compiled {
        /// The UUID of the compiled object.
        uuid: Uuid,
        /// The version of the compiled object.
        version: String,
        /// An unknown string field.
        unknown1: String,
        /// The file name of the compiled object.
        file_name: String,
    },
    /// A reference to a project file.
    Project {
        /// The path to the project file.
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
