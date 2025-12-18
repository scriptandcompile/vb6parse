use std::collections::HashMap;

use serde::Serialize;

/// Represents a VB6 file format version.
/// A VB6 file format version contains a major version number and a minor version number.
///
/// The file format version is locked to the version of the language and IDE but
/// the file format version is not the same as the language version. Some files
/// may be at version 4.0 while others might be at version 5.0. It really depends
/// on how things changed as the language and IDE evolved between major versions
/// of the language.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub struct FileFormatVersion {
    /// The files major version number.
    pub major: u8,
    /// The files minor version number.
    pub minor: u8,
}

/// Represents if a class is in the global or local name space.
///
/// The global name space is the default name space for a class.
/// In the file, `VB_GlobalNameSpace` of 'False' means the class is in the local name space.
/// `VB_GlobalNameSpace` of 'True' means the class is in the global name space.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum NameSpace {
    /// The class is in the global name space.
    Global,
    /// The class is in the local name space.
    #[default]
    Local,
}

/// The creatable attribute is used to determine if the class can be created.
///
/// If True, the class can be created from anywhere. The class is essentially public.
/// If False, the class can only be created from within the class itself.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum Creatable {
    /// The class cannot be created from outside the class itself.
    False,
    /// The class can be created from anywhere.
    #[default]
    True,
}

/// Used to determine if the class has a pre-declared ID.
///
/// If True, the class has a pre-declared ID and can be accessed by
/// the class name without creating an instance of the class.
///
/// If False, the class does not have a pre-declared ID and must be
/// accessed by creating an instance of the class.
///
/// If True and the `VB_GlobalNameSpace` is True, the class shares namespace
/// access semantics with the VB6 intrinsic classes.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum PreDeclaredID {
    /// The class does not have a pre-declared ID.
    #[default]
    False,
    /// The class has a pre-declared ID.
    True,
}

/// Used to determine if the class is exposed.
///
/// The `VB_Exposed` attribute is not normally visible in the code editor region.
///
/// ----------------------------------------------------------------------------
///
/// True is public and False is internal.
/// Used in combination with the Creatable attribute to create a matrix of
/// scoping behavior.
///
/// ----------------------------------------------------------------------------
///
/// Private (Default).
///
/// `VB_Exposed` = False and `VB_Creatable` = False.
/// The class is accessible only within the enclosing project.
///
/// Instances of the class can only be created by modules contained within the
/// project that defines the class.
///
/// ----------------------------------------------------------------------------
///
/// Public Not Creatable.
///
/// `VB_Exposed` = True and `VB_Creatable` = False.
/// The class is accessible within the enclosing project and within projects
/// that reference the enclosing project.
///
/// Instances of the class can only be created by modules within the enclosing
/// project. Modules in other projects can reference the class name as a
/// declared type but can’t instantiate the class using new or the
/// `CreateObject` function.
///
/// ----------------------------------------------------------------------------
///
/// Public Creatable.
///
/// `VB_Exposed` = True and `VB_Creatable` = True.
/// The class is accessible within the enclosing project and within the
/// enclosing project and within projects that reference the enclosing project.
///
/// Any module that can access the class can create instances of it.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum Exposed {
    /// The class is not exposed.
    #[default]
    False,
    /// The class is exposed.
    True,
}

/// Represents the attributes of a VB6 file.
/// The attributes contain the name, global name space, creatable, pre-declared id, and exposed status.
///
/// None of these values are normally visible in the code editor region.
/// They are only visible in the file property explorer.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct FileAttributes {
    /// The name of the file.
    pub name: String, // Attribute VB_Name = "Organism"
    /// The status of the global name space of the file.
    pub global_name_space: NameSpace, // (True/False) Attribute VB_GlobalNameSpace = False
    /// The creatable status of the file.
    pub creatable: Creatable, // (True/False) Attribute VB_Creatable = True
    /// The pre-declared ID status of the file.
    pub predeclared_id: PreDeclaredID, // (True/False) Attribute VB_PredeclaredId = False
    /// The exposed status of the file.
    pub exposed: Exposed, // (True/False) Attribute VB_Exposed = False
    /// The description of the file.
    pub description: Option<String>, // Attribute VB_Description = "Description"
    /// Additional attributes of the file.
    pub ext_key: HashMap<String, String>, // Additional attributes
}

impl Default for FileAttributes {
    /// Creates a default instance of FileAttributes with default values.
    ///
    /// The default values are:
    /// - name: empty string
    /// - global_name_space: Local
    /// - creatable: True
    /// - pre_declared_id: False
    /// - exposed: False
    /// - description: None
    /// - ext_key: empty HashMap
    ///
    /// Returns: A FileAttributes instance with default values.
    fn default() -> Self {
        FileAttributes {
            name: String::new(),
            global_name_space: NameSpace::Local,
            creatable: Creatable::True,
            predeclared_id: PreDeclaredID::False,
            exposed: Exposed::False,
            description: None,
            ext_key: HashMap::new(),
        }
    }
}

/// Extracts the file format version from a CST.
///
/// Searches for a VersionStatement node in the CST and parses the version number.
/// Returns None if no version statement is found or if parsing fails.
///
/// # Arguments
/// * `cst` - A reference to the ConcreteSyntaxTree to extract the version from.
///
/// # Returns
/// An Option containing the FileFormatVersion if found, or None if not found or parsing fails.
pub(crate) fn extract_version(
    cst: &crate::parsers::ConcreteSyntaxTree,
) -> Option<FileFormatVersion> {
    use crate::parsers::SyntaxKind;

    // Find the VersionStatement node
    let version_nodes = cst.find_children_by_kind(SyntaxKind::VersionStatement);

    if version_nodes.is_empty() {
        return None;
    }

    let version_node = &version_nodes[0];

    // Traverse children to find the SingleLiteral token
    // Expected order: VersionKeyword, Whitespace, SingleLiteral, ...
    let mut found_version_keyword = false;

    for child in &version_node.children {
        // Skip until we find the VERSION keyword
        if !found_version_keyword {
            if child.kind == SyntaxKind::VersionKeyword {
                found_version_keyword = true;
            }
            continue;
        }

        // After VERSION keyword, look for SingleLiteral (the version number)
        if child.kind == SyntaxKind::SingleLiteral {
            // Parse the version number from the SingleLiteral text
            // It should be in format "major.minor" (e.g., "5.00" or "1.0")
            let version_str = child.text.trim();
            let parts: Vec<&str> = version_str.split('.').collect();

            if parts.len() != 2 {
                return None;
            }

            let major = parts[0].parse::<u8>().ok()?;
            let minor = parts[1].parse::<u8>().ok()?;

            return Some(FileFormatVersion { major, minor });
        }
    }

    None
}
