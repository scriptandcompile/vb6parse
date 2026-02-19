//! Defines the `ProjectFile` struct and related parsing functions for VB6 Project files.
//!
//! Handles extraction of project type, references, objects, modules, classes, forms,
//! user controls, user documents, properties, and other related information from the Project file.
//!
pub mod compilesettings;
pub mod properties;

use std::collections::HashMap;
use std::convert::TryFrom;
use std::fmt::{Debug, Display};
use std::str::FromStr;

use serde::Serialize;
use strum::{EnumMessage, IntoEnumIterator};
use uuid::Uuid;

use crate::{
    errors::{DiagnosticLabel, ParserContext, ProjectError},
    files::common::ObjectReference,
    files::project::{
        compilesettings::CompilationType,
        properties::{CompileTargetType, ProjectProperties},
    },
    io::{Comparator, SourceFile, SourceStream},
    parsers::ParseResult,
};

/// Represents a VB6 Project file.
///
/// Contains information about the project's type, references, objects, modules, classes, forms,
/// user controls, user documents, properties, and other related information.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub struct ProjectFile<'a> {
    /// The type of the project (e.g., Exe, Dll, etc.).
    pub project_type: CompileTargetType,
    /// The list of references in the project.
    references: Vec<ProjectReference<'a>>,
    /// The list of object references in the project.
    objects: Vec<ObjectReference>,
    /// The list of module references in the project.
    modules: Vec<ProjectModuleReference<'a>>,
    /// The list of class references in the project.
    classes: Vec<ProjectClassReference<'a>>,
    /// The list of related documents in the project.
    related_documents: Vec<&'a str>,
    /// The list of property pages in the project.
    property_pages: Vec<&'a str>,
    /// The list of designers in the project.
    designers: Vec<&'a str>,
    /// The list of forms in the project.
    forms: Vec<&'a str>,
    /// The list of user controls in the project.
    user_controls: Vec<&'a str>,
    /// The list of user documents in the project.
    user_documents: Vec<&'a str>,
    /// Other properties grouped by section headers.
    pub other_properties: HashMap<&'a str, HashMap<&'a str, &'a str>>,
    /// The project properties.
    pub properties: ProjectProperties<'a>,
}

impl Display for ProjectFile<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "VB6 Project File: Type={:?}, References={}, Objects={}, Modules={}, Classes={}, Forms={}, UserControls={}, UserDocuments={}, RelatedDocuments={}, PropertyPages={}", 
            self.project_type,
            self.references.len(),
            self.objects.len(),
            self.modules.len(),
            self.classes.len(),
            self.forms.len(),
            self.user_controls.len(),
            self.user_documents.len(),
            self.related_documents.len(),
            self.property_pages.len()
        )
    }
}

/// Represents a reference to either a compiled object or a sub-project.
#[derive(Debug, PartialEq, Eq, Copy, Clone, Hash)]
pub enum ProjectReference<'a> {
    /// A reference to a compiled object.
    Compiled {
        /// The UUID of the compiled object.
        uuid: Uuid,
        /// An unknown string field.
        unknown1: &'a str,
        /// Another unknown string field.
        unknown2: &'a str,
        /// The path to the compiled object.
        path: &'a str,
        /// The description of the compiled object.
        description: &'a str,
    },
    /// A reference to a sub-project.
    SubProject {
        /// The path to the sub-project file.
        path: &'a str,
    },
}

impl Display for ProjectReference<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ProjectReference::Compiled {
                uuid,
                unknown1,
                unknown2,
                path,
                description,
            } => write!(
                f,
                "Compiled Reference: UUID={uuid}, Unknown1='{unknown1}', Unknown2='{unknown2}', Path='{path}', Description='{description}'"
            ),
            ProjectReference::SubProject { path } => {
                write!(f, "Sub-Project Reference: Path='{path}'")
            }
        }
    }
}

impl Serialize for ProjectReference<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        match self {
            ProjectReference::Compiled {
                uuid,
                unknown1,
                unknown2,
                path,
                description,
            } => {
                let mut state = serializer.serialize_struct("CompiledReference", 5)?;

                state.serialize_field("uuid", &uuid.to_string())?;
                state.serialize_field("unknown1", unknown1)?;
                state.serialize_field("unknown2", unknown2)?;
                state.serialize_field("path", path)?;
                state.serialize_field("description", description)?;

                state.end()
            }
            ProjectReference::SubProject { path } => {
                let mut state = serializer.serialize_struct("SubProjectReference", 1)?;

                state.serialize_field("path", path)?;

                state.end()
            }
        }
    }
}

/// Represents a reference to a module in a VB6 project.
#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, Hash)]
pub struct ProjectModuleReference<'a> {
    /// The name of the module.
    pub name: &'a str,
    /// The path to the module file.
    pub path: &'a str,
}

impl Display for ProjectModuleReference<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "Module Reference: Name='{}', Path='{}'",
            self.name, self.path
        )
    }
}

/// Represents a reference to a class in a VB6 project.
#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, Hash)]
pub struct ProjectClassReference<'a> {
    /// The name of the class.
    pub name: &'a str,
    /// The path to the class file.
    pub path: &'a str,
}

impl Display for ProjectClassReference<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "Class Reference: Name='{}', Path='{}'",
            self.name, self.path
        )
    }
}

/// The result type for parsing a VB6 project file.
///
/// Contains the parsed `ProjectFile` and any `ProjectErrorKind` errors encountered during parsing.
///
/// This is a type alias for `ParseResult<'a, ProjectFile<'a>>`.
pub type ProjectResult<'a> = ParseResult<'a, ProjectFile<'a>>;

impl<'a> ProjectFile<'a> {
    ///
    /// Creates an empty project file with default values.
    ///
    /// This is an internal helper function used by the parser to initialize
    /// a new `ProjectFile` before populating it with parsed values.
    fn new_empty() -> Self {
        ProjectFile {
            project_type: CompileTargetType::Exe,
            references: vec![],
            objects: vec![],
            modules: vec![],
            classes: vec![],
            designers: vec![],
            forms: vec![],
            user_controls: vec![],
            user_documents: vec![],
            related_documents: vec![],
            property_pages: vec![],
            other_properties: HashMap::new(),
            properties: ProjectProperties {
                // We default to using NativeCode because all the possible options
                // sit on this branch of the enum, while the other branch (PCode)
                // has no other options.
                //
                // Hence, if we have a NativeCode value, then we can place the
                // parsed value within it. If on the other hand it is PCode, then
                // we know the compilation type was selected as PCode and we can
                // simply ignore any of the NativeCode options since they will
                // not be used.
                compilation_type: CompilationType::default(),
                ..Default::default()
            },
        }
    }

    /// Returns an iterator over the project references.
    ///
    /// # Returns
    ///
    /// An iterator over references to `ProjectReference` items.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Class=Class1; Class1.cls
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///    for failure in failures.iter() {
    ///        failure.print();
    ///    }
    /// }
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.references().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn references(&self) -> impl Iterator<Item = &ProjectReference<'a>> {
        self.references.iter()
    }

    /// Returns an iterator over the project modules.
    ///
    /// # Returns
    ///
    /// An iterator over references to `ProjectModuleReference` items.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Class=Class1; Class1.cls
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.modules().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn modules(&self) -> impl Iterator<Item = &ProjectModuleReference<'a>> {
        self.modules.iter()
    }

    /// Returns an iterator over the project classes.
    ///
    /// # Returns
    ///
    /// An iterator over references to `ProjectClassReference` items.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Class=Class1; Class1.cls
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.classes().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn classes(&self) -> impl Iterator<Item = &ProjectClassReference<'a>> {
        self.classes.iter()
    }

    /// Returns an iterator over the project object references.
    ///
    /// # Returns
    ///
    /// An iterator over references to `ObjectReference` items.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Class=Class1; Class1.cls
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.objects().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn objects(&self) -> impl Iterator<Item = &ObjectReference> {
        self.objects.iter()
    }

    /// Returns an iterator over the project forms.
    ///
    /// # Returns
    ///
    /// An iterator over references to form file names.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Form=Form1.frm
    /// Form=Form2.frm
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.forms().collect::<Vec<_>>().len(), 2);
    /// ```
    pub fn forms(&self) -> impl Iterator<Item = &&'a str> {
        self.forms.iter()
    }

    /// Returns an iterator over the project user controls.
    ///
    /// # Returns
    ///
    /// An iterator over references to user control file names.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// UserControl=UserControl1.ctl
    /// Form=Form1.frm
    /// Form=Form2.frm
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.user_controls().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn user_controls(&self) -> impl Iterator<Item = &&'a str> {
        self.user_controls.iter()
    }

    /// Returns an iterator over the project user documents.
    ///
    /// # Returns
    ///
    /// An iterator over references to user document file names.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Form=Form2.frm
    /// UserDocument=UserDocument1.udd
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.user_documents().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn user_documents(&self) -> impl Iterator<Item = &&'a str> {
        self.user_documents.iter()
    }

    /// Returns an iterator over the project designers.
    ///
    /// # Returns
    ///
    /// An iterator over references to designer file names.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Form=Form2.frm
    /// Designer=Designer1.des
    /// UserDocument=UserDocument1.udd
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.designers().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn designers(&self) -> impl Iterator<Item = &&'a str> {
        self.designers.iter()
    }

    /// Returns an iterator over the project related documents.
    ///
    /// # Returns
    ///
    /// An iterator over references to related document file names.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Form=Form2.frm
    /// RelatedDoc=RelatedDocument1.rdt
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.related_documents().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn related_documents(&self) -> impl Iterator<Item = &&'a str> {
        self.related_documents.iter()
    }

    /// Returns an iterator over the project property pages.
    ///
    /// # Returns
    ///
    /// An iterator over references to property page file names.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Form=Form2.frm
    /// PropertyPage=PropertyPage1.ppg
    /// UserDocument=UserDocument1.udd
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.property_pages().collect::<Vec<_>>().len(), 1);
    /// ```
    pub fn property_pages(&self) -> impl Iterator<Item = &&'a str> {
        self.property_pages.iter()
    }

    /// Returns a reference to the other properties map.
    ///
    /// # Returns
    ///
    /// A reference to the `other_properties` map.
    ///
    /// # Example
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Form=Form2.frm
    /// PropertyPage=PropertyPage1.ppg
    /// UserDocument=UserDocument1.udd
    ///
    /// [ThirdPartySection]
    /// CustomProperty1=Value1
    /// "#;
    ///
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    ///
    /// let other_props = project.other_properties();
    /// assert_eq!(other_props.len(), 1);
    /// assert!(other_props.contains_key("ThirdPartySection"));
    ///
    /// let third_party_props = other_props.get("ThirdPartySection").unwrap();
    /// assert_eq!(third_party_props.len(), 1);
    /// assert_eq!(third_party_props.get("CustomProperty1").unwrap(), &"Value1");
    /// ```
    #[must_use]
    pub fn other_properties(&self) -> &HashMap<&'a str, HashMap<&'a str, &'a str>> {
        &self.other_properties
    }

    /// Parses a VB6 project file using a dispatch-based property handler system.
    ///
    /// This method uses a registry of property handlers to parse VB6 project files.
    /// The dispatch system provides better maintainability compared to the previous
    /// large match statement approach.
    ///
    /// # Architecture
    ///
    /// The parsing process follows these steps:
    /// 1. Initialize an empty project and property handler registry
    /// 2. Loop through each line of the input
    /// 3. Skip empty lines and handle section headers for third-party properties
    /// 4. Parse the property name
    /// 5. Dispatch to the appropriate handler based on the property name
    /// 6. Collect any errors without stopping the parse
    ///
    /// # Error Handling
    ///
    /// This parser uses an error-collecting approach rather than fail-fast. When
    /// an error occurs, it's added to the failures vector and parsing continues.
    /// This allows reporting multiple errors in a single pass.
    ///
    /// # Returns
    ///
    /// A `ProjectResult` containing the parsed project (if successful) and any
    /// errors encountered during parsing.
    ///
    /// # Panics
    ///
    /// This function can panic if the input is not a valid VB6 project file.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::*;
    /// use vb6parse::files::project::properties::CompileTargetType;
    ///
    /// let input = r#"Type=Exe
    /// Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
    /// Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
    /// Module=Module1; Module1.bas
    /// Class=Class1; Class1.cls
    /// Form=Form1.frm
    /// Form=Form2.frm
    /// UserControl=UserControl1.ctl
    /// UserDocument=UserDocument1.uds
    /// ExeName32="Project1.exe"
    /// Command32=""
    /// Path32=""
    /// Name="Project1"
    /// HelpContextID="0"
    /// CompatibleMode="0"
    /// MajorVer=1
    /// MinorVer=0
    /// RevisionVer=0
    /// AutoIncrementVer=0
    /// StartMode=0
    /// Unattended=0
    /// Retained=0
    /// ThreadPerObject=0
    /// MaxNumberOfThreads=1
    /// DebugStartupOption=0
    /// NoControlUpgrade=0
    /// ServerSupportFiles=0
    /// VersionCompanyName="Company Name"
    /// VersionFileDescription="File Description"
    /// VersionLegalCopyright="Copyright"
    /// VersionLegalTrademarks="Trademark"
    /// VersionProductName="Product Name"
    /// VersionComments="Comments"
    /// CompilationType=0
    /// OptimizationType=0
    /// FavorPentiumPro(tm)=0
    /// CodeViewDebugInfo=0
    /// NoAliasing=0
    /// BoundsCheck=0
    /// OverflowCheck=0
    /// FlPointCheck=0
    /// FDIVCheck=0
    /// UnroundedFP=0
    /// CondComp=""
    /// ResFile32=""
    /// IconForm=""
    /// Startup="Form1"
    /// HelpFile=""
    /// Title="Project1"
    /// [MS Transaction Server]
    /// AutoRefresh=1
    /// "#;
    /// let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => {
    ///         e.print();
    ///         panic!("failed to decode project source code.");
    ///     }
    /// };
    ///
    /// let result = ProjectFile::parse(&project_source_file);
    ///
    /// let (project_opt, failures) = result.unpack();
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    ///     panic!("Failed to parse project with {} errors.", failures.len());
    /// }
    ///
    /// let project = project_opt.expect("Expected project to be parsed successfully.");
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.references().collect::<Vec<_>>().len(), 1);
    /// assert_eq!(project.objects().collect::<Vec<_>>().len(), 1);
    /// assert_eq!(project.modules().collect::<Vec<_>>().len(), 1);
    /// assert_eq!(project.classes().collect::<Vec<_>>().len(), 1);
    /// assert_eq!(project.designers().collect::<Vec<_>>().len(), 0);
    /// assert_eq!(project.forms().collect::<Vec<_>>().len(), 2);
    /// assert_eq!(project.user_controls().collect::<Vec<_>>().len(), 1);
    /// assert_eq!(project.user_documents().collect::<Vec<_>>().len(), 1);
    /// assert_eq!(project.properties.startup, "Form1");
    /// assert_eq!(project.properties.title, "Project1");
    /// assert_eq!(project.properties.exe_32_file_name, "Project1.exe");
    /// ```
    #[must_use]
    pub fn parse(source_file: &'a SourceFile) -> ProjectResult<'a> {
        let mut project = ProjectFile::new_empty();
        let mut input = source_file.source_stream();
        let mut other_property_group: Option<&str> = None;
        let mut ctx = ParserContext::new(input.file_name(), input.contents);
        let handlers = PropertyHandlers::new();

        while !input.is_empty() {
            // Skip empty lines
            if skip_empty_lines(&mut input) {
                continue;
            }

            let line_start = input.start_of_line();

            // Handle section headers
            match parse_section_header_line(&mut ctx, &mut input) {
                Some(SectionHeaderDetection::HeaderName(section_header)) => {
                    handle_section_header(
                        &mut ctx,
                        &mut project,
                        section_header,
                        &mut other_property_group,
                    );
                    continue;
                }
                Some(SectionHeaderDetection::MalformedHeader) => {
                    continue;
                }
                None => {
                    // Not a section header line, parse the line as a normal
                    // VB6 project property line.
                }
            }

            // Parse property name
            let Some(property_name) = parse_property_name(&mut ctx, &mut input) else {
                continue;
            };

            // Handle third-party properties
            if let Some(group) = other_property_group {
                handle_other_property(&mut ctx, &mut input, &mut project, group, property_name);
                continue;
            }

            // Dispatch to appropriate handler
            if !handlers.handle(&mut ctx, &mut input, &mut project, property_name) {
                handle_unknown_property(&mut ctx, &mut input, line_start, property_name);
            }
        }

        ParseResult::new(Some(project), ctx.into_errors())
    }

    /// Gets a collection of references to all sub-project references in the project.
    ///
    /// # Returns
    ///
    /// A vector of references to all sub-project references.
    ///
    #[must_use]
    pub fn subproject_references(&self) -> Vec<&ProjectReference<'a>> {
        self.references
            .iter()
            .filter(|reference| matches!(reference, ProjectReference::SubProject { .. }))
            .collect::<Vec<_>>()
    }

    /// Gets a collection of all project references.
    ///
    /// # Returns
    ///
    /// A vector of all project references.
    ///
    #[must_use]
    pub fn project_references(&self) -> &Vec<ProjectReference<'a>> {
        &self.references
    }

    /// Gets a collection of references to all compiled references in the project.
    ///
    /// # Returns
    ///
    /// A vector of references to all compiled references.
    ///
    #[must_use]
    pub fn compiled_references(&self) -> Vec<&ProjectReference<'a>> {
        self.references
            .iter()
            .filter(|reference| matches!(reference, ProjectReference::Compiled { .. }))
            .collect::<Vec<_>>()
    }

    /// Gets a collection of mutable references to all sub-project references in the project.
    ///
    /// # Returns
    ///
    /// A vector of mutable references to all sub-project references.
    ///
    #[must_use]
    pub fn subproject_references_mut(&mut self) -> Vec<&mut ProjectReference<'a>> {
        self.references
            .iter_mut()
            .filter(|reference| matches!(reference, ProjectReference::SubProject { .. }))
            .collect::<Vec<_>>()
    }

    /// Gets a collection of mutable references to all compiled references in the project.
    ///
    /// # Returns
    ///
    /// A vector of mutable references to all compiled references.
    ///
    #[must_use]
    pub fn compiled_references_mut(&mut self) -> Vec<&mut ProjectReference<'a>> {
        self.references
            .iter_mut()
            .filter(|reference| matches!(reference, ProjectReference::Compiled { .. }))
            .collect::<Vec<_>>()
    }

    /// Gets a mutable reference to the collection of all project references.
    ///
    /// # Returns
    ///
    /// A mutable reference to the vector of all project references.
    ///
    #[must_use]
    pub fn project_references_mut(&mut self) -> &mut Vec<ProjectReference<'a>> {
        &mut self.references
    }
}

/// Type alias for property handler functions.
///
/// Property handlers take a mutable reference to the project being built,
/// a mutable reference to the input stream, the property name, and a mutable
/// reference to the failures vector for error collection.
type PropertyHandler<'a> =
    fn(&mut ParserContext<'a>, &mut SourceStream<'a>, &mut ProjectFile<'a>, &'a str);

/// Registry of property handlers for dispatching property parsing.
///
/// This structure provides a dispatch mechanism that maps VB6 project property names
/// to their corresponding handler functions. This design replaces the previous giant
/// match statement with a more maintainable and extensible system.
///
/// # Design
///
/// The registry uses a `HashMap` to associate property names with handler functions,
/// allowing O(1) lookups and making it easy to add new properties without modifying
/// a large match statement.
///
/// # Property Categories
///
/// The handlers are organized into the following categories:
/// - **File references**: `Type`, `Designer`, `Reference`, `Object`, `Module`, `Class`, `Form`, etc.
/// - **Basic metadata**: `ResFile32`, `IconForm`, `Startup`, `HelpFile`, `Title`, `ExeName32`, etc.
/// - **Version information**: `MajorVer`, `MinorVer`, `VersionCompanyName`, etc.
/// - **Compatibility**: `CompatibleMode`, `VersionCompatible32`, etc.
/// - **Compilation**: `CompilationType`, `OptimizationType`, `BoundsCheck`, etc.
/// - **Threading**: `StartMode`, `ThreadPerObject`, `MaxNumberOfThreads`, etc.
/// - **Debug settings**: `DebugStartupComponent`, `DebugStartupOption`, etc.
struct PropertyHandlers<'a> {
    handlers: HashMap<&'static str, PropertyHandler<'a>>,
}

impl<'a> PropertyHandlers<'a> {
    /// Creates a new property handlers registry with all standard VB6 properties registered.
    fn new() -> Self {
        let mut handlers: HashMap<&'static str, PropertyHandler<'a>> = HashMap::new();

        // File references
        handlers.insert("Type", handle_type);
        handlers.insert("Designer", handle_designer);
        handlers.insert("Reference", handle_reference);
        handlers.insert("Object", handle_object);
        handlers.insert("Module", handle_module);
        handlers.insert("Class", handle_class);
        handlers.insert("Form", handle_form);
        handlers.insert("UserControl", handle_user_control);
        handlers.insert("UserDocument", handle_user_document);
        handlers.insert("RelatedDoc", handle_related_doc);
        handlers.insert("PropertyPage", handle_property_page);

        // Basic project metadata
        handlers.insert("ResFile32", handle_res_file_32);
        handlers.insert("IconForm", handle_icon_form);
        handlers.insert("Startup", handle_startup);
        handlers.insert("HelpFile", handle_help_file);
        handlers.insert("Title", handle_title);
        handlers.insert("ExeName32", handle_exe_name_32);
        handlers.insert("Path32", handle_path_32);
        handlers.insert("Command32", handle_command_32);
        handlers.insert("Name", handle_name);
        handlers.insert("Description", handle_description);

        // Version information
        handlers.insert("MajorVer", handle_major_ver);
        handlers.insert("MinorVer", handle_minor_ver);
        handlers.insert("RevisionVer", handle_revision_ver);
        handlers.insert("AutoIncrementVer", handle_auto_increment_ver);
        handlers.insert("VersionCompanyName", handle_version_company_name);
        handlers.insert("VersionFileDescription", handle_version_file_description);
        handlers.insert("VersionLegalCopyright", handle_version_legal_copyright);
        handlers.insert("VersionLegalTrademarks", handle_version_legal_trademarks);
        handlers.insert("VersionProductName", handle_version_product_name);
        handlers.insert("VersionComments", handle_version_comments);

        // Compatibility settings
        handlers.insert("HelpContextID", handle_help_context_id);
        handlers.insert("CompatibleMode", handle_compatible_mode);
        handlers.insert("VersionCompatible32", handle_version_compatible_32);
        handlers.insert("CompatibleEXE32", handle_compatible_exe_32);

        // DLL/Component settings
        handlers.insert("DllBaseAddress", handle_dll_base_address);
        handlers.insert("RemoveUnusedControlInfo", handle_remove_unused_control_info);

        // Compilation settings
        handlers.insert("CompilationType", handle_compilation_type);
        handlers.insert("OptimizationType", handle_optimization_type);
        handlers.insert("FavorPentiumPro(tm)", handle_favor_pentium_pro);
        handlers.insert("CodeViewDebugInfo", handle_code_view_debug_info);
        handlers.insert("NoAliasing", handle_no_aliasing);
        handlers.insert("BoundsCheck", handle_bounds_check);
        handlers.insert("OverflowCheck", handle_overflow_check);
        handlers.insert("FlPointCheck", handle_fl_point_check);
        handlers.insert("FDIVCheck", handle_fdiv_check);
        handlers.insert("UnroundedFP", handle_unrounded_fp);
        handlers.insert("CondComp", handle_cond_comp);

        // Threading & runtime settings
        handlers.insert("StartMode", handle_start_mode);
        handlers.insert("Unattended", handle_unattended);
        handlers.insert("Retained", handle_retained);
        handlers.insert("ThreadPerObject", handle_thread_per_object);
        handlers.insert("MaxNumberOfThreads", handle_max_number_of_threads);
        handlers.insert("ThreadingModel", handle_threading_model);

        // Debug & development
        handlers.insert("DebugStartupComponent", handle_debug_startup_component);
        handlers.insert("DebugStartupOption", handle_debug_startup_option);
        handlers.insert("UseExistingBrowser", handle_use_existing_browser);
        handlers.insert("NoControlUpgrade", handle_no_control_upgrade);
        handlers.insert("ServerSupportFiles", handle_server_support_files);

        Self { handlers }
    }

    /// Handles a property by dispatching to the appropriate handler function.
    ///
    /// Returns `true` if the property was handled, `false` if no handler was found.
    fn handle(
        &self,
        ctx: &mut ParserContext<'a>,
        input: &mut SourceStream<'a>,
        project: &mut ProjectFile<'a>,
        property_name: &'a str,
    ) -> bool {
        if let Some(handler) = self.handlers.get(property_name) {
            handler(ctx, input, project, property_name);
            true
        } else {
            false
        }
    }
}

/// Skips empty lines in the input stream.
///
/// This function consumes any ASCII whitespace followed by a newline character.
/// It's used to skip blank lines in VB6 project files.
///
/// # Returns
///
/// Returns `true` if an empty line was skipped, `false` otherwise.
fn skip_empty_lines(input: &mut SourceStream) -> bool {
    let _ = input.take_ascii_whitespaces();
    input.take_newline().is_some()
}

/// Handles a section header by registering it in the project's `other_properties` map.
///
/// VB6 project files can contain custom sections for third-party components,
/// marked with section headers like `[MS Transaction Server]`. This function
/// creates a new `HashMap` entry for such sections and updates the tracking variable.
///
/// # Arguments
///
/// * `project` - The project file being constructed
/// * `section_header` - The name of the section (without brackets)
/// * `other_property_group` - Mutable reference to track the current section
fn handle_section_header<'a>(
    _ctx: &mut ParserContext<'a>,
    project: &mut ProjectFile<'a>,
    section_header: &'a str,
    other_property_group: &mut Option<&'a str>,
) {
    if !project.other_properties.contains_key(section_header) {
        project
            .other_properties
            .insert(section_header, HashMap::new());
        *other_property_group = Some(section_header);
    }
}

/// Handles a third-party property by parsing its value and storing it.
///
/// When parsing properties within a custom section (e.g., `[MS Transaction Server]`),
/// this function parses the property value and stores it in the appropriate `HashMap`.
///
/// # Arguments
///
/// * `project` - The project file being constructed
/// * `input` - The input stream containing the property value
/// * `group` - The section name this property belongs to
/// * `property_name` - The name of the property
/// * `failures` - Vector to collect any parsing errors
fn handle_other_property<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    group: &'a str,
    property_name: &'a str,
) {
    let Some(property_value) = parse_property_value(ctx, input, property_name) else {
        return;
    };

    if let Some(map) = project.other_properties.get_mut(group) {
        map.insert(property_name, property_value);
    }
}

/// Handles an unknown property by generating an error and skipping to the next line.
///
/// When a property name is not recognized as a standard VB6 property, this function
/// generates an error but allows parsing to continue with the next line.
///
/// # Arguments
///
/// * `input` - The input stream to skip to the next line
/// * `line_start` - The offset where the line started (for error reporting)
/// * `property_name` - The unrecognized property name
/// * `failures` - Vector to collect the error
fn handle_unknown_property<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    line_start: usize,
    property_name: &'a str,
) {
    input.forward_to_next_line();

    ctx.error(
        input.span_at(line_start),
        ProjectError::ParameterLineUnknown {
            line: property_name.to_string(),
        },
    );
}

// ============================================================================
// Property Handler Functions
// ============================================================================
//
// The following functions handle parsing of individual VB6 project properties.
// Each handler follows a consistent pattern:
// 1. Call the appropriate parsing function
// 2. On success, update the project structure
// 3. On failure, push the error to the failures vector
//
// This design allows for error collection without stopping the parse process,
// enabling the parser to report multiple errors in a single pass.

fn handle_type<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(project_type_value) = parse_converted_value(ctx, input, property_name) {
        project.project_type = project_type_value;
    }
}

fn handle_designer<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(designer) = parse_path_line(ctx, input, property_name) {
        project.designers.push(designer);
    }
}

fn handle_reference<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    _property_name: &'a str,
) {
    if let Some(reference) = parse_reference(ctx, input) {
        project.references.push(reference);
    }
}

fn handle_object<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    _property_name: &'a str,
) {
    if let Some(object) = parse_object(ctx, input) {
        project.objects.push(object);
    }
}

fn handle_module<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    _property_name: &'a str,
) {
    if let Some(module) = parse_module(ctx, input) {
        project.modules.push(module);
    }
}

fn handle_class<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    _property_name: &'a str,
) {
    if let Some(class) = parse_class(ctx, input) {
        project.classes.push(class);
    }
}

fn handle_form<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(form) = parse_path_line(ctx, input, property_name) {
        project.forms.push(form);
    }
}

fn handle_user_control<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(user_control) = parse_path_line(ctx, input, property_name) {
        project.user_controls.push(user_control);
    }
}

fn handle_user_document<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(user_document) = parse_path_line(ctx, input, property_name) {
        project.user_documents.push(user_document);
    }
}

fn handle_related_doc<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(related_document) = parse_path_line(ctx, input, property_name) {
        project.related_documents.push(related_document);
    }
}

fn handle_property_page<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(property_page_value) = parse_path_line(ctx, input, property_name) {
        project.property_pages.push(property_page_value);
    }
}

fn handle_res_file_32<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(res_32_file) = parse_quoted_value(ctx, input, property_name) {
        project.properties.res_file_32_path = res_32_file;
    }
}

fn handle_icon_form<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(icon_form_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.icon_form = icon_form_value;
    }
}

fn handle_startup<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(startup_value) = parse_optional_quoted_value(ctx, input, property_name) {
        project.properties.startup = startup_value;
    }
}

fn handle_help_file<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(help_file) = parse_quoted_value(ctx, input, property_name) {
        project.properties.help_file_path = help_file;
    }
}

fn handle_title<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(title_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.title = title_value;
    }
}

fn handle_exe_name_32<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(exe_32_file_name_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.exe_32_file_name = exe_32_file_name_value;
    }
}

fn handle_path_32<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(path_32_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.path_32 = path_32_value;
    }
}

fn handle_command_32<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(command_line_arguments_value) =
        parse_optional_quoted_value(ctx, input, property_name)
    {
        project.properties.command_line_arguments = command_line_arguments_value;
    }
}

fn handle_name<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(name_value) = parse_optional_quoted_value(ctx, input, property_name) {
        project.properties.name = name_value;
    }
}

fn handle_description<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(description_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.description = description_value;
    }
}

fn handle_major_ver<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(major_value) = parse_numeric(ctx, input, property_name) {
        project.properties.version_info.major = major_value;
    }
}

fn handle_minor_ver<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(minor_value) = parse_numeric(ctx, input, property_name) {
        project.properties.version_info.minor = minor_value;
    }
}

fn handle_revision_ver<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(revision_value) = parse_numeric(ctx, input, property_name) {
        project.properties.version_info.revision = revision_value;
    }
}

fn handle_auto_increment_ver<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(auto_increment_revision_value) = parse_numeric(ctx, input, property_name) {
        project.properties.version_info.auto_increment_revision = auto_increment_revision_value;
    }
}

fn handle_version_company_name<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(company_name_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.version_info.company_name = company_name_value;
    }
}

fn handle_version_file_description<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(file_description_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.version_info.file_description = file_description_value;
    }
}

fn handle_version_legal_copyright<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(copyright_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.version_info.copyright = copyright_value;
    }
}

fn handle_version_legal_trademarks<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(trademark_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.version_info.trademark = trademark_value;
    }
}

fn handle_version_product_name<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(product_name_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.version_info.product_name = product_name_value;
    }
}

fn handle_version_comments<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(comments_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.version_info.comments = comments_value;
    }
}

fn handle_help_context_id<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(help_context_id_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.help_context_id = help_context_id_value;
    }
}

fn handle_compatible_mode<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(compatibility_mode_value) = parse_quoted_converted_value(ctx, input, property_name)
    {
        project.properties.compatibility_mode = compatibility_mode_value;
    }
}

fn handle_version_compatible_32<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(version_32_compatibility_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.version_32_compatibility = version_32_compatibility_value;
    }
}

fn handle_compatible_exe_32<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(exe_32_compatible_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.exe_32_compatible = exe_32_compatible_value;
    }
}

fn handle_dll_base_address<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    _property_name: &'a str,
) {
    if let Some(dll_base_address_value) = parse_dll_base_address(ctx, input) {
        project.properties.dll_base_address = dll_base_address_value;
    }
}

fn handle_remove_unused_control_info<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(unused_control_info_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.unused_control_info = unused_control_info_value;
    }
}

fn handle_compilation_type<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(compilation_type) = parse_numeric(ctx, input, property_name) {
        project.properties.compilation_type = compilation_type;
    }
}

fn handle_optimization_type<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(optimization_type_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_optimization_type(optimization_type_value);
    }
}

fn handle_favor_pentium_pro<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(favor_pentium_pro_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_favor_pentium_pro(favor_pentium_pro_value);
    }
}

fn handle_code_view_debug_info<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(code_view_debug_info_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_code_view_debug_info(code_view_debug_info_value);
    }
}

fn handle_no_aliasing<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(aliasing_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_aliasing(aliasing_value);
    }
}

fn handle_bounds_check<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(bounds_check_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_bounds_check(bounds_check_value);
    }
}

fn handle_overflow_check<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(overflow_check_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_overflow_check(overflow_check_value);
    }
}

fn handle_fl_point_check<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(floating_point_check_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_floating_point_check(floating_point_check_value);
    }
}

fn handle_fdiv_check<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(pentium_fdiv_bug_check_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_pentium_fdiv_bug_check(pentium_fdiv_bug_check_value);
    }
}

fn handle_unrounded_fp<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(unrounded_floating_point_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.compilation_type = project
            .properties
            .compilation_type
            .update_unrounded_floating_point(unrounded_floating_point_value);
    }
}

fn handle_cond_comp<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(conditional_compile_value) = parse_quoted_value(ctx, input, property_name) {
        project.properties.conditional_compile = conditional_compile_value;
    }
}

fn handle_start_mode<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(start_mode_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.start_mode = start_mode_value;
    }
}

fn handle_unattended<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(unattended_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.unattended = unattended_value;
    }
}

fn handle_retained<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(retained_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.retained = retained_value;
    }
}

fn handle_thread_per_object<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(thread_per_object_value) = parse_numeric::<i16>(ctx, input, property_name) {
        if thread_per_object_value <= 0 {
            project.properties.thread_per_object = 0;
        } else {
            project.properties.thread_per_object = thread_per_object_value.cast_unsigned();
        }
    }
}

fn handle_max_number_of_threads<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(max_number_of_threads_value) = parse_numeric(ctx, input, property_name) {
        project.properties.max_number_of_threads = max_number_of_threads_value;
    }
}

fn handle_threading_model<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(threading_model_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.threading_model = threading_model_value;
    }
}

fn handle_debug_startup_component<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(debug_startup_component_value) = parse_path_line(ctx, input, property_name) {
        project.properties.debug_startup_component = debug_startup_component_value;
    }
}

fn handle_debug_startup_option<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(debug_startup_option_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.debug_startup_option = debug_startup_option_value;
    }
}

fn handle_use_existing_browser<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(use_existing_browser_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.use_existing_browser = use_existing_browser_value;
    }
}

fn handle_no_control_upgrade<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(upgrade_controls_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.upgrade_controls = upgrade_controls_value;
    }
}

fn handle_server_support_files<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
    project: &mut ProjectFile<'a>,
    property_name: &'a str,
) {
    if let Some(server_support_files_value) = parse_converted_value(ctx, input, property_name) {
        project.properties.server_support_files = server_support_files_value;
    }
}

enum SectionHeaderDetection<'a> {
    HeaderName(&'a str),
    MalformedHeader,
}

fn parse_section_header_line<'a>(
    ctx: &mut ParserContext<'a>,
    input: &mut SourceStream<'a>,
) -> Option<SectionHeaderDetection<'a>> {
    let line_start = input.start_of_line();

    // We want to grab any section header lines like '[MS Transaction Server]'.
    // Which we will use in parsing 'other properties.'
    let _header_start = input.take("[", Comparator::CaseSensitive)?;

    // We have a section header line.
    let Some((other_property, _)) = input.take_until("]", Comparator::CaseSensitive) else {
        // We have a section header line but it is not terminated properly.
        ctx.error(
            input.span_at(line_start),
            ProjectError::UnterminatedSectionHeader,
        );
        input.forward_to_next_line();

        return Some(SectionHeaderDetection::MalformedHeader);
    };

    let _ = input.take("]", Comparator::CaseSensitive);
    input.forward_to_next_line();

    Some(SectionHeaderDetection::HeaderName(other_property))
}

fn parse_property_name<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
) -> Option<&'a str> {
    let line_start = input.start_of_line();

    // We want to grab the property name.
    let property_name = input.take_until("=", Comparator::CaseSensitive);

    match property_name {
        None => {
            // No property name found, so we can't parse this line.
            // Go to the next line and return the error.
            ctx.error(
                input.span_at(line_start),
                ProjectError::PropertyNameNotFound,
            );
            input.forward_to_next_line();

            None
        }
        Some((property_name, _)) => {
            // We only need the property name not the split on '=' value so we only
            // return the first of the pair in the line split.
            let _ = input.take("=", Comparator::CaseSensitive);

            Some(property_name)
        }
    }
}

fn parse_property_value<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Option<&'a str> {
    // An line starts with the line_type followed by '=', and a value.
    //
    // By this point in the parse the line_type and "=" component should be
    // stripped off since that is how we knew to use this particular parse.;
    let parameter_start = input.offset();

    let Some((parameter_value, _)) = input.take_until_newline() else {
        // No parameter value found, so we can't parse this line.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueNotFound {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    };

    if parameter_value.is_empty() {
        // No parameter value found, so we can't parse this line.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueNotFound {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    }

    Some(parameter_value)
}

fn parse_quoted_value<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Option<&'a str> {
    // An line starts with the line_type followed by '=', and a quoted value.
    //
    // By this point in the parse the line_type and "=" component should be
    // stripped off since that is how we knew to use this particular parse.
    let parameter_start = input.offset();

    let Some((parameter_value, _)) = input.take_until_newline() else {
        // No parameter value found, so we can't parse this line.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueNotFound {
                parameter_line_name: line_type.to_string(),
            },
        );

        return None;
    };

    if parameter_value.is_empty() {
        // No startup value found, so we can't parse this line.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueNotFound {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    }

    let start_quote_found = parameter_value.starts_with('"');
    let end_quote_found = parameter_value.ends_with('"');

    if !start_quote_found && end_quote_found {
        // The value ends with a quote but does not start with one.
        // This is an error, so we return an error.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueMissingOpeningQuote {
                parameter_line_name: line_type.to_string(),
            },
        );

        return None;
    }

    // we have to check the length like this because if we have only a single
    // quote, then obviously the string both starts and ends with a quote (even
    // if that is the same character!) which means we still need to report this
    // failure case.
    let start_without_end = start_quote_found && !end_quote_found;
    let start_with_length_one = start_quote_found && end_quote_found && parameter_value.len() == 1;

    if start_without_end || start_with_length_one {
        // The value starts with a quote but does not end with one.
        // This is an error, so we return an error.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueMissingClosingQuote {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    }

    if !start_quote_found && !end_quote_found {
        // The startup value does not start or end with a quote.
        // This is an error, so we return an error.

        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterWithoutDefaultValueMissingQuotes {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    }

    let parameter_value = &parameter_value[1..parameter_value.len() - 1];

    Some(parameter_value)
}

fn parse_optional_quoted_value<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Option<&'a str> {
    // An optional line starts with 'Startup=' (or another such option starting line)
    // and is followed by the quoted value, "!None!", or "(None)", or "!(None)!" to indicate the
    // parameter value is 'None'.
    //
    // By this point in the parse the "Startup=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let parameter_start = input.offset();

    let Some((parameter_value, _)) = input.take_until_newline() else {
        // No parameter value found, so we can't parse this line.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueNotFound {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    };

    if parameter_value.is_empty() {
        // No parameter value found, so we can't parse this line.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueNotFound {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    }

    if parameter_value == "\"(None)\""
        || parameter_value == "\"!None!\""
        || parameter_value == "\"!(None)!\""
        || parameter_value == "(None)"
        || parameter_value == "!None!"
        || parameter_value == "!(None)!"
    {
        // The parameter has a value of None.
        return Some("");
    }

    let start_quote_found = parameter_value.starts_with('"');
    let end_quote_found = parameter_value.ends_with('"');

    if !start_quote_found && end_quote_found {
        // The value ends with a quote but does not start with one.
        // This is an error, so we return an error.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueMissingOpeningQuote {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    }

    // we have to check the length like this because if we have only a single
    // quote, then obviously the string both starts and ends with a quote (even
    // if that is the same character!) which means we still need to report this
    // failure case.
    let start_without_end = start_quote_found && !end_quote_found;
    let start_with_end_length_one =
        start_quote_found && end_quote_found && parameter_value.len() == 1;

    if start_without_end || start_with_end_length_one {
        // The value starts with a quote but does not end with one.
        // This is an error, so we return an error.
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueMissingClosingQuote {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    }

    if !start_quote_found && !end_quote_found {
        // The parameter value does not start or end with a quote.
        // This is an error, so we return an error.

        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterWithoutDefaultValueMissingQuotes {
                parameter_line_name: line_type.to_string(),
            },
        );
        return None;
    }

    let parameter_value = &parameter_value[1..parameter_value.len() - 1];
    Some(parameter_value)
}

/// Formats all valid values for an enum type as a string.
///
/// Returns a comma-separated list of valid enum values in the format:
/// ```text
/// 'numeric value' "message"
/// ```
/// for each variant, with the final variant
/// being appended with:
/// ```text
/// ", and 'numeric value' "message"
/// ```
/// This makes it slightly nicer to read.
///
/// Long live the Oxford comma!
///
/// # Example
/// For an enum with values 0, 1, 2 this should return:
/// `'0' "No Compatibility", '1' "Project Compatibility", and '2' "Compatible Exe Mode"`
fn format_valid_enum_values<T>() -> String
where
    T: IntoEnumIterator + EnumMessage + Debug + Into<i16> + Copy,
{
    match T::iter()
        .map(|v| {
            let numeric: i16 = v.into();
            format!("'{:?}' {:#?}", numeric, v.get_message().unwrap_or(""))
        })
        .collect::<Vec<_>>()
        .split_last()
    {
        Some((last, elements)) => {
            format!("{}, and {}", elements.join(", "), last)
        } // we shoiuld never get a 'None' here since all
        // the enums should have multiple variants with values, but...
        None => String::new(),
    }
}

fn parse_quoted_converted_value<'a, T>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Option<T>
where
    T: 'a
        + TryFrom<&'a str, Error = String>
        + IntoEnumIterator
        + EnumMessage
        + Debug
        + Into<i16>
        + Default
        + Copy,
{
    // This function is used to parse a quoted value that is expected to be
    // converted into an enum value through TryFrom.
    // This kind of line starts with the line_type followed by '=', and a
    // quoted value.
    let parameter_start = input.offset();

    let text_to_newline = input.take_until_newline();

    let parameter_value = match text_to_newline {
        None => {
            // The input ends right after the equal!
            // weird error and indicates the system is basically done, but still need to
            // spit out a reasonable error message.
            let value_span = input.span_range(parameter_start - 1, parameter_start);
            // We don't have a value so we want the valid values.
            let valid_value_message = format_valid_enum_values::<T>();
            let error = ctx
                .error_with(
                    value_span,
                    ProjectError::ParameterValueNotFoundEOF {
                        parameter_line_name: line_type.to_string(),
                        valid_value_message,
                    },
                )
                .with_label(DiagnosticLabel::new(
                    value_span,
                    format!(
                        "'{line_type}' must have a double qouted value and end with a newline."
                    ),
                )) // only a start quote in the note since we already have the end quote value.
                .with_note(format!("{line_type}=\"{}\"", T::default().into()));
            ctx.push_error(error);
            return None;
        }
        Some((parameter_value, _)) => parameter_value,
    };

    let start_quote_found = parameter_value.starts_with('"');
    let end_quote_found = parameter_value.ends_with('"');

    if !start_quote_found && end_quote_found {
        // The value ends with a quote but does not start with one.
        // This is an error, so we return an error.
        let value_span = input.span_range(parameter_start, parameter_start + parameter_value.len());
        let error = ctx
            .error_with(
                value_span,
                ProjectError::ParameterValueMissingOpeningQuote {
                    parameter_line_name: line_type.to_string(),
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be surrounded by double quotes."),
            )) // only a start quote in the note since we already have the end quote value.
            .with_note(format!("{line_type}=\"{parameter_value}"));
        ctx.push_error(error);
        return None;
    }

    // we have to check the length like this because if we have only a single
    // quote, then obviously the string both starts and ends with a quote (even
    // if that is the same character!) which means we still need to report this
    // failure case.
    let start_and_end_qoute = start_quote_found && end_quote_found;
    let parameter_length_one = parameter_value.len() == 1;

    if start_and_end_qoute && parameter_length_one {
        // The value starts with a quote and is only a single character wide. This means the entire
        // parameter value consists of a single double qoute character: '"'
        let value_span = input.span_range(parameter_start, parameter_start + parameter_value.len());
        // We do not have a valid parameter value, so we return an error.
        let valid_value_message = format_valid_enum_values::<T>();
        let default_value = T::default().into();
        let note_message = format!("{line_type}=\"{default_value}\"");

        let error = ctx
            .error_with(
                value_span,
                ProjectError::ParameterValueMissingClosingQuoteAndValue {
                    parameter_line_name: line_type.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be surrounded by double quotes."),
            )) // only an end quote in the note since we already have the start quote value.
            .with_note(note_message);
        ctx.push_error(error);
        return None;
    }

    if start_quote_found && !end_quote_found {
        // The value ends with a quote but does not start with one.
        // This is an error, so we return an error.
        let value_span = input.span_range(parameter_start, parameter_start + parameter_value.len());
        let error = ctx
            .error_with(
                value_span,
                ProjectError::ParameterValueMissingClosingQuote {
                    parameter_line_name: line_type.to_string(),
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be surrounded by double quotes."),
            )) // only an end quote in the note since we already have the start quote value.
            .with_note(format!("{line_type}={parameter_value}\""));
        ctx.push_error(error);
        return None;
    }

    if !start_quote_found && !end_quote_found && parameter_length_one {
        // The value does not start or end with a quote but there *is* a number here.
        // this is not the same as not having an start or end and having a length of zero.
        // this is likely something like 'CompatibleMode=1' and needs to mention the
        // double qouting.

        // We do not have a valid parameter value, so we return an error.
        let valid_value_message = format_valid_enum_values::<T>();

        // We have a value, but it's not qouted. If the value makes sense
        // for this conversion, we should have the note show the qouted values.
        // if it's an invalid value, show the user an example with the default
        // value.
        let note_message = if T::try_from(parameter_value).is_ok() {
            format!("{line_type}=\"{parameter_value}\"")
        } else {
            let default_value = T::default().into();
            format!("{line_type}=\"{default_value}\"")
        };

        let value_span = input.span_at(parameter_start);
        let error = ctx
            .error_with(
                value_span,
                ProjectError::ParameterValueMissingQuotes {
                    parameter_line_name: line_type.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be contained within double qoutes."),
            ))
            .with_note(note_message);
        ctx.push_error(error);
        return None;
    }

    if !start_quote_found && !end_quote_found && !parameter_length_one {
        // The value does not start or end with a quote but there *is* a number here.
        // this is not the same as not having an start or end and having a length of zero.
        // this is likely something like 'CompatibleMode=' and needs to show the default vale.

        // We do not have a valid parameter value, so we return an error.
        let valid_value_message = format_valid_enum_values::<T>();

        // We don't have a value or qoutes.
        // show the user an example with the default value.

        let default_value = T::default().into();
        let note_message = format!("{line_type}=\"{default_value}\"");

        let value_span = input.span_at(parameter_start);
        let error = ctx
            .error_with(
                value_span,
                ProjectError::ParameterWithDefaultValueNotFound {
                    parameter_line_name: line_type.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(
                value_span,
                format!("'{line_type}' value must be one of the valid values contained within double qoutes."),
            ))
            .with_note(note_message);
        ctx.push_error(error);
        return None;
    }

    // trim off the quote characters.
    let parameter_value = &parameter_value[1..parameter_value.len() - 1];

    let Ok(value) = T::try_from(parameter_value) else {
        // We have a parameter value that is invalid, so we return an error.
        let valid_value_message = format_valid_enum_values::<T>();

        let value_span = input.span_at(parameter_start + 1);
        let error = ctx
            .error_with(
                value_span,
                ProjectError::ParameterValueInvalid {
                    parameter_line_name: line_type.to_string(),
                    invalid_value: parameter_value.to_string(),
                    valid_value_message,
                },
            )
            .with_label(DiagnosticLabel::new(value_span, "invalid value"))
            .with_note("Change the quoted value to one of the valid values.");
        ctx.push_error(error);
        return None;
    };

    Some(value)
}

fn parse_converted_value<'a, T>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Option<T>
where
    T: TryFrom<&'a str, Error = String> + IntoEnumIterator + EnumMessage + Debug,
{
    // This function is used to parse a value that is expected to be
    // converted into an enum value through TryFrom.
    // This kind of line starts with the line_type followed by '=', and a
    // value.
    let parameter_start = input.offset();

    let text_to_newline = input.take_until_newline();

    let parameter_value = match text_to_newline {
        None => {
            // No type text found, so we can't parse this line.
            // Go to the next line and return the error.
            ctx.error(
                input.span_at(parameter_start),
                ProjectError::ParameterValueMissingOpeningQuote {
                    parameter_line_name: line_type.to_string(),
                },
            );
            return None;
        }
        Some((parameter_value, _)) => parameter_value,
    };

    let Ok(value) = T::try_from(parameter_value) else {
        // We have a parameter value that is invalid, so we return an error.

        let valid_value_message = T::iter()
            .map(|v| format!("'{:?}' ({})", v, v.get_message().unwrap()))
            .collect::<Vec<_>>()
            .join(", ");

        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueInvalid {
                parameter_line_name: line_type.to_string(),
                invalid_value: parameter_value.to_string(),
                valid_value_message,
            },
        );
        return None;
    };

    Some(value)
}

fn parse_numeric<'a, T>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Option<T>
where
    T: FromStr,
{
    // This function is used to parse a value that is expected to be
    // converted into a value through TryFrom.
    // This kind of line starts with the line_type followed by '=', and a
    // value.
    let parameter_start = input.offset();

    let text_to_newline = input.take_until_newline();

    let parameter_value = match text_to_newline {
        None => {
            // No type text found, so we can't parse this line.
            // Go to the next line and return the error.
            ctx.error(
                input.span_at(parameter_start),
                ProjectError::ParameterValueMissingOpeningQuote {
                    parameter_line_name: line_type.to_string(),
                },
            );
            return None;
        }
        Some((parameter_value, _)) => parameter_value,
    };

    let Ok(value) = parameter_value.parse::<T>() else {
        // We have a parameter value that is invalid, so we return an error.
        let valid_value_message = format!(
            "Failed attempting to parse as {0}. '{parameter_value}' is not a valid {0}",
            std::any::type_name::<T>()
        );
        ctx.error(
            input.span_at(parameter_start),
            ProjectError::ParameterValueInvalid {
                parameter_line_name: line_type.to_string(),
                invalid_value: parameter_value.to_string(),
                valid_value_message,
            },
        );
        return None;
    };

    Some(value)
}

fn parse_reference<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
) -> Option<ProjectReference<'a>> {
    // A Reference line starts with a 'Reference=' and is followed by either a
    // project reference or a compiled reference.
    //
    // By this point in the parse the "Reference=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let reference_start = input.offset();

    // Compiled references start with "*\\G{" and are followed by a UUID.
    let compiled_reference_signature = "*\\G{";
    if input.peek(compiled_reference_signature.len()) == Some(compiled_reference_signature) {
        let _ = input.take(compiled_reference_signature, Comparator::CaseSensitive);
        // This is a compiled reference.
        return parse_compiled_reference(ctx, input);
    }

    // This is a project reference, but not a compiled reference.
    let Some((path, _)) = input.take_until_newline() else {
        // No path found, so we can't parse this line.
        ctx.error(
            input.span_at(reference_start),
            ProjectError::ReferenceProjectPathNotFound,
        );
        return None;
    };

    if path.is_empty() {
        // No path found, so we can't parse this line.
        ctx.error(
            input.span_at(reference_start),
            ProjectError::ReferenceProjectPathNotFound,
        );
        return None;
    }

    if !path.starts_with("*\\A") {
        // The path does not start with "*\A", which is not allowed.
        ctx.error(
            input.span_at(reference_start),
            ProjectError::ReferenceProjectPathInvalid {
                value: path.to_string(),
            },
        );
        return None;
    }

    let path = &path[3..]; // Strip off the "*\A" prefix

    Some(ProjectReference::SubProject { path })
}

fn parse_compiled_reference<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
) -> Option<ProjectReference<'a>> {
    // A compiled reference starts with "*\\G{" and is followed by a UUID.
    // We have already checked that the input starts with "*\\G{".
    // By this point in the parse the "*\\G{" component should be stripped off.
    let uuid_start = input.offset();

    // This is a compiled reference.
    let Some((uuid_text, _)) = input.take_until("}#", Comparator::CaseSensitive) else {
        // No UUID found, so we can't parse this line.
        ctx.error(
            input.span_at(uuid_start),
            ProjectError::ReferenceCompiledUuidMissingMatchingBrace,
        );
        input.forward_to_next_line();

        return None;
    };

    let Ok(uuid) = Uuid::parse_str(uuid_text) else {
        // The UUID is not a valid UUID, so we can't parse this line.
        ctx.error(
            input.span_at(uuid_start),
            ProjectError::ReferenceCompiledUuidInvalid,
        );
        input.forward_to_next_line();

        return None;
    };

    let _ = input.take("}#", Comparator::CaseSensitive);
    let unknown1_start = input.offset();

    let Some((unknown1, _)) = input.take_until("#", Comparator::CaseSensitive) else {
        // No unknown1 found, so we can't parse this line.
        ctx.error(
            input.span_at(unknown1_start),
            ProjectError::ReferenceCompiledUnknown1Missing,
        );
        input.forward_to_next_line();

        return None;
    };

    let _ = input.take("#", Comparator::CaseSensitive);
    let unknown2_start = input.offset();

    let Some((unknown2, _)) = input.take_until("#", Comparator::CaseSensitive) else {
        // No unknown2 found, so we can't parse this line.
        ctx.error(
            input.span_at(unknown2_start),
            ProjectError::ReferenceCompiledUnknown2Missing,
        );
        input.forward_to_next_line();

        return None;
    };

    let _ = input.take("#", Comparator::CaseSensitive);
    let path_start = input.offset();

    let Some((path, _)) = input.take_until("#", Comparator::CaseSensitive) else {
        // No path found, so we can't parse this line.
        ctx.error(
            input.span_at(path_start),
            ProjectError::ReferenceCompiledPathNotFound,
        );
        input.forward_to_next_line();

        return None;
    };

    let _ = input.take("#", Comparator::CaseSensitive);
    let description_start = input.offset();

    let Some((description, _)) = input.take_until_newline() else {
        // No description found, so we can't parse this line.
        ctx.error(
            input.span_at(description_start),
            ProjectError::ReferenceCompiledDescriptionNotFound,
        );
        return None;
    };

    if description.contains('#') {
        // The description contains a '#', which is not allowed.
        ctx.error(
            input.span_at(description_start),
            ProjectError::ReferenceCompiledDescriptionInvalid,
        );
        return None;
    }

    // We have a compiled reference.
    let reference = ProjectReference::Compiled {
        uuid,
        unknown1,
        unknown2,
        path,
        description,
    };

    Some(reference)
}

fn parse_object(ctx: &mut ParserContext, input: &mut SourceStream) -> Option<ObjectReference> {
    // An Object line starts with an 'Object=' and is followed by either a
    // compiled object or a project object.
    //
    // By this point in the parse the "Object=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let object_start = input.offset();

    // Project objects start with "\"*\\A" and are followed by the path to the
    // object ending with a single quote.
    // Usually this is a single file with a .vbp extension but we do not enforce that currently.
    let project_object_signature = "\"*\\A";
    if input.peek(project_object_signature.len()) == Some(project_object_signature) {
        let _ = input.take(project_object_signature, Comparator::CaseSensitive);
        // This is a project object.
        let object_path_start = input.offset();

        let Some((path, _)) = input.take_until("\"", Comparator::CaseSensitive) else {
            // No path found, so we can't parse this line.
            ctx.error(
                input.span_at(object_path_start),
                ProjectError::ObjectProjectPathNotFound,
            );
            input.forward_to_next_line();

            return None;
        };
        input.forward_to_next_line();

        return Some(ObjectReference::Project { path: path.into() });
    }

    // It looks like we have a compiled object line instead. Hopefully.
    if input.peek(1) != Some("{") {
        // We do not have a compiled object line, so we can't parse this line.
        ctx.error(
            input.span_at(object_start),
            ProjectError::ObjectCompiledMissingOpeningBrace,
        );
        input.forward_to_next_line();

        return None;
    }
    let _ = input.take("{", Comparator::CaseSensitive);

    let Some((uuid_text, _)) = input.take_until("}", Comparator::CaseSensitive) else {
        // No UUID found, so we can't parse this line.
        ctx.error(
            input.span_at(object_start),
            ProjectError::ObjectCompiledUuidMissingMatchingBrace,
        );
        input.forward_to_next_line();

        return None;
    };

    let _ = input.take("}", Comparator::CaseSensitive);

    let Ok(uuid) = Uuid::parse_str(uuid_text) else {
        // The UUID is not a valid UUID, so we can't parse this line.
        ctx.error(
            input.span_at(object_start),
            ProjectError::ObjectCompiledUuidInvalid,
        );
        input.forward_to_next_line();

        return None;
    };

    let _ = input.take("#", Comparator::CaseSensitive);

    let version_start = input.offset();
    let Some((version, invalid_version_character)) = input.take_until_not(
        &["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."],
        Comparator::CaseSensitive,
    ) else {
        // No version found, so we can't parse this line.
        ctx.error(
            input.span_at(version_start),
            ProjectError::ObjectCompiledVersionMissing,
        );
        input.forward_to_next_line();

        return None;
    };

    if invalid_version_character != "#" {
        // The version contains an invalid character, so we can't parse this line.
        ctx.error(
            input.span_at(version_start + version.len()),
            ProjectError::ObjectCompiledVersionInvalid,
        );
        input.forward_to_next_line();

        return None;
    }
    let _ = input.take("#", Comparator::CaseSensitive);
    let unknown1_start = input.offset();

    let Some((unknown1, _)) = input.take_until("; ", Comparator::CaseSensitive) else {
        // No unknown1 found, so we can't parse this line.
        ctx.error(
            input.span_at(unknown1_start),
            ProjectError::ObjectCompiledUnknown1Missing,
        );
        input.forward_to_next_line();

        return None;
    };
    let _ = input.take("; ", Comparator::CaseSensitive);
    let file_name_start = input.offset();

    let file_name = input.take_until_newline();
    match file_name {
        None => {
            // No file name found, so we can't parse this line.
            ctx.error(
                input.span_at(file_name_start),
                ProjectError::ObjectCompiledFileNameNotFound,
            );
            None
        }
        Some((file_name, _)) => {
            if file_name.is_empty() {
                // No file name found, so we can't parse this line.
                ctx.error(
                    input.span_at(file_name_start),
                    ProjectError::ObjectCompiledFileNameNotFound,
                );
                return None;
            }

            Some(ObjectReference::Compiled {
                uuid,
                version: version.into(),
                unknown1: unknown1.into(),
                file_name: file_name.into(),
            })
        }
    }
}

fn parse_module<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
) -> Option<ProjectModuleReference<'a>> {
    // A Module line starts with a 'Module=' and is followed by a name and a path.
    //
    // By this point in the parse the "Module=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let module_start = input.offset();

    let Some((module_name, _)) = input.take_until("; ", Comparator::CaseSensitive) else {
        // No name found, so we can't parse this line.
        ctx.error(
            input.span_at(module_start),
            ProjectError::ModuleNameNotFound,
        );
        input.forward_to_next_line();

        return None;
    };
    let _ = input.take("; ", Comparator::CaseSensitive);
    let module_path_start = input.offset();

    let Some((module_path, _)) = input.take_until_newline() else {
        // No path found, so we can't parse this line.
        ctx.error(
            input.span_at(module_path_start),
            ProjectError::ModuleFileNameNotFound,
        );
        return None;
    };

    if module_path.is_empty() {
        // No path found, so we can't parse this line.
        ctx.error(
            input.span_at(module_path_start),
            ProjectError::ModuleFileNameNotFound,
        );
        return None;
    }

    let module = ProjectModuleReference {
        name: module_name,
        path: module_path,
    };
    Some(module)
}

fn parse_class<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
) -> Option<ProjectClassReference<'a>> {
    // A Class line starts with a 'Class=' and is followed by a name and a path.
    //
    // By this point in the parse the "Class=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let class_start = input.offset();

    let Some((class_name, _)) = input.take_until("; ", Comparator::CaseSensitive) else {
        // No name found, so we can't parse this line.
        ctx.error(input.span_at(class_start), ProjectError::ClassNameNotFound);
        input.forward_to_next_line();

        return None;
    };

    let _ = input.take("; ", Comparator::CaseSensitive);
    let class_path_start = input.offset();

    let Some((class_path, _)) = input.take_until_newline() else {
        // No path found, so we can't parse this line.
        ctx.error(
            input.span_at(class_path_start),
            ProjectError::ClassFileNameNotFound,
        );
        return None;
    };

    if class_path.is_empty() {
        // No path found, so we can't parse this line.
        ctx.error(
            input.span_at(class_path_start),
            ProjectError::ClassFileNameNotFound,
        );
        return None;
    }

    let class = ProjectClassReference {
        name: class_name,
        path: class_path,
    };

    Some(class)
}

fn parse_path_line<'a>(
    ctx: &mut ParserContext,
    input: &mut SourceStream<'a>,
    parameter_line_name: &'a str,
) -> Option<&'a str> {
    // A single element line starts with a 'Form=', 'Designer=', or 'RelatedDoc='
    // and is followed by a path to the corresponding file.
    //
    // By this point in the parse the "Form=", 'Designer=', or 'RelatedDoc='
    // component should be stripped off since that is how we knew to use this
    // particular parse.
    let path_start = input.offset();

    let path_line = input.take_until_newline();
    match path_line {
        None => {
            // No file_path text found, so we can't parse this line.
            // Go to the next line and return the error.
            ctx.error(
                input.span_at(path_start),
                ProjectError::PathValueNotFound {
                    parameter_line_name: parameter_line_name.to_string(),
                },
            );
            None
        }
        Some((file_path, _)) => {
            if file_path.is_empty() {
                // No file_path text found, so we can't parse this line.
                // Go to the next line and return the error.
                ctx.error(
                    input.span_at(path_start),
                    ProjectError::PathValueNotFound {
                        parameter_line_name: parameter_line_name.to_string(),
                    },
                );
                return None;
            }

            Some(file_path)
        }
    }
}

fn parse_dll_base_address(ctx: &mut ParserContext, input: &mut SourceStream) -> Option<u32> {
    // A DllBaseAddress line starts with a 'DllBaseAddress=' and is followed by a
    // hexadecimal value.
    //
    // By this point in the parse the "DllBaseAddress=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let dll_base_address_start = input.offset();

    let Some((base_address_hex_text, _)) = input.take_until_newline() else {
        // No base address found, so we can't parse this line.
        ctx.error(
            input.span_at(dll_base_address_start),
            ProjectError::DllBaseAddressNotFound,
        );
        return None;
    };

    if base_address_hex_text.is_empty() {
        // The base address is empty, so we can't parse this line.
        ctx.error(
            input.span_at(dll_base_address_start),
            ProjectError::DllBaseAddressUnparsableEmpty,
        );
        return None;
    }

    if !base_address_hex_text.starts_with("&H") {
        // The base address does not start with "&H", so we can't parse this line.
        ctx.error(
            input.span_at(dll_base_address_start),
            ProjectError::DllBaseAddressMissingHexPrefix,
        );
        return None;
    }

    let dll_base_address_start = dll_base_address_start + 2; // Skip the "&H" prefix

    let trimmed_base_address_hex_text = base_address_hex_text.trim_start_matches("&H");

    let Ok(dll_base_address) = u32::from_str_radix(trimmed_base_address_hex_text, 16) else {
        // The base address is not a valid hexadecimal value, so we can't parse this line.
        ctx.error(
            input.span_at(dll_base_address_start),
            ProjectError::DllBaseAddressUnparsable {
                hex_value: trimmed_base_address_hex_text.to_string(),
            },
        );
        return None;
    };

    Some(dll_base_address)
}

#[cfg(test)]
mod tests {
    use crate::errors::{ErrorKind, ParserContext, ProjectError, Severity};
    use crate::files::common::ObjectReference;
    use crate::files::project::compilesettings::*;
    use crate::files::project::properties::*;
    use crate::files::project::ProjectReference;
    use crate::io::{Comparator, SourceFile, SourceStream};
    use crate::ProjectFile;
    use uuid::Uuid;

    #[test]
    fn compatibility_mode_eof_after_equal() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert!(matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueNotFoundEOF { .. })
        ));
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 15);
        assert_eq!(errors[0].labels[0].span.offset, 14);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' must have a double qouted value and end with a newline."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_is_invalid() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"5\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert!(matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueInvalid { .. })
        ));
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 21);
        assert_eq!(errors[0].labels[0].span.offset, 16);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(errors[0].labels[0].message, "invalid value");
        assert_eq!(
            errors[0].notes[0],
            "Change the quoted value to one of the valid values."
        );
    }

    #[test]
    fn compatibility_mode_without_qoutes() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=0\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert!(matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingQuotes { .. })
        ));
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 18);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be contained within double qoutes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"0\"");
    }

    #[test]
    fn compatibility_mode_invalid_without_qoutes() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=5\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert!(matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingQuotes { .. })
        ));
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 18);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be contained within double qoutes."
        ); // Since the unqouted value is invalid, we should show a note with the default for 'CompatibleMode'
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_without_value() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert!(matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterWithDefaultValueNotFound { .. })
        ));
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);

        assert_eq!(errors[0].labels[0].span.line_end, 16);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 1);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be one of the valid values contained within double qoutes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_without_end_qoute() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"1\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert!(matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingClosingQuote { .. })
        ));
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);
        assert_eq!(errors[0].labels[0].span.line_end, 19);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 2);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be surrounded by double quotes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"1\"");
    }

    #[test]
    fn compatibility_mode_without_start_qoute() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=2\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let _compatibility_mode: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert!(matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueMissingOpeningQuote { .. })
        ));
        assert_eq!(errors[0].severity, Severity::Error);
        assert_eq!(errors[0].labels.len(), 1);
        assert_eq!(errors[0].labels[0].span.line_start, 0);
        assert_eq!(errors[0].labels[0].span.line_end, 19);
        assert_eq!(errors[0].labels[0].span.offset, 15);
        assert_eq!(errors[0].labels[0].span.length, 2);
        assert_eq!(
            errors[0].labels[0].message,
            "'CompatibleMode' value must be surrounded by double quotes."
        );
        assert_eq!(errors[0].notes[0], "CompatibleMode=\"2\"");
    }

    #[test]
    fn compatibility_mode_is_no_compatibility() {
        use crate::files::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"0\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 0);
        assert_eq!(result.unwrap(), CompatibilityMode::NoCompatibility);
    }

    #[test]
    fn compatibility_mode_is_project() {
        use crate::files::project::parse_quoted_converted_value;
        use crate::io::{Comparator, SourceStream};

        let mut input = SourceStream::new("", "CompatibleMode=\"1\"\r\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 0);
        assert_eq!(result.unwrap(), CompatibilityMode::Project);
    }

    #[test]
    fn compatibility_mode_is_compatible_exe() {
        use crate::files::project::parse_quoted_converted_value;
        use crate::io::{Comparator, SourceStream};

        let mut input = SourceStream::new("", "CompatibleMode=\"2\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result: Option<CompatibilityMode> =
            parse_quoted_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 0);
        assert_eq!(result.unwrap(), CompatibilityMode::CompatibleExe);
    }

    #[test]
    fn project_type_is_exe() {
        use crate::files::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=Exe\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result: Option<CompileTargetType> =
            parse_converted_value(&mut ctx, &mut input, parameter_name);

        let errors = ctx.errors();

        assert_eq!(errors.len(), 0);
        assert_eq!(result.unwrap(), CompileTargetType::Exe);
    }

    #[test]
    fn project_type_is_oledll() {
        use crate::files::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=OleDll\r\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result: Option<CompileTargetType> =
            parse_converted_value(&mut ctx, &mut input, parameter_name);

        assert_eq!(result.unwrap(), CompileTargetType::OleDll);
    }

    #[test]
    fn project_type_is_control() {
        use crate::files::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=Control\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result: Option<CompileTargetType> =
            parse_converted_value(&mut ctx, &mut input, parameter_name);

        assert_eq!(result.unwrap(), CompileTargetType::Control);
    }

    #[test]
    fn project_type_is_ole_exe() {
        use crate::files::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=OleExe\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result: Option<CompileTargetType> =
            parse_converted_value(&mut ctx, &mut input, parameter_name);

        assert_eq!(result.unwrap(), CompileTargetType::OleExe);
    }

    #[test]
    fn project_type_is_unknown_type() {
        use crate::files::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=blah\r\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result: Option<CompileTargetType> =
            parse_converted_value(&mut ctx, &mut input, parameter_name);

        assert!(result.is_none());

        let errors = ctx.errors();

        assert_eq!(errors.len(), 1);
        assert!(matches!(
            *errors[0].kind,
            ErrorKind::Project(ProjectError::ParameterValueInvalid { .. })
        ));
    }

    #[test]
    fn reference_compiled_line_valid() {
        use crate::files::project::parse_reference;

        let mut input = SourceStream::new("", "Reference=*\\G{000440D8-E9ED-4435-A9A2-06B05387BB16}#c.0#0#..\\DBCommon\\Libs\\VbIntellisenseFix.dll#VbIntellisenseFix\r\n");

        let _ = input.take("Reference", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result = parse_reference(&mut ctx, &mut input);

        let expected_uuid = Uuid::parse_str("000440D8-E9ED-4435-A9A2-06B05387BB16").unwrap();

        assert!(input.is_empty());
        let result = result.unwrap();
        assert!(matches!(result, ProjectReference::Compiled { .. }));

        match result {
            ProjectReference::Compiled {
                uuid,
                unknown1,
                unknown2,
                path,
                description,
            } => {
                assert_eq!(uuid, expected_uuid);
                assert_eq!(unknown1, "c.0");
                assert_eq!(unknown2, "0");
                assert_eq!(path, r"..\DBCommon\Libs\VbIntellisenseFix.dll");
                assert_eq!(description, r"VbIntellisenseFix");
            }
            ProjectReference::SubProject { .. } => panic!("Expected a compiled reference"),
        }
    }

    #[test]
    fn reference_subproject_line_valid() {
        use crate::files::project::parse_reference;

        let mut input = SourceStream::new("", "Reference=*\\Atest.vbp\r\n");

        let _ = input.take("Reference", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result = parse_reference(&mut ctx, &mut input);

        assert!(input.is_empty());
        assert_eq!(
            result.unwrap(),
            ProjectReference::SubProject { path: "test.vbp" }
        );
    }

    #[test]
    fn module_line_valid() {
        use crate::files::project::parse_module;

        let mut input = SourceStream::new("", "Module=modDBAssist; ..\\DBCommon\\DBAssist.bas\r\n");

        let _ = input.take("Module", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result = parse_module(&mut ctx, &mut input).unwrap();

        assert!(input.is_empty());
        assert_eq!(result.name, "modDBAssist");
        assert_eq!(result.path, "..\\DBCommon\\DBAssist.bas");
    }

    #[test]
    fn class_line_valid() {
        use crate::files::project::parse_class;

        let mut input = SourceStream::new(
            "",
            "Class=CStatusBarClass; ..\\DBCommon\\CStatusBarClass.cls\r\n",
        );

        let _ = input.take("Class", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result = parse_class(&mut ctx, &mut input).unwrap();

        assert!(input.is_empty());
        assert_eq!(result.name, "CStatusBarClass");
        assert_eq!(result.path, "..\\DBCommon\\CStatusBarClass.cls");
    }

    #[test]
    fn object_line_valid() {
        use crate::files::project::parse_object;

        let mut input = SourceStream::new(
            "",
            "Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb\r\n",
        );

        let _ = input.take("Object", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let result = parse_object(&mut ctx, &mut input);

        let object = result.unwrap();

        assert!(input.is_empty());
        match object {
            ObjectReference::Compiled {
                uuid,
                version,
                unknown1,
                file_name,
            } => {
                let expected_uuid =
                    Uuid::parse_str("00020430-0000-0000-C000-000000000046").unwrap();
                assert_eq!(uuid, expected_uuid);
                assert_eq!(version, "2.0");
                assert_eq!(unknown1, "0");
                assert_eq!(file_name, "stdole2.tlb");
            }
            ObjectReference::Project { .. } => panic!("Expected a compiled object"),
        }
    }

    #[allow(clippy::too_many_lines)]
    #[test]
    fn thread_per_object_negative() {
        let input = r#"Type=Exe
     Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
     Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
     Module=Module1; Module1.bas
     Class=Class1; Class1.cls
     Form=Form1.frm
     Form=Form2.frm
     UserControl=UserControl1.ctl
     UserDocument=UserDocument1.uds
     ExeName32="Project1.exe"
     Command32=""
     Path32=""
     Name="Project1"
     HelpContextID="0"
     CompatibleMode="0"
     MajorVer=1
     MinorVer=0
     RevisionVer=0
     AutoIncrementVer=0
     StartMode=0
     Unattended=0
     Retained=0
     ThreadPerObject=-1
     MaxNumberOfThreads=1
     DebugStartupOption=0
     NoControlUpgrade=0
     ServerSupportFiles=0
     VersionCompanyName="Company Name"
     VersionFileDescription="File Description"
     VersionLegalCopyright="Copyright"
     VersionLegalTrademarks="Trademark"
     VersionProductName="Product Name"
     VersionComments="Comments"
     CompilationType=0
     OptimizationType=0
     FavorPentiumPro(tm)=0
     CodeViewDebugInfo=0
     NoAliasing=0
     BoundsCheck=0
     OverflowCheck=0
     FlPointCheck=0
     FDIVCheck=0
     UnroundedFP=0
     CondComp=""
     ResFile32=""
     IconForm=""
     Startup=!(None)!
     HelpFile=""
     Title="Project1"
    
     [MS Transaction Server]
     AutoRefresh=1
"#;

        let project_source_file = SourceFile::decode("project1.vbp", input.as_bytes()).unwrap();

        let result = ProjectFile::parse(&project_source_file);

        if result.has_failures() {
            for failure in result.failures() {
                failure.print();
            }

            panic!("Project parse had failures");
        }

        let project = result.unwrap();

        assert_eq!(project.project_type, CompileTargetType::Exe);
        assert_eq!(project.references.len(), 1);
        assert_eq!(project.objects.len(), 1);
        assert_eq!(project.modules.len(), 1);
        assert_eq!(project.classes.len(), 1);
        assert_eq!(project.designers.len(), 0);
        assert_eq!(project.forms.len(), 2);
        assert_eq!(project.user_controls.len(), 1);
        assert_eq!(project.user_documents.len(), 1);
        assert_eq!(
            project.properties.upgrade_controls,
            UpgradeControls::Upgrade
        );
        assert_eq!(project.properties.res_file_32_path, "");
        assert_eq!(project.properties.icon_form, "");
        assert_eq!(project.properties.startup, "");
        assert_eq!(project.properties.help_file_path, "");
        assert_eq!(project.properties.title, "Project1");
        assert_eq!(project.properties.exe_32_file_name, "Project1.exe");
        assert_eq!(project.properties.exe_32_compatible, "");
        assert_eq!(project.properties.command_line_arguments, "");
        assert_eq!(project.properties.path_32, "");
        assert_eq!(project.properties.name, "Project1");
        assert_eq!(project.properties.help_context_id, "0");
        assert_eq!(
            project.properties.compatibility_mode,
            CompatibilityMode::NoCompatibility
        );
        assert_eq!(project.properties.version_info.major, 1);
        assert_eq!(project.properties.version_info.minor, 0);
        assert_eq!(project.properties.version_info.revision, 0);
        assert_eq!(project.properties.version_info.auto_increment_revision, 0);
        assert_eq!(project.properties.version_info.company_name, "Company Name");
        assert_eq!(
            project.properties.version_info.file_description,
            "File Description"
        );
        assert_eq!(project.properties.version_info.trademark, "Trademark");
        assert_eq!(project.properties.version_info.product_name, "Product Name");
        assert_eq!(project.properties.version_info.comments, "Comments");
        assert_eq!(
            project.properties.server_support_files,
            ServerSupportFiles::Local,
            "server_support_files check"
        );
        assert_eq!(project.properties.conditional_compile, "");
        assert!(matches!(
            project.properties.compilation_type,
            CompilationType::NativeCode(NativeCodeSettings {
                optimization_type: OptimizationType::FavorFastCode,
                favor_pentium_pro: FavorPentiumPro::False,
                code_view_debug_info: CodeViewDebugInfo::NotCreated,
                aliasing: Aliasing::AssumeAliasing,
                bounds_check: BoundsCheck::CheckBounds,
                overflow_check: OverflowCheck::CheckOverflow,
                floating_point_check: FloatingPointErrorCheck::CheckFloatingPointError,
                pentium_fdiv_bug_check: PentiumFDivBugCheck::CheckPentiumFDivBug,
                unrounded_floating_point: UnroundedFloatingPoint::DoNotAllow
            })
        ));
        assert_eq!(project.properties.start_mode, StartMode::StandAlone);
        assert_eq!(project.properties.unattended, InteractionMode::Interactive);
        assert_eq!(project.properties.retained, Retained::UnloadOnExit);
        assert_eq!(project.properties.thread_per_object, 0);
        assert_eq!(project.properties.max_number_of_threads, 1);
        assert_eq!(
            project.properties.debug_startup_option,
            DebugStartupOption::WaitForComponentCreation,
            "debug_startup_option check"
        );
    }

    #[test]
    fn two_line_with_spaces() {
        use super::parse_converted_value;
        use super::parse_property_name;
        use super::parse_reference;

        let mut input = SourceStream::new(
            "project.vbp",
            r"Type=Exe
     Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation",
        );

        let _ = input.take_ascii_whitespaces();

        let mut ctx = ParserContext::new(input.file_name(), input.contents);

        let line_type = parse_property_name(&mut ctx, &mut input).unwrap();
        let type_result: Option<CompileTargetType> =
            parse_converted_value(&mut ctx, &mut input, line_type);

        assert!(type_result.is_some());
        assert_eq!(type_result.unwrap(), CompileTargetType::Exe);

        let _ = input.take_ascii_whitespaces();

        let _ = parse_property_name(&mut ctx, &mut input).unwrap();
        let reference_result = parse_reference(&mut ctx, &mut input);

        assert!(reference_result.is_some());
        let reference = reference_result.unwrap();
        assert_eq!(
            reference,
            ProjectReference::Compiled {
                uuid: Uuid::parse_str("00020430-0000-0000-C000-000000000046").unwrap(),
                unknown1: "2.0",
                unknown2: "0",
                path: r"C:\Windows\System32\stdole2.tlb",
                description: "OLE Automation",
            }
        );
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn no_startup_selected() {
        let input = r#"Type=Exe
     Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
     Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
     Module=Module1; Module1.bas
     Class=Class1; Class1.cls
     Form=Form1.frm
     Form=Form2.frm
     UserControl=UserControl1.ctl
     UserDocument=UserDocument1.uds
     ExeName32="Project1.exe"
     Command32=""
     Path32=""
     Name="Project1"
     HelpContextID="0"
     CompatibleMode="0"
     MajorVer=1
     MinorVer=0
     RevisionVer=0
     AutoIncrementVer=0
     StartMode=0
     Unattended=0
     Retained=0
     ThreadPerObject=0
     MaxNumberOfThreads=1
     DebugStartupOption=0
     NoControlUpgrade=0
     ServerSupportFiles=0
     VersionCompanyName="Company Name"
     VersionFileDescription="File Description"
     VersionLegalCopyright="Copyright"
     VersionLegalTrademarks="Trademark"
     VersionProductName="Product Name"
     VersionComments="Comments"
     CompilationType=0
     OptimizationType=0
     FavorPentiumPro(tm)=0
     CodeViewDebugInfo=0
     NoAliasing=0
     BoundsCheck=0
     OverflowCheck=0
     FlPointCheck=0
     FDIVCheck=0
     UnroundedFP=0
     CondComp=""
     ResFile32=""
     IconForm=""
     Startup=!(None)!
     HelpFile=""
     Title="Project1"

     [MS Transaction Server]
     AutoRefresh=1
"#;

        let project_source_file = match SourceFile::decode("project1.vbp", input.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                panic!("{}", e.print_to_string().unwrap());
            }
        };

        let result = ProjectFile::parse(&project_source_file);

        if result.has_failures() {
            for failure in result.failures() {
                failure.print();
            }

            panic!("Project parse had failures");
        }

        let project = result.unwrap();

        match project.properties.compilation_type {
            CompilationType::PCode => {}
            CompilationType::NativeCode(val) => {
                println!("{:?}", val.pentium_fdiv_bug_check);
            }
        }

        assert_eq!(project.project_type, CompileTargetType::Exe);
        assert_eq!(project.references.len(), 1);
        assert_eq!(project.objects.len(), 1);
        assert_eq!(project.modules.len(), 1);
        assert_eq!(project.classes.len(), 1);
        assert_eq!(project.designers.len(), 0);
        assert_eq!(project.forms.len(), 2);
        assert_eq!(project.user_controls.len(), 1);
        assert_eq!(project.user_documents.len(), 1);
        assert_eq!(
            project.properties.upgrade_controls,
            UpgradeControls::Upgrade
        );
        assert_eq!(project.properties.res_file_32_path, "");
        assert_eq!(project.properties.icon_form, "");
        assert_eq!(project.properties.startup, "");
        assert_eq!(project.properties.help_file_path, "");
        assert_eq!(project.properties.title, "Project1");
        assert_eq!(project.properties.exe_32_file_name, "Project1.exe");
        assert_eq!(project.properties.exe_32_compatible, "");
        assert_eq!(project.properties.command_line_arguments, "");
        assert_eq!(project.properties.path_32, "");
        assert_eq!(project.properties.name, "Project1");
        assert_eq!(project.properties.help_context_id, "0");
        assert_eq!(
            project.properties.compatibility_mode,
            CompatibilityMode::NoCompatibility
        );
        assert_eq!(project.properties.version_info.major, 1);
        assert_eq!(project.properties.version_info.minor, 0);
        assert_eq!(project.properties.version_info.revision, 0);
        assert_eq!(project.properties.version_info.auto_increment_revision, 0);
        assert_eq!(project.properties.version_info.company_name, "Company Name");
        assert_eq!(
            project.properties.version_info.file_description,
            "File Description"
        );
        assert_eq!(project.properties.version_info.trademark, "Trademark");
        assert_eq!(project.properties.version_info.product_name, "Product Name");
        assert_eq!(project.properties.version_info.comments, "Comments");
        assert_eq!(
            project.properties.server_support_files,
            ServerSupportFiles::Local,
            "server_support_files check"
        );
        assert_eq!(project.properties.conditional_compile, "");
        assert_eq!(
            project.properties.compilation_type,
            CompilationType::NativeCode(NativeCodeSettings {
                optimization_type: OptimizationType::FavorFastCode,
                favor_pentium_pro: FavorPentiumPro::False,
                code_view_debug_info: CodeViewDebugInfo::NotCreated,
                aliasing: Aliasing::AssumeAliasing,
                bounds_check: BoundsCheck::CheckBounds,
                overflow_check: OverflowCheck::CheckOverflow,
                floating_point_check: FloatingPointErrorCheck::CheckFloatingPointError,
                pentium_fdiv_bug_check: PentiumFDivBugCheck::CheckPentiumFDivBug,
                unrounded_floating_point: UnroundedFloatingPoint::DoNotAllow,
            })
        );
        assert_eq!(project.properties.start_mode, StartMode::StandAlone);
        assert_eq!(project.properties.unattended, InteractionMode::Interactive);
        assert_eq!(project.properties.retained, Retained::UnloadOnExit);
        assert_eq!(project.properties.thread_per_object, 0);
        assert_eq!(project.properties.max_number_of_threads, 1);
        assert_eq!(
            project.properties.debug_startup_option,
            DebugStartupOption::WaitForComponentCreation,
            "debug_startup_option check"
        );
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn extra_property_sections() {
        let input = r#"Type=Exe
     Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
     Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
     Module=Module1; Module1.bas
     Class=Class1; Class1.cls
     Form=Form1.frm
     Form=Form2.frm
     UserControl=UserControl1.ctl
     UserDocument=UserDocument1.uds
     ExeName32="Project1.exe"
     Command32=""
     Path32=""
     Name="Project1"
     HelpContextID="0"
     CompatibleMode="0"
     MajorVer=1
     MinorVer=0
     RevisionVer=0
     AutoIncrementVer=0
     StartMode=0
     Unattended=0
     Retained=0
     ThreadPerObject=0
     MaxNumberOfThreads=1
     DebugStartupOption=0
     NoControlUpgrade=0
     ServerSupportFiles=0
     VersionCompanyName="Company Name"
     VersionFileDescription="File Description"
     VersionLegalCopyright="Copyright"
     VersionLegalTrademarks="Trademark"
     VersionProductName="Product Name"
     VersionComments="Comments"
     CompilationType=0
     OptimizationType=0
     FavorPentiumPro(tm)=0
     CodeViewDebugInfo=0
     NoAliasing=0
     BoundsCheck=0
     OverflowCheck=0
     FlPointCheck=0
     FDIVCheck=0
     UnroundedFP=0
     CondComp=""
     ResFile32=""
     IconForm=""
     Startup=!(None)!
     HelpFile=""
     Title="Project1"

     [MS Transaction Server]
     AutoRefresh=1
     
     [VBCompiler]
     LinkSwitches=/STACK:32180000
     Comment=Nouveaut�s :�- ajout d'options dans le menu du widget��Am�liorations :�- position de la fenetre sauvegard�e��Bugs corrig�s :�- 1.4.12 - L'erreur 383 s'est produite dans la fen�tre frmConfig de la proc�dure TimerStart_Timer � la ligne 780 : Propri�t� 'Text' en lecture seule.�- Position de la fen�tre non restaur� en cas de r�duction auto au d�marrage.
"#;

        let project_source_file =
            match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
                Ok(source_file) => source_file,
                Err(e) => {
                    e.print();
                    panic!("failed to decode project source code.");
                }
            };

        let result = ProjectFile::parse(&project_source_file);

        if result.has_failures() {
            for failure in result.failures() {
                failure.print();
            }

            panic!("Project parse had failures");
        }

        let project = result.unwrap();

        assert_eq!(project.project_type, CompileTargetType::Exe);
        assert_eq!(project.references.len(), 1);
        assert_eq!(project.objects.len(), 1);
        assert_eq!(project.modules.len(), 1);
        assert_eq!(project.classes.len(), 1);
        assert_eq!(project.designers.len(), 0);
        assert_eq!(project.forms.len(), 2);
        assert_eq!(project.user_controls.len(), 1);
        assert_eq!(project.user_documents.len(), 1);
        assert_eq!(project.other_properties.len(), 2);
        assert_eq!(
            project.properties.upgrade_controls,
            UpgradeControls::Upgrade
        );
        assert_eq!(project.properties.res_file_32_path, "");
        assert_eq!(project.properties.icon_form, "");
        assert_eq!(project.properties.startup, "");
        assert_eq!(project.properties.help_file_path, "");
        assert_eq!(project.properties.title, "Project1");
        assert_eq!(project.properties.exe_32_file_name, "Project1.exe");
        assert_eq!(project.properties.exe_32_compatible, "");
        assert_eq!(project.properties.command_line_arguments, "");
        assert_eq!(project.properties.path_32, "");
        assert_eq!(project.properties.name, "Project1");
        assert_eq!(project.properties.help_context_id, "0");
        assert_eq!(
            project.properties.compatibility_mode,
            CompatibilityMode::NoCompatibility
        );
        assert_eq!(project.properties.version_info.major, 1);
        assert_eq!(project.properties.version_info.minor, 0);
        assert_eq!(project.properties.version_info.revision, 0);
        assert_eq!(project.properties.version_info.auto_increment_revision, 0);
        assert_eq!(project.properties.version_info.company_name, "Company Name");
        assert_eq!(
            project.properties.version_info.file_description,
            "File Description"
        );
        assert_eq!(project.properties.version_info.trademark, "Trademark");
        assert_eq!(project.properties.version_info.product_name, "Product Name");
        assert_eq!(project.properties.version_info.comments, "Comments");
        assert_eq!(
            project.properties.server_support_files,
            ServerSupportFiles::Local,
            "server_support_files check"
        );
        assert_eq!(project.properties.conditional_compile, "");

        assert_eq!(
            project.properties.compilation_type,
            CompilationType::NativeCode(NativeCodeSettings {
                optimization_type: OptimizationType::FavorFastCode,
                favor_pentium_pro: FavorPentiumPro::False,
                code_view_debug_info: CodeViewDebugInfo::NotCreated,
                aliasing: Aliasing::AssumeAliasing,
                bounds_check: BoundsCheck::CheckBounds,
                overflow_check: OverflowCheck::CheckOverflow,
                floating_point_check: FloatingPointErrorCheck::CheckFloatingPointError,
                pentium_fdiv_bug_check: PentiumFDivBugCheck::CheckPentiumFDivBug,
                unrounded_floating_point: UnroundedFloatingPoint::DoNotAllow,
            })
        );
        assert_eq!(project.properties.start_mode, StartMode::StandAlone);
        assert_eq!(project.properties.unattended, InteractionMode::Interactive);
        assert_eq!(project.properties.retained, Retained::UnloadOnExit);
        assert_eq!(project.properties.thread_per_object, 0);
        assert_eq!(project.properties.max_number_of_threads, 1);
        assert_eq!(
            project.properties.debug_startup_option,
            DebugStartupOption::WaitForComponentCreation,
            "debug_startup_option check"
        );
    }
}
