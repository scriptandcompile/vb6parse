//! Defines the `ProjectFile` struct and related parsing functions for VB6 Project files.
//!
//! Handles extraction of project type, references, objects, modules, classes, forms,
//! user controls, user documents, properties, and other related information from the Project file.
//!
pub mod compilesettings;
pub mod properties;

use std::collections::HashMap;
use std::convert::TryFrom;
use std::fmt::Debug;
use std::str::FromStr;

use serde::Serialize;
use strum::{EnumMessage, IntoEnumIterator};
use uuid::Uuid;

use crate::{
    errors::{ErrorDetails, ProjectErrorKind},
    objectreference::ObjectReference,
    parseresults::ParseResult,
    parsers::project::{
        compilesettings::CompilationType,
        properties::{CompileTargetType, ProjectProperties},
    },
    sourcefile::SourceFile,
    sourcestream::{Comparator, SourceStream},
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
    pub references: Vec<ProjectReference<'a>>,
    /// The list of object references in the project.
    pub objects: Vec<ObjectReference>,
    /// The list of module references in the project.
    pub modules: Vec<ProjectModuleReference<'a>>,
    /// The list of class references in the project.
    pub classes: Vec<ProjectClassReference<'a>>,
    /// The list of related documents in the project.
    pub related_documents: Vec<&'a str>,
    /// The list of property pages in the project.
    pub property_pages: Vec<&'a str>,
    /// The list of designers in the project.
    pub designers: Vec<&'a str>,
    /// The list of forms in the project.
    pub forms: Vec<&'a str>,
    /// The list of user controls in the project.
    pub user_controls: Vec<&'a str>,
    /// The list of user documents in the project.
    pub user_documents: Vec<&'a str>,
    /// Other properties grouped by section headers.
    pub other_properties: HashMap<&'a str, HashMap<&'a str, &'a str>>,
    /// The project properties.
    pub properties: ProjectProperties<'a>,
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

/// Represents a reference to a class in a VB6 project.
#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, Hash)]
pub struct ProjectClassReference<'a> {
    /// The name of the class.
    pub name: &'a str,
    /// The path to the class file.
    pub path: &'a str,
}

/// The result type for parsing a VB6 project file.
///
/// Contains the parsed `ProjectFile` and any `ProjectErrorKind` errors encountered during parsing.
///
/// This is a type alias for `ParseResult<'a, ProjectFile<'a>, ProjectErrorKind<'a>>`.
pub type ProjectResult<'a> = ParseResult<'a, ProjectFile<'a>, ProjectErrorKind<'a>>;

impl<'a> ProjectFile<'a> {
    /// Parses a VB6 project file.
    ///
    /// # Arguments
    ///
    /// * `input` - The input to parse.
    ///
    /// # Returns
    ///
    /// A `ProjectResult` containing the parsed project and/or error(s).
    ///
    /// # Errors
    ///
    /// This function can return a collection of `ErrorDetails` if the input is not a valid VB6 project file.
    ///
    /// # Panics
    ///
    /// This function can panic if the input is not a valid VB6 project file.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::*;
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
    /// if result.has_failures() {
    ///     for failure in result.failures {
    ///         failure.print();
    ///     }
    ///     panic!("Project parse produced warnings/errors");
    /// }
    ///
    /// let project = result.unwrap();
    ///
    /// assert_eq!(project.project_type, CompileTargetType::Exe);
    /// assert_eq!(project.references.len(), 1);
    /// assert_eq!(project.objects.len(), 1);
    /// assert_eq!(project.modules.len(), 1);
    /// assert_eq!(project.classes.len(), 1);
    /// assert_eq!(project.designers.len(), 0);
    /// assert_eq!(project.forms.len(), 2);
    /// assert_eq!(project.user_controls.len(), 1);
    /// assert_eq!(project.user_documents.len(), 1);
    /// assert_eq!(project.properties.startup, "Form1");
    /// assert_eq!(project.properties.title, "Project1");
    /// assert_eq!(project.properties.exe_32_file_name, "Project1.exe");
    /// ```
    #[must_use]
    pub fn parse(source_file: &'a SourceFile) -> ProjectResult<'a> {
        let mut failures = vec![];

        let mut project = ProjectFile {
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
                compilation_type: CompilationType::NativeCode(Default::default()),
                ..Default::default()
            },
        };

        let mut input = source_file.get_source_stream();

        let mut other_property_group: Option<&str> = None;

        while !input.is_empty() {
            // We also want to skip any '[MS Transaction Server]' header lines.
            // There should only be one in the file since it's only used once,
            // but we want to be flexible in what we accept so we skip any of
            // these kinds of header lines.

            // skip empty lines.
            let _ = input.take_ascii_whitespaces();
            if input.take_newline().is_some() {
                continue;
            }

            let line_start = input.start_of_line();

            // We want to grab any section header lines like '[MS Transaction Server]'.
            // Which we will use in parsing 'other properties.'
            match parse_section_header_line(&mut input) {
                Ok(Some(section_header)) => {
                    // Looks like we are no longer parsing the standard VB6 property section
                    // Now we are parsing some third party properties.
                    if !project.other_properties.contains_key(section_header) {
                        project
                            .other_properties
                            .insert(section_header, HashMap::new());
                        other_property_group = Some(section_header);
                    }
                    continue;
                }
                Ok(None) => {
                    // Not a section header line, parse the line as a normal
                    // VB6 project property line.
                }
                Err(e) => {
                    // If we fail to parse the section header line, we will
                    // need to handle the error and continue parsing.
                    failures.push(e);
                    continue;
                }
            }

            let property_name = match parse_property_name(&mut input) {
                Ok(property_name) => property_name,
                Err(e) => {
                    failures.push(e);
                    continue;
                }
            };

            // Looks like we are no longer parsing the standard VB6 property section
            // Now we are parsing some third party properties.
            if other_property_group.is_some() {
                let property_value = match parse_property_value(&mut input, property_name) {
                    Ok(property_value) => property_value,
                    Err(e) => {
                        failures.push(e);
                        continue;
                    }
                };

                project
                    .other_properties
                    .get_mut(other_property_group.unwrap())
                    .unwrap()
                    .insert(property_name, property_value);

                continue;
            }

            match property_name {
                "Type" => match parse_converted_value(&mut input, property_name) {
                    Ok(project_type_value) => {
                        project.project_type = project_type_value;
                    }
                    Err(e) => failures.push(e),
                },
                "Designer" => match parse_path_line(&mut input, property_name) {
                    Ok(designer) => project.designers.push(designer),
                    Err(e) => failures.push(e),
                },
                "Reference" => match parse_reference(&mut input) {
                    Ok(reference) => project.references.push(reference),
                    Err(e) => failures.push(e),
                },
                "Object" => match parse_object(&mut input) {
                    Ok(object) => project.objects.push(object),
                    Err(e) => failures.push(e),
                },
                "Module" => match parse_module(&mut input) {
                    Ok(module) => project.modules.push(module),
                    Err(e) => failures.push(e),
                },
                "Class" => match parse_class(&mut input) {
                    Ok(class) => project.classes.push(class),
                    Err(e) => failures.push(e),
                },
                "RelatedDoc" => match parse_path_line(&mut input, property_name) {
                    Ok(related_document) => project.related_documents.push(related_document),
                    Err(e) => failures.push(e),
                },
                "PropertyPage" => match parse_path_line(&mut input, property_name) {
                    Ok(property_page_value) => {
                        project.property_pages.push(property_page_value);
                    }
                    Err(e) => failures.push(e),
                },
                "Form" => match parse_path_line(&mut input, property_name) {
                    Ok(form) => project.forms.push(form),
                    Err(e) => failures.push(e),
                },
                "UserControl" => match parse_path_line(&mut input, property_name) {
                    Ok(user_control) => project.user_controls.push(user_control),
                    Err(e) => failures.push(e),
                },
                "UserDocument" => match parse_path_line(&mut input, property_name) {
                    Ok(user_document) => project.user_documents.push(user_document),
                    Err(e) => failures.push(e),
                },
                "ResFile32" => match parse_quoted_value(&mut input, property_name) {
                    Ok(res_32_file) => project.properties.res_file_32_path = res_32_file,
                    Err(e) => failures.push(e),
                },
                "IconForm" => match parse_quoted_value(&mut input, property_name) {
                    Ok(icon_form_value) => project.properties.icon_form = icon_form_value,
                    Err(e) => failures.push(e),
                },
                "Startup" => match parse_optional_quoted_value(&mut input, property_name) {
                    Ok(startup_value) => project.properties.startup = startup_value,
                    Err(e) => failures.push(e),
                },
                "HelpFile" => match parse_quoted_value(&mut input, property_name) {
                    Ok(help_file) => project.properties.help_file_path = help_file,
                    Err(e) => failures.push(e),
                },
                "Title" => match parse_quoted_value(&mut input, property_name) {
                    Ok(title_value) => project.properties.title = title_value,
                    Err(e) => failures.push(e),
                },
                "ExeName32" => match parse_quoted_value(&mut input, property_name) {
                    Ok(exe_32_file_name_value) => {
                        project.properties.exe_32_file_name = exe_32_file_name_value;
                    }
                    Err(e) => failures.push(e),
                },
                "Path32" => match parse_quoted_value(&mut input, property_name) {
                    Ok(path_32_value) => project.properties.path_32 = path_32_value,
                    Err(e) => failures.push(e),
                },
                "Command32" => match parse_optional_quoted_value(&mut input, property_name) {
                    Ok(command_line_arguments_value) => {
                        project.properties.command_line_arguments = command_line_arguments_value;
                    }
                    Err(e) => failures.push(e),
                },
                "Name" => match parse_optional_quoted_value(&mut input, property_name) {
                    Ok(name_value) => project.properties.name = name_value,
                    Err(e) => failures.push(e),
                },
                "Description" => match parse_quoted_value(&mut input, property_name) {
                    Ok(description_value) => project.properties.description = description_value,
                    Err(e) => failures.push(e),
                },
                "HelpContextID" => match parse_quoted_value(&mut input, property_name) {
                    Ok(help_context_id_value) => {
                        project.properties.help_context_id = help_context_id_value;
                    }
                    Err(e) => failures.push(e),
                },
                "CompatibleMode" => match parse_quoted_converted_value(&mut input, property_name) {
                    Ok(compatibility_mode_value) => {
                        project.properties.compatibility_mode = compatibility_mode_value;
                    }
                    Err(e) => failures.push(e),
                },
                "VersionCompatible32" => match parse_quoted_value(&mut input, property_name) {
                    Ok(version_32_compatibility_value) => {
                        project.properties.version_32_compatibility =
                            version_32_compatibility_value;
                    }
                    Err(e) => failures.push(e),
                },
                "CompatibleEXE32" => match parse_quoted_value(&mut input, property_name) {
                    Ok(exe_32_compatible_value) => {
                        project.properties.exe_32_compatible = exe_32_compatible_value;
                    }
                    Err(e) => failures.push(e),
                },
                "DllBaseAddress" => match parse_dll_base_address(&mut input) {
                    Ok(dll_base_address_value) => {
                        project.properties.dll_base_address = dll_base_address_value;
                    }
                    Err(e) => failures.push(e),
                },
                "RemoveUnusedControlInfo" => {
                    match parse_converted_value(&mut input, property_name) {
                        Ok(unused_control_info_value) => {
                            project.properties.unused_control_info = unused_control_info_value;
                        }
                        Err(e) => failures.push(e),
                    }
                }
                "MajorVer" => match parse_numeric(&mut input, property_name) {
                    Ok(major_value) => project.properties.version_info.major = major_value,
                    Err(e) => failures.push(e),
                },
                "MinorVer" => match parse_numeric(&mut input, property_name) {
                    Ok(minor_value) => project.properties.version_info.minor = minor_value,
                    Err(e) => failures.push(e),
                },
                "RevisionVer" => match parse_numeric(&mut input, property_name) {
                    Ok(revision_value) => project.properties.version_info.revision = revision_value,
                    Err(e) => failures.push(e),
                },
                "ThreadingModel" => match parse_converted_value(&mut input, property_name) {
                    Ok(threading_model_value) => {
                        project.properties.threading_model = threading_model_value;
                    }
                    Err(e) => failures.push(e),
                },
                "AutoIncrementVer" => match parse_numeric(&mut input, property_name) {
                    Ok(auto_increment_revision_value) => {
                        project.properties.version_info.auto_increment_revision =
                            auto_increment_revision_value;
                    }
                    Err(e) => failures.push(e),
                },
                "DebugStartupComponent" => match parse_path_line(&mut input, property_name) {
                    Ok(debug_startup_component_value) => {
                        project.properties.debug_startup_component = debug_startup_component_value;
                    }
                    Err(e) => failures.push(e),
                },
                "NoControlUpgrade" => match parse_converted_value(&mut input, property_name) {
                    Ok(upgrade_controls_value) => {
                        project.properties.upgrade_controls = upgrade_controls_value;
                    }
                    Err(e) => failures.push(e),
                },
                "ServerSupportFiles" => match parse_converted_value(&mut input, property_name) {
                    Ok(server_support_files_value) => {
                        project.properties.server_support_files = server_support_files_value;
                    }
                    Err(e) => failures.push(e),
                },
                "VersionCompanyName" => match parse_quoted_value(&mut input, property_name) {
                    Ok(company_name_value) => {
                        project.properties.version_info.company_name = company_name_value;
                    }
                    Err(e) => failures.push(e),
                },
                "VersionFileDescription" => match parse_quoted_value(&mut input, property_name) {
                    Ok(file_description_value) => {
                        project.properties.version_info.file_description = file_description_value;
                    }
                    Err(e) => failures.push(e),
                },
                "VersionLegalCopyright" => match parse_quoted_value(&mut input, property_name) {
                    Ok(copyright_value) => {
                        project.properties.version_info.copyright = copyright_value;
                    }
                    Err(e) => failures.push(e),
                },
                "VersionLegalTrademarks" => match parse_quoted_value(&mut input, property_name) {
                    Ok(trademark_value) => {
                        project.properties.version_info.trademark = trademark_value;
                    }
                    Err(e) => failures.push(e),
                },
                "VersionProductName" => match parse_quoted_value(&mut input, property_name) {
                    Ok(product_name_value) => {
                        project.properties.version_info.product_name = product_name_value;
                    }
                    Err(e) => failures.push(e),
                },
                "VersionComments" => match parse_quoted_value(&mut input, property_name) {
                    Ok(comments_value) => project.properties.version_info.comments = comments_value,
                    Err(e) => failures.push(e),
                },
                "CondComp" => match parse_quoted_value(&mut input, property_name) {
                    Ok(conditional_compile_value) => {
                        project.properties.conditional_compile = conditional_compile_value;
                    }
                    Err(e) => failures.push(e),
                },
                "CompilationType" => match parse_numeric(&mut input, property_name) {
                    Ok(compilation_type) => project.properties.compilation_type = compilation_type,
                    Err(e) => failures.push(e),
                },
                "OptimizationType" => match parse_converted_value(&mut input, property_name) {
                    Ok(optimization_type_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_optimization_type(optimization_type_value);
                    }
                    Err(e) => failures.push(e),
                },
                "FavorPentiumPro(tm)" => match parse_converted_value(&mut input, property_name) {
                    Ok(favor_pentium_pro_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_favor_pentium_pro(favor_pentium_pro_value);
                    }
                    Err(e) => failures.push(e),
                },
                "CodeViewDebugInfo" => match parse_converted_value(&mut input, property_name) {
                    Ok(code_view_debug_info_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_code_view_debug_info(code_view_debug_info_value);
                    }
                    Err(e) => failures.push(e),
                },
                "NoAliasing" => match parse_converted_value(&mut input, property_name) {
                    Ok(aliasing_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_aliasing(aliasing_value);
                    }
                    Err(e) => failures.push(e),
                },
                "BoundsCheck" => match parse_converted_value(&mut input, property_name) {
                    Ok(bounds_check_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_bounds_check(bounds_check_value);
                    }
                    Err(e) => failures.push(e),
                },
                "OverflowCheck" => match parse_converted_value(&mut input, property_name) {
                    Ok(overflow_check_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_overflow_check(overflow_check_value);
                    }
                    Err(e) => failures.push(e),
                },
                "FlPointCheck" => match parse_converted_value(&mut input, property_name) {
                    Ok(floating_point_check_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_floating_point_check(floating_point_check_value);
                    }
                    Err(e) => failures.push(e),
                },
                "FDIVCheck" => match parse_converted_value(&mut input, property_name) {
                    Ok(pentium_fdiv_bug_check_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_pentium_fdiv_bug_check(pentium_fdiv_bug_check_value);
                    }
                    Err(e) => failures.push(e),
                },
                "UnroundedFP" => match parse_converted_value(&mut input, property_name) {
                    Ok(unrounded_floating_point_value) => {
                        project.properties.compilation_type = project
                            .properties
                            .compilation_type
                            .update_unrounded_floating_point(unrounded_floating_point_value);
                    }
                    Err(e) => failures.push(e),
                },
                "StartMode" => match parse_converted_value(&mut input, property_name) {
                    Ok(start_mode_value) => project.properties.start_mode = start_mode_value,
                    Err(e) => failures.push(e),
                },
                "Unattended" => match parse_converted_value(&mut input, property_name) {
                    Ok(unattended_value) => project.properties.unattended = unattended_value,
                    Err(e) => failures.push(e),
                },
                "Retained" => match parse_converted_value(&mut input, property_name) {
                    Ok(retained_value) => project.properties.retained = retained_value,
                    Err(e) => failures.push(e),
                },
                "ThreadPerObject" => match parse_numeric::<i16>(&mut input, property_name) {
                    Ok(thread_per_object_value) => {
                        if thread_per_object_value <= 0 {
                            project.properties.thread_per_object = 0;
                        } else {
                            project.properties.thread_per_object = thread_per_object_value as u16;
                        }
                    }
                    Err(e) => failures.push(e),
                },
                "MaxNumberOfThreads" => match parse_numeric(&mut input, property_name) {
                    Ok(max_number_of_threads_value) => {
                        project.properties.max_number_of_threads = max_number_of_threads_value;
                    }
                    Err(e) => failures.push(e),
                },
                "DebugStartupOption" => match parse_converted_value(&mut input, property_name) {
                    Ok(debug_startup_option_value) => {
                        project.properties.debug_startup_option = debug_startup_option_value;
                    }
                    Err(e) => failures.push(e),
                },
                "UseExistingBrowser" => match parse_converted_value(&mut input, property_name) {
                    Ok(use_existing_browser_value) => {
                        project.properties.use_existing_browser = use_existing_browser_value;
                    }
                    Err(e) => failures.push(e),
                },
                _ => {
                    // Unknown property, skip it.
                    input.forward_to_next_line();

                    let e = input.generate_error_at(
                        line_start,
                        ProjectErrorKind::ParameterLineUnknown {
                            parameter_line_name: property_name,
                        },
                    );
                    failures.push(e);
                }
            }
        }

        ParseResult {
            result: Some(project),
            failures,
        }
    }

    /// Gets a collection of mutable references to all sub-project references in the project.
    ///
    /// # Returns
    ///
    /// A vector of mutable references to all sub-project references.
    ///
    #[must_use]
    pub fn with_subproject_references_mut(&mut self) -> Vec<&ProjectReference<'a>> {
        self.references
            .iter()
            .filter(|reference| matches!(reference, ProjectReference::SubProject { .. }))
            .collect::<Vec<_>>()
    }

    /// Gets a collection of references to all compiled references in the project.
    ///
    /// # Returns
    ///
    /// A vector of references to all compiled references.
    ///
    #[must_use]
    pub fn get_compiled_references(&self) -> Vec<&ProjectReference<'a>> {
        self.references
            .iter()
            .filter(|reference| matches!(reference, ProjectReference::Compiled { .. }))
            .collect::<Vec<_>>()
    }
}

fn parse_section_header_line<'a>(
    input: &mut SourceStream<'a>,
) -> Result<Option<&'a str>, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    // We want to grab any section header lines like '[MS Transaction Server]'.
    // Which we will use in parsing 'other properties.'
    let header_start = input.take("[", Comparator::CaseSensitive);

    if header_start.is_none() {
        // No section header line, so we can continue parsing.
        return Ok(None);
    }

    // We have a section header line.
    let Some((other_property, _)) = input.take_until("]", Comparator::CaseSensitive) else {
        // We have a section header line but it is not terminated properly.
        let fail = input.generate_error(ProjectErrorKind::UnterminatedSectionHeader);
        input.forward_to_next_line();

        return Err(fail);
    };

    let _ = input.take("]", Comparator::CaseSensitive);
    input.forward_to_next_line();

    Ok(Some(other_property))
}

fn parse_property_name<'a>(
    input: &mut SourceStream<'a>,
) -> Result<&'a str, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    let line_start = input.start_of_line();

    // We want to grab the property name.
    let property_name = input.take_until("=", Comparator::CaseSensitive);

    match property_name {
        None => {
            // No property name found, so we can't parse this line.
            // Go to the next line and return the error.
            let fail = input.generate_error_at(line_start, ProjectErrorKind::PropertyNameNotFound);
            input.forward_to_next_line();

            Err(fail)
        }
        Some((property_name, _)) => {
            // We only need the property name not the split on '=' value so we only
            // return the first of the pair in the line split.
            let _ = input.take("=", Comparator::CaseSensitive);

            Ok(property_name)
        }
    }
}

fn parse_property_value<'a>(
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Result<&'a str, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    // An line starts with the line_type followed by '=', and a value.
    //
    // By this point in the parse the line_type and "=" component should be
    // stripped off since that is how we knew to use this particular parse.;
    let parameter_start = input.offset();

    let Some((parameter_value, _)) = input.take_until_newline() else {
        // No parameter value found, so we can't parse this line.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueNotFound {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    };

    if parameter_value.is_empty() {
        // No parameter value found, so we can't parse this line.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueNotFound {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    Ok(parameter_value)
}

fn parse_quoted_value<'a>(
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Result<&'a str, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    // An line starts with the line_type followed by '=', and a quoted value.
    //
    // By this point in the parse the line_type and "=" component should be
    // stripped off since that is how we knew to use this particular parse.
    let parameter_start = input.offset();

    let Some((parameter_value, _)) = input.take_until_newline() else {
        // No parameter value found, so we can't parse this line.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueNotFound {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    };

    if parameter_value.is_empty() {
        // No startup value found, so we can't parse this line.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueNotFound {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    let start_quote_found = parameter_value.starts_with('"');
    let end_quote_found = parameter_value.ends_with('"');

    if !start_quote_found && end_quote_found {
        // The value ends with a quote but does not start with one.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueMissingOpeningQuote {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    // we have to check the length like this because if we have only a single
    // quote, then obviously the string both starts and ends with a quote (even
    // if that is the same character!) which means we still need to report this
    // failure case.
    if (start_quote_found && !end_quote_found)
        || (start_quote_found && end_quote_found && parameter_value.len() == 1)
    {
        // The value starts with a quote but does not end with one.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start + parameter_value.len(),
            ProjectErrorKind::ParameterValueMissingMatchingQuote {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    if !start_quote_found && !end_quote_found {
        // The startup value does not start or end with a quote.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueMissingQuotes {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    let parameter_value = &parameter_value[1..parameter_value.len() - 1];

    Ok(parameter_value)
}

fn parse_optional_quoted_value<'a>(
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Result<&'a str, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    // An optional line starts with 'Startup=' (or another such option starting line)
    // and is followed by the quoted value, "!None!", or "(None)", or "!(None)!" to indicate the
    // parameter value is 'None'.
    //
    // By this point in the parse the "Startup=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let parameter_start = input.offset();

    let Some((parameter_value, _)) = input.take_until_newline() else {
        // No parameter value found, so we can't parse this line.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueNotFound {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    };

    if parameter_value.is_empty() {
        // No parameter value found, so we can't parse this line.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueNotFound {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    if parameter_value == "\"(None)\""
        || parameter_value == "\"!None!\""
        || parameter_value == "\"!(None)!\""
        || parameter_value == "(None)"
        || parameter_value == "!None!"
        || parameter_value == "!(None)!"
    {
        // The parameter has a value of None.
        return Ok("");
    }

    let start_quote_found = parameter_value.starts_with('"');
    let end_quote_found = parameter_value.ends_with('"');

    if !start_quote_found && end_quote_found {
        // The value ends with a quote but does not start with one.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueMissingOpeningQuote {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    // we have to check the length like this because if we have only a single
    // quote, then obviously the string both starts and ends with a quote (even
    // if that is the same character!) which means we still need to report this
    // failure case.
    if (start_quote_found && !end_quote_found)
        || (start_quote_found && end_quote_found && parameter_value.len() == 1)
    {
        // The value starts with a quote but does not end with one.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start + parameter_value.len(),
            ProjectErrorKind::ParameterValueMissingMatchingQuote {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    if !start_quote_found && !end_quote_found {
        // The parameter value does not start or end with a quote.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueMissingQuotes {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    let parameter_value = &parameter_value[1..parameter_value.len() - 1];
    Ok(parameter_value)
}

fn parse_quoted_converted_value<'a, T>(
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Result<T, ErrorDetails<'a, ProjectErrorKind<'a>>>
where
    T: TryFrom<&'a str, Error = String> + 'a + IntoEnumIterator + EnumMessage + Debug,
{
    // This function is used to parse a quoted value that is expected to be
    // converted into an enum value through TryFrom.
    // This kind of line starts with the line_type followed by '=', and a
    // quoted value.
    let parameter_start = input.offset();

    let text_to_newline = input.take_until_newline();

    let parameter_value = match text_to_newline {
        None => {
            // No type text found, so we can't parse this line.
            // Go to the next line and return the error.
            let fail = input.generate_error_at(
                parameter_start,
                ProjectErrorKind::ParameterValueMissingOpeningQuote {
                    parameter_line_name: line_type,
                },
            );
            return Err(fail);
        }
        Some((parameter_value, _)) => parameter_value,
    };

    let start_quote_found = parameter_value.starts_with('"');
    let end_quote_found = parameter_value.ends_with('"');

    if !start_quote_found && end_quote_found {
        // The value ends with a quote but does not start with one.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueMissingOpeningQuote {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    // we have to check the length like this because if we have only a single
    // quote, then obviously the string both starts and ends with a quote (even
    // if that is the same character!) which means we still need to report this
    // failure case.
    if start_quote_found && !end_quote_found || start_quote_found && parameter_value.len() == 1 {
        // The value starts with a quote but does not end with one.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start + parameter_value.len(),
            ProjectErrorKind::ParameterValueMissingMatchingQuote {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    if !start_quote_found && !end_quote_found {
        // The value does not start or end with a quote.
        // This is an error, so we return an error.
        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueMissingQuotes {
                parameter_line_name: line_type,
            },
        );
        return Err(fail);
    }

    // trim off the quote characters.
    let parameter_value = &parameter_value[1..parameter_value.len() - 1];

    let Ok(value) = T::try_from(parameter_value) else {
        // We have a parameter value that is invalid, so we return an error.
        let valid_value_message = T::iter()
            .map(|v| format!("'{:?}' ({:#?})", v, v.get_message()))
            .collect::<Vec<_>>()
            .join(", ");

        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueInvalid {
                parameter_line_name: line_type,
                invalid_value: parameter_value,
                valid_value_message,
            },
        );
        return Err(fail);
    };

    Ok(value)
}

fn parse_converted_value<'a, T>(
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Result<T, ErrorDetails<'a, ProjectErrorKind<'a>>>
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
            let fail = input.generate_error_at(
                parameter_start,
                ProjectErrorKind::ParameterValueMissingOpeningQuote {
                    parameter_line_name: line_type,
                },
            );
            return Err(fail);
        }
        Some((parameter_value, _)) => parameter_value,
    };

    let Ok(value) = T::try_from(parameter_value) else {
        // We have a parameter value that is invalid, so we return an error.

        let valid_value_message = T::iter()
            .map(|v| format!("'{:?}' ({})", v, v.get_message().unwrap()))
            .collect::<Vec<_>>()
            .join(", ");

        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueInvalid {
                parameter_line_name: line_type,
                invalid_value: parameter_value,
                valid_value_message,
            },
        );
        return Err(fail);
    };

    Ok(value)
}

fn parse_numeric<'a, T>(
    input: &mut SourceStream<'a>,
    line_type: &'a str,
) -> Result<T, ErrorDetails<'a, ProjectErrorKind<'a>>>
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
            let fail = input.generate_error_at(
                parameter_start,
                ProjectErrorKind::ParameterValueMissingOpeningQuote {
                    parameter_line_name: line_type,
                },
            );
            return Err(fail);
        }
        Some((parameter_value, _)) => parameter_value,
    };

    let Ok(value) = parameter_value.parse::<T>() else {
        // We have a parameter value that is invalid, so we return an error.
        let valid_value_message = format!(
            "Failed attempting to parse as {0}. '{parameter_value}' is not a valid {0}",
            std::any::type_name::<T>()
        );

        let fail = input.generate_error_at(
            parameter_start,
            ProjectErrorKind::ParameterValueInvalid {
                parameter_line_name: line_type,
                invalid_value: parameter_value,
                valid_value_message,
            },
        );
        return Err(fail);
    };

    Ok(value)
}

fn parse_reference<'a>(
    input: &mut SourceStream<'a>,
) -> Result<ProjectReference<'a>, ErrorDetails<'a, ProjectErrorKind<'a>>> {
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
        return parse_compiled_reference(input);
    }

    // This is a project reference, but not a compiled reference.
    let Some((path, _)) = input.take_until_newline() else {
        // No path found, so we can't parse this line.
        let fail = input.generate_error_at(
            reference_start,
            ProjectErrorKind::ReferenceProjectPathNotFound,
        );
        return Err(fail);
    };

    if path.is_empty() {
        // No path found, so we can't parse this line.
        let fail = input.generate_error_at(
            reference_start,
            ProjectErrorKind::ReferenceProjectPathNotFound,
        );
        return Err(fail);
    }

    if !path.starts_with("*\\A") {
        // The path does not start with "*\A", which is not allowed.
        let fail = input.generate_error_at(
            reference_start,
            ProjectErrorKind::ReferenceProjectPathInvalid { value: path },
        );
        return Err(fail);
    }

    let path = &path[3..]; // Strip off the "*\A" prefix

    Ok(ProjectReference::SubProject { path })
}

fn parse_compiled_reference<'a>(
    input: &mut SourceStream<'a>,
) -> Result<ProjectReference<'a>, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    // A compiled reference starts with "*\\G{" and is followed by a UUID.
    // We have already checked that the input starts with "*\\G{".
    // By this point in the parse the "*\\G{" component should be stripped off.
    let uuid_start = input.offset();

    // This is a compiled reference.
    let Some((uuid_text, _)) = input.take_until("}#", Comparator::CaseSensitive) else {
        // No UUID found, so we can't parse this line.
        let fail = input.generate_error_at(
            uuid_start,
            ProjectErrorKind::ReferenceCompiledUuidMissingMatchingBrace,
        );
        input.forward_to_next_line();

        return Err(fail);
    };

    let uuid = if let Ok(uuid) = Uuid::parse_str(uuid_text) {
        uuid
    } else {
        // The UUID is not a valid UUID, so we can't parse this line.
        let fail =
            input.generate_error_at(uuid_start, ProjectErrorKind::ReferenceCompiledUuidInvalid);
        input.forward_to_next_line();

        return Err(fail);
    };

    let _ = input.take("}#", Comparator::CaseSensitive);
    let unknown1_start = input.offset();

    let Some((unknown1, _)) = input.take_until("#", Comparator::CaseSensitive) else {
        // No unknown1 found, so we can't parse this line.
        let fail = input.generate_error_at(
            unknown1_start,
            ProjectErrorKind::ReferenceCompiledUnknown1Missing,
        );
        input.forward_to_next_line();

        return Err(fail);
    };

    let _ = input.take("#", Comparator::CaseSensitive);
    let unknown2_start = input.offset();

    let Some((unknown2, _)) = input.take_until("#", Comparator::CaseSensitive) else {
        // No unknown2 found, so we can't parse this line.
        let fail = input.generate_error_at(
            unknown2_start,
            ProjectErrorKind::ReferenceCompiledUnknown2Missing,
        );
        input.forward_to_next_line();

        return Err(fail);
    };

    let _ = input.take("#", Comparator::CaseSensitive);
    let path_start = input.offset();

    let Some((path, _)) = input.take_until("#", Comparator::CaseSensitive) else {
        // No path found, so we can't parse this line.
        let fail =
            input.generate_error_at(path_start, ProjectErrorKind::ReferenceCompiledPathNotFound);
        input.forward_to_next_line();

        return Err(fail);
    };

    let _ = input.take("#", Comparator::CaseSensitive);
    let description_start = input.offset();

    let Some((description, _)) = input.take_until_newline() else {
        // No description found, so we can't parse this line.
        let fail = input.generate_error_at(
            description_start,
            ProjectErrorKind::ReferenceCompiledDescriptionNotFound,
        );
        return Err(fail);
    };

    if description.contains('#') {
        // The description contains a '#', which is not allowed.
        let fail = input.generate_error_at(
            description_start,
            ProjectErrorKind::ReferenceCompiledDescriptionInvalid,
        );
        return Err(fail);
    }

    // We have a compiled reference.
    let reference = ProjectReference::Compiled {
        uuid,
        unknown1,
        unknown2,
        path,
        description,
    };

    Ok(reference)
}

fn parse_object<'a>(
    input: &mut SourceStream<'a>,
) -> Result<ObjectReference, ErrorDetails<'a, ProjectErrorKind<'a>>> {
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
            let fail = input.generate_error_at(
                object_path_start,
                ProjectErrorKind::ObjectProjectPathNotFound,
            );
            input.forward_to_next_line();

            return Err(fail);
        };
        input.forward_to_next_line();

        return Ok(ObjectReference::Project { path: path.into() });
    }

    // It looks like we have a compiled object line instead. Hopefully.
    if input.peek(1) != Some("{") {
        // We do not have a compiled object line, so we can't parse this line.
        let fail = input.generate_error_at(
            object_start,
            ProjectErrorKind::ObjectCompiledMissingOpeningBrace,
        );
        input.forward_to_next_line();

        return Err(fail);
    }
    let _ = input.take("{", Comparator::CaseSensitive);

    let Some((uuid_text, _)) = input.take_until("}", Comparator::CaseSensitive) else {
        // No UUID found, so we can't parse this line.
        let fail = input.generate_error_at(
            object_start,
            ProjectErrorKind::ObjectCompiledUuidMissingMatchingBrace,
        );
        input.forward_to_next_line();

        return Err(fail);
    };

    let _ = input.take("}", Comparator::CaseSensitive);

    let uuid = if let Ok(uuid) = Uuid::parse_str(uuid_text) {
        uuid
    } else {
        // The UUID is not a valid UUID, so we can't parse this line.
        let fail =
            input.generate_error_at(object_start, ProjectErrorKind::ObjectCompiledUuidInvalid);
        input.forward_to_next_line();

        return Err(fail);
    };
    let _ = input.take("#", Comparator::CaseSensitive);

    let version_start = input.offset();
    let Some((version, invalid_version_character)) = input.take_until_not(
        &["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."],
        Comparator::CaseSensitive,
    ) else {
        // No version found, so we can't parse this line.
        let fail = input.generate_error_at(
            version_start,
            ProjectErrorKind::ObjectCompiledVersionMissing,
        );
        input.forward_to_next_line();

        return Err(fail);
    };

    if invalid_version_character != "#" {
        // The version contains an invalid character, so we can't parse this line.
        let fail = input.generate_error_at(
            version_start + version.len(),
            ProjectErrorKind::ObjectCompiledVersionInvalid,
        );
        input.forward_to_next_line();

        return Err(fail);
    }
    let _ = input.take("#", Comparator::CaseSensitive);
    let unknown1_start = input.offset();

    let Some((unknown1, _)) = input.take_until("; ", Comparator::CaseSensitive) else {
        // No unknown1 found, so we can't parse this line.
        let fail = input.generate_error_at(
            unknown1_start,
            ProjectErrorKind::ObjectCompiledUnknown1Missing,
        );
        input.forward_to_next_line();

        return Err(fail);
    };
    let _ = input.take("; ", Comparator::CaseSensitive);
    let file_name_start = input.offset();

    let file_name = input.take_until_newline();
    match file_name {
        None => {
            // No file name found, so we can't parse this line.
            let fail = input.generate_error_at(
                file_name_start,
                ProjectErrorKind::ObjectCompiledFileNameNotFound,
            );
            Err(fail)
        }
        Some((file_name, _)) => {
            if file_name.is_empty() {
                // No file name found, so we can't parse this line.
                let fail = input.generate_error_at(
                    file_name_start,
                    ProjectErrorKind::ObjectCompiledFileNameNotFound,
                );
                return Err(fail);
            }

            Ok(ObjectReference::Compiled {
                uuid,
                version: version.into(),
                unknown1: unknown1.into(),
                file_name: file_name.into(),
            })
        }
    }
}

fn parse_module<'a>(
    input: &mut SourceStream<'a>,
) -> Result<ProjectModuleReference<'a>, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    // A Module line starts with a 'Module=' and is followed by a name and a path.
    //
    // By this point in the parse the "Module=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let module_start = input.offset();

    let Some((module_name, _)) = input.take_until("; ", Comparator::CaseSensitive) else {
        // No name found, so we can't parse this line.
        let fail = input.generate_error_at(module_start, ProjectErrorKind::ModuleNameNotFound);
        input.forward_to_next_line();

        return Err(fail);
    };
    let _ = input.take("; ", Comparator::CaseSensitive);
    let module_path_start = input.offset();

    let Some((module_path, _)) = input.take_until_newline() else {
        // No path found, so we can't parse this line.
        let fail =
            input.generate_error_at(module_path_start, ProjectErrorKind::ModuleFileNameNotFound);
        return Err(fail);
    };

    if module_path.is_empty() {
        // No path found, so we can't parse this line.
        let fail =
            input.generate_error_at(module_path_start, ProjectErrorKind::ModuleFileNameNotFound);
        return Err(fail);
    }

    let module = ProjectModuleReference {
        name: module_name,
        path: module_path,
    };
    Ok(module)
}

fn parse_class<'a>(
    input: &mut SourceStream<'a>,
) -> Result<ProjectClassReference<'a>, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    // A Class line starts with a 'Class=' and is followed by a name and a path.
    //
    // By this point in the parse the "Class=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let class_start = input.offset();

    let Some((class_name, _)) = input.take_until("; ", Comparator::CaseSensitive) else {
        // No name found, so we can't parse this line.
        let fail = input.generate_error_at(class_start, ProjectErrorKind::ClassNameNotFound);
        input.forward_to_next_line();

        return Err(fail);
    };

    let _ = input.take("; ", Comparator::CaseSensitive);
    let class_path_start = input.offset();

    let Some((class_path, _)) = input.take_until_newline() else {
        // No path found, so we can't parse this line.
        let fail =
            input.generate_error_at(class_path_start, ProjectErrorKind::ClassFileNameNotFound);
        return Err(fail);
    };

    if class_path.is_empty() {
        // No path found, so we can't parse this line.
        let fail =
            input.generate_error_at(class_path_start, ProjectErrorKind::ClassFileNameNotFound);

        return Err(fail);
    }

    let class = ProjectClassReference {
        name: class_name,
        path: class_path,
    };

    Ok(class)
}

fn parse_path_line<'a>(
    input: &mut SourceStream<'a>,
    parameter_line_name: &'a str,
) -> Result<&'a str, ErrorDetails<'a, ProjectErrorKind<'a>>> {
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
            let fail = input.generate_error_at(
                path_start,
                ProjectErrorKind::PathValueNotFound {
                    parameter_line_name,
                },
            );
            Err(fail)
        }
        Some((file_path, _)) => {
            if file_path.is_empty() {
                // No file_path text found, so we can't parse this line.
                // Go to the next line and return the error.
                let fail = input.generate_error_at(
                    path_start,
                    ProjectErrorKind::PathValueNotFound {
                        parameter_line_name,
                    },
                );
                return Err(fail);
            }

            Ok(file_path)
        }
    }
}

fn parse_dll_base_address<'a>(
    input: &mut SourceStream<'a>,
) -> Result<u32, ErrorDetails<'a, ProjectErrorKind<'a>>> {
    // A DllBaseAddress line starts with a 'DllBaseAddress=' and is followed by a
    // hexadecimal value.
    //
    // By this point in the parse the "DllBaseAddress=" component should be stripped off
    // since that is how we knew to use this particular parse.
    let dll_base_address_start = input.offset();

    let Some((base_address_hex_text, _)) = input.take_until_newline() else {
        // No base address found, so we can't parse this line.
        let fail = input.generate_error_at(
            dll_base_address_start,
            ProjectErrorKind::DllBaseAddressNotFound,
        );
        return Err(fail);
    };

    if base_address_hex_text.is_empty() {
        // The base address is empty, so we can't parse this line.
        let fail = input.generate_error_at(
            dll_base_address_start,
            ProjectErrorKind::DllBaseAddressUnparsableEmpty,
        );
        return Err(fail);
    }

    if !base_address_hex_text.starts_with("&H") {
        // The base address does not start with "&H", so we can't parse this line.
        let fail = input.generate_error_at(
            dll_base_address_start,
            ProjectErrorKind::DllBaseAddressMissingHexPrefix,
        );
        return Err(fail);
    }

    let dll_base_address_start = dll_base_address_start + 2; // Skip the "&H" prefix

    let trimmed_base_address_hex_text = base_address_hex_text.trim_start_matches("&H");

    let Ok(dll_base_address) = u32::from_str_radix(trimmed_base_address_hex_text, 16) else {
        // The base address is not a valid hexadecimal value, so we can't parse this line.
        let fail = input.generate_error_at(
            dll_base_address_start,
            ProjectErrorKind::DllBaseAddressUnparsable {
                hex_value: trimmed_base_address_hex_text,
            },
        );
        return Err(fail);
    };

    Ok(dll_base_address)
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn compatibility_mode_is_unknown() {
        use crate::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"5\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompatibilityMode, ErrorDetails<ProjectErrorKind>> =
            parse_quoted_converted_value(&mut input, &parameter_name);

        assert!(matches!(
            result.err().unwrap().kind,
            ProjectErrorKind::ParameterValueInvalid { .. }
        ));
    }

    #[test]
    fn compatibility_mode_is_no_compatibility() {
        use crate::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"0\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompatibilityMode, ErrorDetails<ProjectErrorKind>> =
            parse_quoted_converted_value(&mut input, &parameter_name);

        assert_eq!(result.unwrap(), CompatibilityMode::NoCompatibility);
    }

    #[test]
    fn compatibility_mode_is_project() {
        use crate::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"1\"\r\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompatibilityMode, ErrorDetails<ProjectErrorKind>> =
            parse_quoted_converted_value(&mut input, &parameter_name);

        assert_eq!(result.unwrap(), CompatibilityMode::Project);
    }

    #[test]
    fn compatibility_mode_is_compatible_exe() {
        use crate::project::parse_quoted_converted_value;

        let mut input = SourceStream::new("", "CompatibleMode=\"2\"\n");

        let parameter_name = input
            .take("CompatibleMode", Comparator::CaseSensitive)
            .unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompatibilityMode, ErrorDetails<ProjectErrorKind>> =
            parse_quoted_converted_value(&mut input, &parameter_name);

        assert_eq!(result.unwrap(), CompatibilityMode::CompatibleExe);
    }

    #[test]
    fn project_type_is_exe() {
        use crate::parsers::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=Exe\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompileTargetType, ErrorDetails<ProjectErrorKind>> =
            parse_converted_value(&mut input, &parameter_name);

        assert_eq!(result.unwrap(), CompileTargetType::Exe);
    }

    #[test]
    fn project_type_is_oledll() {
        use crate::parsers::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=OleDll\r\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompileTargetType, ErrorDetails<ProjectErrorKind>> =
            parse_converted_value(&mut input, &parameter_name);

        assert_eq!(result.unwrap(), CompileTargetType::OleDll);
    }

    #[test]
    fn project_type_is_control() {
        use crate::parsers::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=Control\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompileTargetType, ErrorDetails<ProjectErrorKind>> =
            parse_converted_value(&mut input, &parameter_name);

        assert_eq!(result.unwrap(), CompileTargetType::Control);
    }

    #[test]
    fn project_type_is_ole_exe() {
        use crate::parsers::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=OleExe\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompileTargetType, ErrorDetails<ProjectErrorKind>> =
            parse_converted_value(&mut input, &parameter_name);

        assert_eq!(result.unwrap(), CompileTargetType::OleExe);
    }

    #[test]
    fn project_type_is_unknown_type() {
        use crate::parsers::project::parse_converted_value;

        let mut input = SourceStream::new("", "Type=blah\r\n");

        let parameter_name = input.take("Type", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result: Result<CompileTargetType, ErrorDetails<ProjectErrorKind>> =
            parse_converted_value(&mut input, &parameter_name);

        assert!(result.is_err());

        let error = result.err().unwrap();

        assert_eq!(
            matches!(error.kind, ProjectErrorKind::ParameterValueInvalid { .. }),
            true
        );
    }

    #[test]
    fn reference_compiled_line_valid() {
        use crate::parsers::project::parse_reference;

        let mut input = SourceStream::new("", "Reference=*\\G{000440D8-E9ED-4435-A9A2-06B05387BB16}#c.0#0#..\\DBCommon\\Libs\\VbIntellisenseFix.dll#VbIntellisenseFix\r\n");

        let _ = input.take("Reference", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result = parse_reference(&mut input);

        let expected_uuid = Uuid::parse_str("000440D8-E9ED-4435-A9A2-06B05387BB16").unwrap();

        assert_eq!(input.is_empty(), true);
        let result = result.unwrap();
        assert_eq!(matches!(result, ProjectReference::Compiled { .. }), true);

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
            _ => panic!("Expected a compiled reference"),
        }
    }

    #[test]
    fn reference_subproject_line_valid() {
        use crate::parsers::project::parse_reference;

        let mut input = SourceStream::new("", "Reference=*\\Atest.vbp\r\n");

        let _ = input.take("Reference", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result = parse_reference(&mut input);

        if result.is_err() {
            for error in result.err().iter() {
                error.print();
            }
            panic!("Failed to parse reference line");
        }

        assert_eq!(input.is_empty(), true);
        assert_eq!(
            result.unwrap(),
            ProjectReference::SubProject { path: "test.vbp" }
        );
    }

    #[test]
    fn module_line_valid() {
        use crate::parsers::project::parse_module;

        let mut input = SourceStream::new("", "Module=modDBAssist; ..\\DBCommon\\DBAssist.bas\r\n");

        let _ = input.take("Module", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result = parse_module(&mut input).unwrap();

        assert_eq!(input.is_empty(), true);
        assert_eq!(result.name, "modDBAssist");
        assert_eq!(result.path, "..\\DBCommon\\DBAssist.bas");
    }

    #[test]
    fn class_line_valid() {
        use crate::parsers::project::parse_class;

        let mut input = SourceStream::new(
            "",
            "Class=CStatusBarClass; ..\\DBCommon\\CStatusBarClass.cls\r\n",
        );

        let _ = input.take("Class", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result = parse_class(&mut input).unwrap();

        assert_eq!(input.is_empty(), true);
        assert_eq!(result.name, "CStatusBarClass");
        assert_eq!(result.path, "..\\DBCommon\\CStatusBarClass.cls");
    }

    #[test]
    fn object_line_valid() {
        use crate::parsers::project::parse_object;

        let mut input = SourceStream::new(
            "",
            "Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb\r\n",
        );

        let _ = input.take("Object", Comparator::CaseSensitive).unwrap();
        let _ = input.take("=", Comparator::CaseSensitive).unwrap();

        let result = parse_object(&mut input);

        if result.is_err() {
            for error in result.err().iter() {
                error.print();
            }
            panic!("Failed to parse object line");
        }

        let object = result.unwrap();

        assert_eq!(input.is_empty(), true);
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
            _ => panic!("Expected a compiled object"),
        }
    }

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
            for failure in result.failures {
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
            r#"Type=Exe
     Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"#,
        );

        let _ = input.take_ascii_whitespaces();

        let line_type = parse_property_name(&mut input).unwrap();
        let type_result: Result<CompileTargetType, ErrorDetails<ProjectErrorKind>> =
            parse_converted_value(&mut input, line_type);

        assert!(type_result.is_ok());
        assert_eq!(type_result.unwrap(), CompileTargetType::Exe);

        let _ = input.take_ascii_whitespaces();

        let _ = parse_property_name(&mut input).unwrap();
        let reference_result = parse_reference(&mut input);

        assert!(reference_result.is_ok());
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
            for failure in result.failures {
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
            for failure in result.failures {
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
