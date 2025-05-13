use std::collections::HashMap;
use std::convert::TryFrom;
use std::str::FromStr;

use bstr::{BStr, ByteSlice};
use num_enum::TryFromPrimitive;
use serde::Serialize;
use uuid::Uuid;
use winnow::{
    ascii::{line_ending, space0},
    combinator::{alt, opt},
    error::ErrMode,
    token::{literal, take_until, take_while},
    Parser,
};

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    parsers::{
        compilesettings::{
            Aliasing, BoundsCheck, CodeViewDebugInfo, CompilationType, FavorPentiumPro,
            FloatingPointErrorCheck, OptimizationType, OverflowCheck, PentiumFDivBugCheck,
            UnroundedFloatingPoint,
        },
        header::object_parse,
        vb6stream::VB6Stream,
        VB6ObjectReference,
    },
    vb6::{line_comment_parse, take_until_line_ending, VB6Result},
};

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6Project<'a> {
    pub project_type: CompileTargetType,
    pub references: Vec<VB6ProjectReference<'a>>,
    pub objects: Vec<VB6ObjectReference<'a>>,
    pub modules: Vec<VB6ProjectModule<'a>>,
    pub classes: Vec<VB6ProjectClass<'a>>,
    pub related_documents: Vec<&'a BStr>,
    pub designers: Vec<&'a BStr>,
    pub forms: Vec<&'a BStr>,
    pub user_controls: Vec<&'a BStr>,
    pub user_documents: Vec<&'a BStr>,
    pub other_properties: HashMap<&'a BStr, HashMap<&'a BStr, &'a BStr>>,

    pub unused_control_info: UnusedControlInfo,
    pub upgrade_controls: UpgradeControls,
    pub res_file_32_path: Option<&'a BStr>,
    pub icon_form: Option<&'a BStr>,
    pub startup: Option<&'a BStr>,
    pub help_file_path: Option<&'a BStr>,
    pub title: Option<&'a BStr>,
    pub exe_32_file_name: Option<&'a BStr>,
    pub exe_32_compatible: Option<&'a BStr>,
    pub dll_base_address: u32,
    pub path_32: Option<&'a BStr>,
    pub command_line_arguments: Option<&'a BStr>,
    pub name: Option<&'a BStr>,
    pub description: Option<&'a BStr>,
    pub debug_startup_component: Option<&'a BStr>,
    // May need to be switched to a u32. Not sure yet.
    pub help_context_id: Option<&'a BStr>,
    pub compatibility_mode: CompatibilityMode,
    pub version_32_compatibility: Option<&'a BStr>,
    pub version_info: VersionInformation<'a>,
    pub server_support_files: ServerSupportFiles,
    pub conditional_compile: Option<&'a BStr>,
    pub compilation_type: CompilationType,

    pub start_mode: StartMode,
    pub unattended: InteractionMode,
    pub retained: Retained,
    pub thread_per_object: Option<u16>,
    pub threading_model: ThreadingModel,
    pub max_number_of_threads: u16,
    pub debug_startup_option: DebugStartupOption,
    pub use_existing_browser: UseExistingBrowser,
    pub property_page: Option<&'a BStr>,
}

/// Retained mode of the VB6 project.
///
/// Hints to the loading program whether the project DLL should be retained in
/// memory or unloaded when no longer in use.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum Retained {
    /// The DLL is unloaded when no longer in use.
    #[default]
    UnloadOnExit = 0,
    /// `RetainedInMemory` only indicates to the loading program that the DLL
    /// should be retained in memory, it does not guarantee that the DLL will be
    /// retained in memory. Retaining a DLL in memory comes with a memory and
    /// performance cost that the host program may not wish to sustain.
    RetainedInMemory = 1,
}

/// Indicates whether to use an existing browser instance.
///
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum UseExistingBrowser {
    /// Do not use an existing browser instance.
    DoNotUse = 0,
    /// If Internet Explorer is already running, use the existing instance.
    /// Otherwise, launch a new instance.
    #[default]
    Use = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum StartMode {
    #[default]
    StandAlone = 0,
    Automation = 1,
}

/// Interaction mode for VB6 projects.
///
/// Indicates if the project is intended to run without user interaction.
/// Unattended projects have no interface elements.
/// Any run-time functions such as messages that normally result in user
/// interaction are written to an event log.
///
/// Interactive is the default mode, where the program can show dialogs and
/// interact with the user.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum InteractionMode {
    /// The program can show dialogs and interact with the user.
    #[default]
    Interactive = 0,
    /// The program cannot show dialogs and will not interact with the user.
    Unattended = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum ServerSupportFiles {
    #[default]
    Local = 0,
    Remote = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum UpgradeControls {
    #[default]
    Upgrade = 0,
    NoUpgrade = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum UnusedControlInfo {
    Retain = 0,
    #[default]
    Remove = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum CompatibilityMode {
    NoCompatibility = 0,
    #[default]
    Project = 1,
    CompatibleExe = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum DebugStartupOption {
    #[default]
    WaitForComponentCreation = 0,
    StartComponent = 1,
    StartProgram = 2,
    StartBrowser = 3,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VersionInformation<'a> {
    pub major: u16,
    pub minor: u16,
    pub revision: u16,
    pub auto_increment_revision: u16,
    pub company_name: Option<&'a BStr>,
    pub file_description: Option<&'a BStr>,
    pub copyright: Option<&'a BStr>,
    pub trademark: Option<&'a BStr>,
    pub product_name: Option<&'a BStr>,
    pub comments: Option<&'a BStr>,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum CompileTargetType {
    Exe,
    Control,
    OleExe,
    OleDll,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum ThreadingModel {
    /// Single-threaded.
    SingleThreaded = 0,
    /// Apartment-threaded.
    #[default]
    ApartmentThreaded = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum VB6ProjectReference<'a> {
    Compiled {
        uuid: Uuid,
        unknown1: &'a BStr,
        unknown2: &'a BStr,
        path: &'a BStr,
        description: &'a BStr,
    },
    SubProject {
        path: &'a BStr,
    },
}

impl Serialize for VB6ProjectReference<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        match self {
            VB6ProjectReference::Compiled {
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
            VB6ProjectReference::SubProject { path } => {
                let mut state = serializer.serialize_struct("SubProjectReference", 1)?;

                state.serialize_field("path", path)?;

                state.end()
            }
        }
    }
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6ProjectModule<'a> {
    pub name: &'a BStr,
    pub path: &'a BStr,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6ProjectClass<'a> {
    pub name: &'a BStr,
    pub path: &'a BStr,
}

impl<'a> VB6Project<'a> {
    /// Parses a VB6 project file.
    ///
    /// # Arguments
    ///
    /// * `input` - The input to parse.
    ///
    /// # Returns
    ///
    /// A `Result` containing the parsed project or an error.
    ///
    /// # Errors
    ///
    /// This function can return a `VB6Error` if the input is not a valid VB6 project file.
    ///
    /// # Panics
    ///
    /// This function can panic if the input is not a valid VB6 project file.
    ///
    /// # Example
    ///
    /// ```rust
    /// use crate::vb6parse::VB6Project;
    /// use crate::vb6parse::parsers::CompileTargetType;
    /// use crate::vb6parse::parsers::project::CompatibilityMode;
    /// use bstr::BStr;
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
    ///
    /// [MS Transaction Server]
    /// AutoRefresh=1
    /// "#;
    ///
    /// let project = VB6Project::parse("project1.vbp", input.as_bytes()).unwrap();
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
    /// assert_eq!(project.startup, Some(BStr::new(b"Form1")));
    /// assert_eq!(project.title, Some(BStr::new(b"Project1")));
    /// assert_eq!(project.exe_32_file_name, Some(BStr::new(b"Project1.exe")));
    /// ```
    pub fn parse(file_name: impl Into<String>, source_code: &'a [u8]) -> Result<Self, VB6Error> {
        let mut input = VB6Stream::new(file_name, source_code);

        let mut references = vec![];
        let mut user_documents = vec![];
        let mut objects = vec![];
        let mut modules = vec![];
        let mut classes = vec![];
        let mut designers = vec![];
        let mut forms = vec![];
        let mut user_controls = vec![];
        let mut related_documents = vec![];
        let mut other_properties = HashMap::new();

        let mut debug_startup_component = Some(BStr::new(b""));
        let mut unused_control_info = UnusedControlInfo::Remove;
        let mut project_type: Option<CompileTargetType> = None;
        let mut res_file_32_path = Some(BStr::new(b""));
        let mut icon_form = Some(BStr::new(b""));
        let mut startup = Some(BStr::new(b""));
        let mut help_file_path = Some(BStr::new(b""));
        let mut title = Some(BStr::new(b""));
        let mut exe_32_file_name = Some(BStr::new(b""));
        let mut exe_32_compatible = Some(BStr::new(b""));
        let mut dll_base_address = 0x1100_0000_u32;
        let mut version_32_compatibility = Some(BStr::new(b""));
        let mut path_32 = Some(BStr::new(b""));
        let mut command_line_arguments = Some(BStr::new(b""));
        let mut name = Some(BStr::new(b""));
        let mut description = Some(BStr::new(b""));
        let mut help_context_id = Some(BStr::new(b""));
        let mut compatibility_mode = CompatibilityMode::Project;
        let mut threading_model = ThreadingModel::ApartmentThreaded;
        let mut upgrade_controls = UpgradeControls::Upgrade;
        let mut server_support_files = ServerSupportFiles::Local;
        let mut conditional_compile = Some(BStr::new(b""));
        let mut compilation_type = CompilationType::PCode;
        let mut compilation_type_value = -1;
        let mut optimization_type = OptimizationType::FavorFastCode;
        let mut favor_pentium_pro = FavorPentiumPro::default();
        let mut code_view_debug_info = CodeViewDebugInfo::NotCreated;
        let mut aliasing = Aliasing::AssumeAliasing;
        let mut bounds_check = BoundsCheck::CheckBounds;
        let mut overflow_check = OverflowCheck::CheckOverflow;
        let mut floating_point_check = FloatingPointErrorCheck::CheckFloatingPointError;
        let mut pentium_fdiv_bug_check = PentiumFDivBugCheck::NoPentiumFDivBugCheck;
        let mut unrounded_floating_point = UnroundedFloatingPoint::DoNotAllow;
        let mut start_mode = StartMode::StandAlone;
        let mut unattended = InteractionMode::Interactive;
        let mut retained = Retained::UnloadOnExit;
        let mut thread_per_object = None;
        let mut max_number_of_threads = 1;
        let mut debug_startup_option = DebugStartupOption::WaitForComponentCreation;
        let mut company_name = Some(BStr::new(b""));
        let mut file_description = Some(BStr::new(b""));
        let mut major = 0u16;
        let mut minor = 0u16;
        let mut revision = 0u16;
        let mut auto_increment_revision = 0;
        let mut copyright = Some(BStr::new(b""));
        let mut trademark = Some(BStr::new(b""));
        let mut product_name = Some(BStr::new(b""));
        let mut comments = Some(BStr::new(b""));
        let mut use_existing_browser = UseExistingBrowser::Use;
        let mut property_page = Some(BStr::new(b""));

        let mut other_property_group: Option<&'a BStr> = None;

        while !input.is_empty() {
            // We also want to skip any '[MS Transaction Server]' header lines.
            // There should only be one in the file since it's only used once,
            // but we want to be flexible in what we accept so we skip any of
            // these kinds of header lines.

            // skip empty lines.
            if (space0, line_ending::<_, VB6Error>)
                .parse_next(&mut input)
                .is_ok()
            {
                continue;
            }

            // We want to grab any '[MS Transaction Server]' or other section header lines.
            // Which we will use in parsing 'other properties.'
            if let Ok((_, other_property, _, _, _)) = (
                '[',
                take_until(0.., ']'),
                ']',
                space0,
                line_ending::<_, VB6Error>,
            )
                .parse_next(&mut input)
            {
                if !other_properties.contains_key(other_property) {
                    other_properties.insert(other_property, HashMap::new());

                    other_property_group = Some(other_property);
                }
                continue;
            }

            // Looks like we are no longer parsing the standard VB6 property section
            // Now we are parsing some third party properties.
            if other_property_group.is_some() {
                let (property_name, property_value) = match other_property_parse(&mut input) {
                    Ok((property_name, property_value)) => (property_name, property_value),
                    Err(e) => {
                        return Err(input.error(e.into_inner().unwrap()));
                    }
                };

                other_properties
                    .get_mut(other_property_group.unwrap())
                    .unwrap()
                    .insert(property_name, property_value);

                continue;
            }

            let _: VB6Result<_> = space0.parse_next(&mut input);

            // Type
            if literal::<_, _, VB6Error>("Type")
                .parse_next(&mut input)
                .is_ok()
            {
                project_type = match project_type_parse.parse_next(&mut input) {
                    Ok(project_type) => Some(project_type),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("Designer")
                .parse_next(&mut input)
                .is_ok()
            {
                let designer = match designer_parse.parse_next(&mut input) {
                    Ok(designer) => designer,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                designers.push(designer);

                continue;
            }

            if literal::<_, _, VB6Error>("Reference")
                .parse_next(&mut input)
                .is_ok()
            {
                let reference = match reference_parse.parse_next(&mut input) {
                    Ok(reference) => reference,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                references.push(reference);

                continue;
            }

            if literal::<_, _, VB6Error>("Object")
                .parse_next(&mut input)
                .is_ok()
            {
                let object = match object_parse.parse_next(&mut input) {
                    Ok(object) => object,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                objects.push(object);

                continue;
            }

            if literal::<_, _, VB6Error>("Module")
                .parse_next(&mut input)
                .is_ok()
            {
                let module = match module_parse.parse_next(&mut input) {
                    Ok(module) => module,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                modules.push(module);

                continue;
            }

            if literal::<_, _, VB6Error>("Class")
                .parse_next(&mut input)
                .is_ok()
            {
                let class = match class_parse.parse_next(&mut input) {
                    Ok(class) => class,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                classes.push(class);

                continue;
            }

            if literal::<_, _, VB6Error>("RelatedDoc")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, '=', space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                let Ok(related_document) = take_until_line_ending.parse_next(&mut input) else {
                    return Err(input.error(VB6ErrorKind::RelatedDocLineUnparseable));
                };

                related_documents.push(related_document);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Form")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, '=', space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                let Ok(form): VB6Result<_> = take_until_line_ending.parse_next(&mut input) else {
                    return Err(input.error(VB6ErrorKind::FormLineUnparseable));
                };

                forms.push(form);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("UserControl")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, '=', space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                let Ok(user_control): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::UserControlLineUnparseable));
                };

                user_controls.push(user_control);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("UserDocument")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, '=', space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                let Ok(user_document): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::UserDocumentLineUnparseable));
                };

                user_documents.push(user_document);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("ResFile32")
                .parse_next(&mut input)
                .is_ok()
            {
                res_file_32_path = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("IconForm")
                .parse_next(&mut input)
                .is_ok()
            {
                icon_form = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("Startup")
                .parse_next(&mut input)
                .is_ok()
            {
                // if the project lacks a startup object/function/etc it will be !(None)! or !! in the file
                // which is distinct from the double qouted way of targeting a specific object.
                startup = process_qouted_optional_parameter(
                    &mut input,
                    VB6ErrorKind::StartupUnparseable,
                )?;

                continue;
            }

            if literal::<_, _, VB6Error>("HelpFile")
                .parse_next(&mut input)
                .is_ok()
            {
                help_file_path = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("Title")
                .parse_next(&mut input)
                .is_ok()
            {
                title = match title_parse.parse_next(&mut input) {
                    Ok(title) => Some(title),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("ExeName32")
                .parse_next(&mut input)
                .is_ok()
            {
                exe_32_file_name = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("Path32")
                .parse_next(&mut input)
                .is_ok()
            {
                path_32 = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("Command32")
                .parse_next(&mut input)
                .is_ok()
            {
                // if the project lacks a commandline it will be !(None)! or !! or "" in the file
                // which is distinct from the double qouted way of targeting a specific object.
                command_line_arguments = process_qouted_optional_parameter(
                    &mut input,
                    VB6ErrorKind::CommandLineUnparseable,
                )?;

                continue;
            }

            if literal::<_, _, VB6Error>("Name")
                .parse_next(&mut input)
                .is_ok()
            {
                // if the project lacks a name it will be !(None)! or !! or "" in the file..
                name =
                    process_qouted_optional_parameter(&mut input, VB6ErrorKind::NameUnparseable)?;

                continue;
            }

            if literal::<_, _, VB6Error>("Description")
                .parse_next(&mut input)
                .is_ok()
            {
                description = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("HelpContextID")
                .parse_next(&mut input)
                .is_ok()
            {
                // if the project lacks a help_context_id it will be !(None)! or !! or "" in the file..
                help_context_id = process_qouted_optional_parameter(
                    &mut input,
                    VB6ErrorKind::HelpContextIdUnparseable,
                )?;

                continue;
            }

            if literal::<_, _, VB6Error>("CompatibleMode")
                .parse_next(&mut input)
                .is_ok()
            {
                compatibility_mode = match compatibility_mode_parse.parse_next(&mut input) {
                    Ok(compatibility_mode) => compatibility_mode,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("VersionCompatible32")
                .parse_next(&mut input)
                .is_ok()
            {
                version_32_compatibility = process_qouted_parameter(&mut input)?;
                continue;
            }

            if literal::<_, _, VB6Error>("CompatibleEXE32")
                .parse_next(&mut input)
                .is_ok()
            {
                exe_32_compatible = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("DllBaseAddress")
                .parse_next(&mut input)
                .is_ok()
            {
                dll_base_address = process_dll_base_address(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("RemoveUnusedControlInfo")
                .parse_next(&mut input)
                .is_ok()
            {
                unused_control_info =
                    process_parameter(&mut input, VB6ErrorKind::UnusedControlInfoUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("MajorVer")
                .parse_next(&mut input)
                .is_ok()
            {
                major =
                    process_numeric_parameter(&mut input, VB6ErrorKind::MajorVersionUnparseable)?;

                continue;
            }

            if literal::<_, _, VB6Error>("MinorVer")
                .parse_next(&mut input)
                .is_ok()
            {
                minor =
                    process_numeric_parameter(&mut input, VB6ErrorKind::MinorVersionUnparseable)?;

                continue;
            }

            if literal::<_, _, VB6Error>("RevisionVer")
                .parse_next(&mut input)
                .is_ok()
            {
                revision = process_numeric_parameter(
                    &mut input,
                    VB6ErrorKind::RevisionVersionUnparseable,
                )?;

                continue;
            }

            if literal::<_, _, VB6Error>("ThreadingModel")
                .parse_next(&mut input)
                .is_ok()
            {
                threading_model = process_parameter::<ThreadingModel>(
                    &mut input,
                    VB6ErrorKind::ThreadingModelInvalid,
                )?;

                continue;
            }

            if literal::<_, _, VB6Error>("AutoIncrementVer")
                .parse_next(&mut input)
                .is_ok()
            {
                auto_increment_revision =
                    process_numeric_parameter(&mut input, VB6ErrorKind::AutoIncrementUnparseable)?;

                continue;
            }

            if literal::<_, _, VB6Error>("PropertyPage")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, '=', space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                property_page = match take_until_line_ending.parse_next(&mut input) {
                    Ok(property_page) => Some(property_page),
                    Err(_) => {
                        return Err(input.error(VB6ErrorKind::PropertyPageUnparseable));
                    }
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("DebugStartupComponent")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, '=', space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                debug_startup_component = match take_until_line_ending.parse_next(&mut input) {
                    Ok(debug_startup_component) => Some(debug_startup_component),
                    Err(_) => {
                        return Err(input.error(VB6ErrorKind::CommentUnparseable));
                    }
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("NoControlUpgrade")
                .parse_next(&mut input)
                .is_ok()
            {
                upgrade_controls =
                    process_parameter(&mut input, VB6ErrorKind::NoControlUpgradeUnparsable)?;

                continue;
            }

            if literal::<_, _, VB6Error>("ServerSupportFiles")
                .parse_next(&mut input)
                .is_ok()
            {
                server_support_files =
                    process_parameter(&mut input, VB6ErrorKind::ServerSupportFilesUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("VersionCompanyName")
                .parse_next(&mut input)
                .is_ok()
            {
                company_name = process_qouted_parameter(&mut input)?;
                continue;
            }

            if literal::<_, _, VB6Error>("VersionFileDescription")
                .parse_next(&mut input)
                .is_ok()
            {
                file_description = process_qouted_parameter(&mut input)?;
                continue;
            }

            if literal::<_, _, VB6Error>("VersionLegalCopyright")
                .parse_next(&mut input)
                .is_ok()
            {
                copyright = process_qouted_parameter(&mut input)?;
                continue;
            }

            if literal::<_, _, VB6Error>("VersionLegalTrademarks")
                .parse_next(&mut input)
                .is_ok()
            {
                trademark = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("VersionProductName")
                .parse_next(&mut input)
                .is_ok()
            {
                product_name = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("VersionComments")
                .parse_next(&mut input)
                .is_ok()
            {
                comments = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("CondComp")
                .parse_next(&mut input)
                .is_ok()
            {
                conditional_compile = process_qouted_parameter(&mut input)?;

                continue;
            }

            if literal::<_, _, VB6Error>("CompilationType")
                .parse_next(&mut input)
                .is_ok()
            {
                compilation_type_value = process_numeric_parameter(
                    &mut input,
                    VB6ErrorKind::CompilationTypeUnparseable,
                )?;

                continue;
            }

            if literal::<_, _, VB6Error>("OptimizationType")
                .parse_next(&mut input)
                .is_ok()
            {
                optimization_type = process_parameter::<OptimizationType>(
                    &mut input,
                    VB6ErrorKind::OptimizationTypeUnparseable,
                )?;

                continue;
            }

            if literal::<_, _, VB6Error>("FavorPentiumPro(tm)")
                .parse_next(&mut input)
                .is_ok()
            {
                favor_pentium_pro =
                    process_parameter(&mut input, VB6ErrorKind::FavorPentiumProUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("CodeViewDebugInfo")
                .parse_next(&mut input)
                .is_ok()
            {
                code_view_debug_info =
                    process_parameter(&mut input, VB6ErrorKind::CodeViewDebugInfoUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("NoAliasing")
                .parse_next(&mut input)
                .is_ok()
            {
                aliasing = process_parameter(&mut input, VB6ErrorKind::NoAliasingUnparseable)?;

                continue;
            }

            if literal::<_, _, VB6Error>("BoundsCheck")
                .parse_next(&mut input)
                .is_ok()
            {
                bounds_check = process_parameter(&mut input, VB6ErrorKind::BoundsCheckUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("OverflowCheck")
                .parse_next(&mut input)
                .is_ok()
            {
                overflow_check =
                    process_parameter(&mut input, VB6ErrorKind::OverflowCheckUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("FlPointCheck")
                .parse_next(&mut input)
                .is_ok()
            {
                floating_point_check =
                    process_parameter(&mut input, VB6ErrorKind::FlPointCheckUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("FDIVCheck")
                .parse_next(&mut input)
                .is_ok()
            {
                pentium_fdiv_bug_check =
                    process_parameter(&mut input, VB6ErrorKind::FDIVCheckUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("UnroundedFP")
                .parse_next(&mut input)
                .is_ok()
            {
                unrounded_floating_point =
                    process_parameter(&mut input, VB6ErrorKind::UnroundedFPUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("StartMode")
                .parse_next(&mut input)
                .is_ok()
            {
                start_mode = process_parameter(&mut input, VB6ErrorKind::StartModeUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("Unattended")
                .parse_next(&mut input)
                .is_ok()
            {
                unattended = process_parameter::<InteractionMode>(
                    &mut input,
                    VB6ErrorKind::UnattendedUnparseable,
                )?;
                continue;
            }

            if literal::<_, _, VB6Error>("Retained")
                .parse_next(&mut input)
                .is_ok()
            {
                retained =
                    process_parameter::<Retained>(&mut input, VB6ErrorKind::RetainedUnparseable)?;
                continue;
            }

            if literal::<_, _, VB6Error>("ThreadPerObject")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, '=', space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                let Ok(threads): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::ThreadPerObjectUnparseable));
                };

                if threads.trim() == b"-1" {
                    thread_per_object = None;
                } else {
                    thread_per_object = match threads.to_string().as_str().parse::<u16>() {
                        Ok(thread_per_object) => Some(thread_per_object),
                        Err(_) => return Err(input.error(VB6ErrorKind::ThreadPerObjectUnparseable)),
                    }
                }

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("MaxNumberOfThreads")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, '=', space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                let Ok(max_threads): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::MaxThreadsUnparseable));
                };

                max_number_of_threads = match max_threads.to_string().as_str().parse::<u16>() {
                    Ok(max_number_of_threads) => max_number_of_threads,
                    Err(_) => return Err(input.error(VB6ErrorKind::MaxThreadsUnparseable)),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("DebugStartupOption")
                .parse_next(&mut input)
                .is_ok()
            {
                debug_startup_option = process_parameter::<DebugStartupOption>(
                    &mut input,
                    VB6ErrorKind::DebugStartupOptionUnparseable,
                )?;

                continue;
            }

            if literal::<_, _, VB6Error>("UseExistingBrowser")
                .parse_next(&mut input)
                .is_ok()
            {
                use_existing_browser = process_parameter::<UseExistingBrowser>(
                    &mut input,
                    VB6ErrorKind::UseExistingBrowserUnparseable,
                )?;

                continue;
            }

            return Err(input.error(VB6ErrorKind::LineTypeUnknown));
        }

        if project_type.is_none() {
            return Err(input.error(VB6ErrorKind::FirstLineNotProject));
        }

        let version_info = VersionInformation {
            major,
            minor,
            revision,
            auto_increment_revision,
            company_name,
            file_description,
            copyright,
            trademark,
            product_name,
            comments,
        };

        // native code
        if compilation_type_value == 0 {
            compilation_type = CompilationType::NativeCode {
                optimization_type,
                favor_pentium_pro,
                code_view_debug_info,
                aliasing,
                bounds_check,
                overflow_check,
                floating_point_check,
                pentium_fdiv_bug_check,
                unrounded_floating_point,
            };
        }

        let project = VB6Project {
            project_type: project_type.unwrap(),
            references,
            objects,
            modules,
            classes,
            designers,
            forms,
            user_controls,
            user_documents,
            related_documents,
            other_properties,
            unused_control_info,
            upgrade_controls,
            debug_startup_component,
            res_file_32_path,
            icon_form,
            startup,
            help_file_path,
            title,
            exe_32_file_name,
            exe_32_compatible,
            dll_base_address,
            version_32_compatibility,
            path_32,
            command_line_arguments,
            name,
            description,
            help_context_id,
            compatibility_mode,
            version_info,
            server_support_files,
            conditional_compile,
            compilation_type,
            start_mode,
            unattended,
            retained,
            thread_per_object,
            max_number_of_threads,
            threading_model,
            debug_startup_option,
            use_existing_browser,
            property_page,
        };

        Ok(project)
    }

    #[must_use]
    pub fn get_subproject_references(&self) -> Vec<&VB6ProjectReference> {
        self.references
            .iter()
            .filter(|reference| matches!(reference, VB6ProjectReference::SubProject { .. }))
            .collect::<Vec<_>>()
    }

    #[must_use]
    pub fn get_compiled_references(&self) -> Vec<&VB6ProjectReference> {
        self.references
            .iter()
            .filter(|reference| matches!(reference, VB6ProjectReference::Compiled { .. }))
            .collect::<Vec<_>>()
    }
}

fn process_parameter<T>(
    input: &mut VB6Stream<'_>,
    error_on_conversion: VB6ErrorKind,
) -> Result<T, VB6Error>
where
    T: TryFrom<i16>,
{
    if (space0::<_, VB6ErrorKind>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoEqualSplit));
    }

    let Ok(result_ascii) =
        take_while::<_, _, VB6ErrorKind>(1.., ('-', '0'..='9')).parse_next(input)
    else {
        return Err(input.error(error_on_conversion));
    };

    let Ok(result) = result_ascii.to_string().as_str().parse::<i16>() else {
        return Err(input.error(error_on_conversion));
    };

    let Ok(conversion) = T::try_from(result) else {
        return Err(input.error(error_on_conversion));
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoLineEnding));
    }

    Ok(conversion)
}

fn process_numeric_parameter<F>(
    input: &mut VB6Stream,
    error_on_conversion: VB6ErrorKind,
) -> Result<F, VB6Error>
where
    F: FromStr,
{
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoEqualSplit));
    }

    let Ok(result_ascii) =
        take_while::<_, _, VB6ErrorKind>(1.., ('-', '0'..='9')).parse_next(input)
    else {
        return Err(input.error(error_on_conversion));
    };

    let Ok(value) = result_ascii.to_string().as_str().parse::<F>() else {
        return Err(input.error(error_on_conversion));
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoLineEnding));
    }

    Ok(value)
}

fn process_qouted_parameter<'a>(input: &mut VB6Stream<'a>) -> Result<Option<&'a BStr>, VB6Error> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoEqualSplit));
    }

    let value = match qouted_value_parse("\"").parse_next(input) {
        Ok(value) => Some(value),
        Err(e) => return Err(input.error(e.into_inner().unwrap())),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoLineEnding));
    }

    Ok(value)
}

fn process_qouted_optional_parameter<'a>(
    input: &mut VB6Stream<'a>,
    error_on_conversion: VB6ErrorKind,
) -> Result<Option<&'a BStr>, VB6Error> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoEqualSplit));
    }

    // if the project lacks this value it will be !(None)! or !! or "" in the file..
    let value = match alt((qouted_value_parse("\""), qouted_value_parse("!"))).parse_next(input) {
        Ok(value) => {
            // if we have !(None)! or !! then we have no value line.
            if value == "(None)" || value.is_empty() {
                None
            } else {
                Some(value)
            }
        }
        Err(_) => return Err(input.error(error_on_conversion)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoLineEnding));
    }

    Ok(value)
}

fn process_dll_base_address(input: &mut VB6Stream<'_>) -> Result<u32, VB6Error> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoEqualSplit));
    }

    let Ok(base_address_hex_text): VB6Result<_> = take_until_line_ending.parse_next(input) else {
        return Err(input.error(VB6ErrorKind::DllBaseAddressUnparseable));
    };

    let Ok(dll_base_address) = u32::from_str_radix(
        base_address_hex_text.to_string().trim_start_matches("&H"),
        16,
    ) else {
        return Err(input.error(VB6ErrorKind::DllBaseAddressUnparseable));
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(input.error(VB6ErrorKind::NoLineEnding));
    }

    Ok(dll_base_address)
}

fn other_property_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<(&'a BStr, &'a BStr)> {
    let property_name =
        match (opt(space0::<_, VB6ErrorKind>), take_until(0.., '=')).parse_next(input) {
            Ok((_, property_name)) => property_name,
            Err(e) => return Err(ErrMode::Cut(e)),
        };

    if (opt(space0::<_, VB6Error>), '=').parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    let property_value = match take_until::<_, _, VB6ErrorKind>(0.., ("\r", "\n")).parse_next(input)
    {
        Ok(property_value) => property_value,
        Err(e) => return Err(ErrMode::Cut(e)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    Ok((property_name, property_value))
}

fn designer_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    let Ok(designer) = take_until_line_ending.parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::DesignerLineUnparseable));
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    Ok(designer)
}

fn compatibility_mode_parse(input: &mut VB6Stream<'_>) -> VB6Result<CompatibilityMode> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    let compatibility_mode =
        match alt((qouted_value_parse("\""), qouted_value_parse("!"))).parse_next(input) {
            Ok(compatible_mode) => match compatible_mode.as_bytes() {
                b"0" => CompatibilityMode::NoCompatibility,
                b"1" => CompatibilityMode::Project,
                b"2" => CompatibilityMode::CompatibleExe,
                _ => {
                    return Err(ErrMode::Cut(VB6ErrorKind::CompatibilityModeUnparseable));
                }
            },
            Err(e) => return Err(ErrMode::Cut(e.into_inner().unwrap())),
        };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    Ok(compatibility_mode)
}

fn title_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    if (space0::<_, VB6Error>, '"').parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::TitleUnparseable));
    }

    // it's perfectly possible to use '"' within the title string.
    // VB6 being the language it is, there is no escape sequence for
    // this. Instead, the title is wrapped in quotes and the quotes
    // are just simply included in the text. This means we can't use
    // string_parser here.
    let Ok(title): VB6Result<_> =
        alt((take_until(1.., "\"\r\n"), take_until(1.., "\"\n"))).parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ErrorKind::TitleUnparseable));
    };

    // We need to skip the closing quote.
    // But we also need to make sure we don't skip the line ending.
    let _: VB6Result<_> = '"'.parse_next(input);

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    Ok(title)
}

fn qouted_value_parse<'a>(
    qoute_char: &'a str,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<&'a BStr> {
        literal(qoute_char).parse_next(input)?;
        let qouted_value = take_until(0.., qoute_char).parse_next(input)?;
        literal(qoute_char).parse_next(input)?;

        Ok(qouted_value)
    }
}

fn module_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ProjectModule<'a>> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    let (name, path) = semicolon_space_split_parse.parse_next(input)?;

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    let name = name.as_bstr();
    let path = path.as_bstr();

    let module = VB6ProjectModule { name, path };

    Ok(module)
}

fn class_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ProjectClass<'a>> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    let (name, path) = semicolon_space_split_parse.parse_next(input)?;

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    let name = name.as_bstr();
    let path = path.as_bstr();

    let module = VB6ProjectClass { name, path };

    Ok(module)
}

fn semicolon_space_split_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<(&'a [u8], &'a [u8])> {
    let left = take_until(1.., "; ").parse_next(input)?;

    "; ".parse_next(input)?;

    let right = take_until_line_ending.parse_next(input)?;

    Ok((left, right))
}

fn project_reference_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ProjectReference<'a>> {
    space0.parse_next(input)?;

    "*\\A".parse_next(input)?;

    let Ok(path): VB6Result<_> = take_until_line_ending.parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::ReferenceMissingSections));
    };

    let reference = VB6ProjectReference::SubProject { path };

    Ok(reference)
}

fn compiled_reference_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ProjectReference<'a>> {
    // This is not the cleanest way to handle this but we need to replace the
    // first instance of "*\\G{" from the start of the segment. Notice the '\\'
    // escape sequence which is just a single slash in the file itself.
    // Then remove
    let (_, uuid_segment, _) = ("*\\G{", take_until(1.., "}#"), "}#").parse_next(input)?;

    let Ok(uuid) = Uuid::parse_str(uuid_segment.to_str().unwrap()) else {
        return Err(ErrMode::Cut(VB6ErrorKind::UnableToParseUuid));
    };

    // still not sure what this element or the next represents.
    let Ok((unknown1, _)): VB6Result<_> = (take_until(1.., "#"), "#").parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::ReferenceMissingSections));
    };

    let Ok((unknown2, _)): VB6Result<_> = (take_until(1.., "#"), "#").parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::ReferenceMissingSections));
    };

    let Ok((path, _)): VB6Result<_> = (take_until(1.., "#"), "#").parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::ReferenceMissingSections));
    };

    let Ok(description): VB6Result<_> = take_until_line_ending.parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::ReferenceMissingSections));
    };

    if description.contains(&b'#') {
        return Err(ErrMode::Cut(VB6ErrorKind::ReferenceExtraSections));
    }

    let reference = VB6ProjectReference::Compiled {
        uuid,
        unknown1,
        unknown2,
        path,
        description,
    };

    Ok(reference)
}

fn reference_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ProjectReference<'a>> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    let reference = alt((project_reference_parse, compiled_reference_parse)).parse_next(input)?;

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    Ok(reference)
}

fn project_type_parse(input: &mut VB6Stream<'_>) -> VB6Result<CompileTargetType> {
    if (space0::<_, VB6Error>, '=', space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    }

    // The first line of any VB6 project file (vbp) is a type line that
    // tells us what kind of project we have.
    // this should be in every project file, even an empty one, and it must
    // be one of these four options.
    //
    // The project type line starts with a 'Type=' has either 'Exe', 'OleDll',
    // 'Control', or 'OleExe'.
    //
    // By this point in the parse the "Type=" component should be stripped off
    // since that is how we knew to use this particular parse.

    let Ok(project_type) = alt::<_, CompileTargetType, VB6ErrorKind, _>((
        "Exe".value(CompileTargetType::Exe),
        "Control".value(CompileTargetType::Control),
        "OleExe".value(CompileTargetType::OleExe),
        "OleDll".value(CompileTargetType::OleDll),
    ))
    .parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::ProjectTypeUnknown));
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    Ok(project_type)
}

#[cfg(test)]
mod tests {
    use winnow::stream::StreamIsPartial;

    use super::*;

    #[test]
    fn compatibility_mode_is_unknown() {
        let mut input = VB6Stream::new("", b"CompatibleMode=\"5\"\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "CompatibleMode".parse_next(&mut input);

        let result = compatibility_mode_parse.parse_next(&mut input);

        assert!(matches!(
            result.err().unwrap().into_inner().unwrap(),
            VB6ErrorKind::CompatibilityModeUnparseable
        ));
    }

    #[test]
    fn compatibility_mode_is_no_compatibility() {
        let mut input = VB6Stream::new("", b"CompatibleMode=\"0\"\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "CompatibleMode".parse_next(&mut input);

        let result = compatibility_mode_parse.parse_next(&mut input).unwrap();

        assert_eq!(result, CompatibilityMode::NoCompatibility);
    }

    #[test]
    fn compatibility_mode_is_project() {
        let mut input = VB6Stream::new("", b"CompatibleMode=\"1\"\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "CompatibleMode".parse_next(&mut input);

        let result = compatibility_mode_parse.parse_next(&mut input).unwrap();
        assert_eq!(result, CompatibilityMode::Project);
    }

    #[test]
    fn compatibility_mode_is_compatible_exe() {
        let mut input = VB6Stream::new("", b"CompatibleMode=\"2\"\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "CompatibleMode".parse_next(&mut input);

        let result = compatibility_mode_parse.parse_next(&mut input).unwrap();

        assert_eq!(result, CompatibilityMode::CompatibleExe);
    }

    #[test]
    fn project_type_is_exe() {
        let mut input = VB6Stream::new("", b"Type=Exe\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Type".parse_next(&mut input);

        let result = project_type_parse.parse_next(&mut input).unwrap();

        assert_eq!(result, CompileTargetType::Exe);
    }

    #[test]
    fn project_type_is_oledll() {
        let mut input = VB6Stream::new("", b"Type=OleDll\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Type".parse_next(&mut input);

        let result = project_type_parse.parse_next(&mut input).unwrap();
        assert_eq!(result, CompileTargetType::OleDll);
    }

    #[test]
    fn project_type_is_control() {
        let mut input = VB6Stream::new("", b"Type=Control\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Type".parse_next(&mut input);

        let result = project_type_parse.parse_next(&mut input).unwrap();

        assert_eq!(result, CompileTargetType::Control);
    }

    #[test]
    fn project_type_is_oleexe() {
        let mut input = VB6Stream::new("", b"Type=OleExe\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Type".parse_next(&mut input);

        let result = project_type_parse.parse_next(&mut input).unwrap();

        assert_eq!(result, CompileTargetType::OleExe);
    }

    #[test]
    fn project_type_is_unknown_type() {
        let mut input = VB6Stream::new("", b"Type=blah\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Type".parse_next(&mut input);

        let result = project_type_parse.parse_next(&mut input);

        assert!(result.is_err());
        assert!(matches!(
            result.err().unwrap().into_inner().unwrap(),
            VB6ErrorKind::ProjectTypeUnknown
        ));
    }

    #[test]
    fn reference_compiled_line_valid() {
        let mut input = VB6Stream::new("", b"Reference=*\\G{000440D8-E9ED-4435-A9A2-06B05387BB16}#c.0#0#..\\DBCommon\\Libs\\VbIntellisenseFix.dll#VbIntellisenseFix\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Reference".parse_next(&mut input);

        let result = reference_parse.parse_next(&mut input).unwrap();

        let expected_uuid = Uuid::parse_str("000440D8-E9ED-4435-A9A2-06B05387BB16").unwrap();

        assert_eq!(input.complete(), 0);
        assert_eq!(matches!(result, VB6ProjectReference::Compiled { .. }), true);
        let result = match result {
            VB6ProjectReference::Compiled {
                uuid,
                unknown1,
                unknown2,
                path,
                description,
            } => (uuid, unknown1, unknown2, path, description),
            _ => unreachable!(),
        };
        assert_eq!(result.0, expected_uuid);
        assert_eq!(result.1, "c.0");
        assert_eq!(result.2, "0");
        assert_eq!(result.3, r"..\DBCommon\Libs\VbIntellisenseFix.dll");
        assert_eq!(result.4, r"VbIntellisenseFix");
    }

    #[test]
    fn reference_subproject_line_valid() {
        let mut input = VB6Stream::new("", b"Reference=*\\Atest.vbp\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Reference".parse_next(&mut input);

        let result = reference_parse.parse_next(&mut input).unwrap();

        assert_eq!(input.complete(), 0);
        assert_eq!(
            result,
            VB6ProjectReference::SubProject {
                path: BStr::new("test.vbp")
            }
        );
    }

    #[test]
    fn module_line_valid() {
        let mut input = VB6Stream::new("", b"Module=modDBAssist; ..\\DBCommon\\DBAssist.bas\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Module".parse_next(&mut input);
        let result = module_parse.parse_next(&mut input).unwrap();

        assert_eq!(input.complete(), 0);
        assert_eq!(result.name, "modDBAssist");
        assert_eq!(result.path, "..\\DBCommon\\DBAssist.bas");
    }

    #[test]
    fn class_line_valid() {
        let mut input = VB6Stream::new(
            "",
            b"Class=CStatusBarClass; ..\\DBCommon\\CStatusBarClass.cls\r\n",
        );

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Class".parse_next(&mut input);
        let result = class_parse.parse_next(&mut input).unwrap();

        assert_eq!(input.complete(), 0);
        assert_eq!(result.name, "CStatusBarClass");
        assert_eq!(result.path, "..\\DBCommon\\CStatusBarClass.cls");
    }

    #[test]
    fn thread_per_object_negative() {
        use bstr::BStr;

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

        let project = VB6Project::parse("project1.vbp", input.as_bytes()).unwrap();

        assert_eq!(project.project_type, CompileTargetType::Exe);
        assert_eq!(project.references.len(), 1);
        assert_eq!(project.objects.len(), 1);
        assert_eq!(project.modules.len(), 1);
        assert_eq!(project.classes.len(), 1);
        assert_eq!(project.designers.len(), 0);
        assert_eq!(project.forms.len(), 2);
        assert_eq!(project.user_controls.len(), 1);
        assert_eq!(project.user_documents.len(), 1);
        assert_eq!(project.upgrade_controls, UpgradeControls::Upgrade);
        assert_eq!(project.res_file_32_path, Some(BStr::new(b"")));
        assert_eq!(project.icon_form, Some(BStr::new(b"")));
        assert_eq!(project.startup, None);
        assert_eq!(project.help_file_path, Some(BStr::new(b"")));
        assert_eq!(project.title, Some(BStr::new(b"Project1")));
        assert_eq!(project.exe_32_file_name, Some(BStr::new(b"Project1.exe")));
        assert_eq!(project.exe_32_compatible, Some(BStr::new(b"")));
        assert_eq!(project.command_line_arguments, None);
        assert_eq!(project.path_32, Some(BStr::new(b"")));
        assert_eq!(project.name, Some(BStr::new(b"Project1")));
        assert_eq!(project.help_context_id, Some(BStr::new(b"0")));
        assert_eq!(
            project.compatibility_mode,
            CompatibilityMode::NoCompatibility
        );
        assert_eq!(project.version_info.major, 1);
        assert_eq!(project.version_info.minor, 0);
        assert_eq!(project.version_info.revision, 0);
        assert_eq!(project.version_info.auto_increment_revision, 0);
        assert_eq!(
            project.version_info.company_name,
            Some(BStr::new(b"Company Name"))
        );
        assert_eq!(
            project.version_info.file_description,
            Some(BStr::new(b"File Description"))
        );
        assert_eq!(
            project.version_info.trademark,
            Some(BStr::new(b"Trademark"))
        );
        assert_eq!(
            project.version_info.product_name,
            Some(BStr::new(b"Product Name"))
        );
        assert_eq!(project.version_info.comments, Some(BStr::new(b"Comments")));
        assert_eq!(
            project.server_support_files,
            ServerSupportFiles::Local,
            "server_support_files check"
        );
        assert_eq!(project.conditional_compile, Some(BStr::new(b"")));
        assert_eq!(
            project.compilation_type,
            CompilationType::NativeCode {
                optimization_type: OptimizationType::FavorFastCode,
                favor_pentium_pro: FavorPentiumPro::False,
                code_view_debug_info: CodeViewDebugInfo::NotCreated,
                aliasing: Aliasing::AssumeAliasing,
                bounds_check: BoundsCheck::CheckBounds,
                overflow_check: OverflowCheck::CheckOverflow,
                floating_point_check: FloatingPointErrorCheck::CheckFloatingPointError,
                pentium_fdiv_bug_check: PentiumFDivBugCheck::CheckPentiumFDivBug,
                unrounded_floating_point: UnroundedFloatingPoint::DoNotAllow,
            }
        );
        assert_eq!(project.start_mode, StartMode::StandAlone);
        assert_eq!(project.unattended, InteractionMode::Interactive);
        assert_eq!(project.retained, Retained::UnloadOnExit);
        assert_eq!(project.thread_per_object, None);
        assert_eq!(project.max_number_of_threads, 1);
        assert_eq!(
            project.debug_startup_option,
            DebugStartupOption::WaitForComponentCreation,
            "debug_startup_option check"
        );
    }

    #[test]
    fn no_startup_selected() {
        use bstr::BStr;

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

        let project = VB6Project::parse("project1.vbp", input.as_bytes()).unwrap();

        assert_eq!(project.project_type, CompileTargetType::Exe);
        assert_eq!(project.references.len(), 1);
        assert_eq!(project.objects.len(), 1);
        assert_eq!(project.modules.len(), 1);
        assert_eq!(project.classes.len(), 1);
        assert_eq!(project.designers.len(), 0);
        assert_eq!(project.forms.len(), 2);
        assert_eq!(project.user_controls.len(), 1);
        assert_eq!(project.user_documents.len(), 1);
        assert_eq!(project.upgrade_controls, UpgradeControls::Upgrade);
        assert_eq!(project.res_file_32_path, Some(BStr::new(b"")));
        assert_eq!(project.icon_form, Some(BStr::new(b"")));
        assert_eq!(project.startup, None);
        assert_eq!(project.help_file_path, Some(BStr::new(b"")));
        assert_eq!(project.title, Some(BStr::new(b"Project1")));
        assert_eq!(project.exe_32_file_name, Some(BStr::new(b"Project1.exe")));
        assert_eq!(project.exe_32_compatible, Some(BStr::new(b"")));
        assert_eq!(project.command_line_arguments, None);
        assert_eq!(project.path_32, Some(BStr::new(b"")));
        assert_eq!(project.name, Some(BStr::new(b"Project1")));
        assert_eq!(project.help_context_id, Some(BStr::new(b"0")));
        assert_eq!(
            project.compatibility_mode,
            CompatibilityMode::NoCompatibility
        );
        assert_eq!(project.version_info.major, 1);
        assert_eq!(project.version_info.minor, 0);
        assert_eq!(project.version_info.revision, 0);
        assert_eq!(project.version_info.auto_increment_revision, 0);
        assert_eq!(
            project.version_info.company_name,
            Some(BStr::new(b"Company Name"))
        );
        assert_eq!(
            project.version_info.file_description,
            Some(BStr::new(b"File Description"))
        );
        assert_eq!(
            project.version_info.trademark,
            Some(BStr::new(b"Trademark"))
        );
        assert_eq!(
            project.version_info.product_name,
            Some(BStr::new(b"Product Name"))
        );
        assert_eq!(project.version_info.comments, Some(BStr::new(b"Comments")));
        assert_eq!(
            project.server_support_files,
            ServerSupportFiles::Local,
            "server_support_files check"
        );
        assert_eq!(project.conditional_compile, Some(BStr::new(b"")));
        assert_eq!(
            project.compilation_type,
            CompilationType::NativeCode {
                optimization_type: OptimizationType::FavorFastCode,
                favor_pentium_pro: FavorPentiumPro::False,
                code_view_debug_info: CodeViewDebugInfo::NotCreated,
                aliasing: Aliasing::AssumeAliasing,
                bounds_check: BoundsCheck::CheckBounds,
                overflow_check: OverflowCheck::CheckOverflow,
                floating_point_check: FloatingPointErrorCheck::CheckFloatingPointError,
                pentium_fdiv_bug_check: PentiumFDivBugCheck::CheckPentiumFDivBug,
                unrounded_floating_point: UnroundedFloatingPoint::DoNotAllow,
            }
        );
        assert_eq!(project.start_mode, StartMode::StandAlone);
        assert_eq!(project.unattended, InteractionMode::Interactive);
        assert_eq!(project.retained, Retained::UnloadOnExit);
        assert_eq!(project.thread_per_object, Some(0));
        assert_eq!(project.max_number_of_threads, 1);
        assert_eq!(
            project.debug_startup_option,
            DebugStartupOption::WaitForComponentCreation,
            "debug_startup_option check"
        );
    }

    #[test]
    fn extra_property_sections() {
        use bstr::BStr;

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

        let project = VB6Project::parse("project1.vbp", input.as_bytes()).unwrap();

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
        assert_eq!(project.upgrade_controls, UpgradeControls::Upgrade);
        assert_eq!(project.res_file_32_path, Some(BStr::new(b"")));
        assert_eq!(project.icon_form, Some(BStr::new(b"")));
        assert_eq!(project.startup, None);
        assert_eq!(project.help_file_path, Some(BStr::new(b"")));
        assert_eq!(project.title, Some(BStr::new(b"Project1")));
        assert_eq!(project.exe_32_file_name, Some(BStr::new(b"Project1.exe")));
        assert_eq!(project.exe_32_compatible, Some(BStr::new(b"")));
        assert_eq!(project.command_line_arguments, None);
        assert_eq!(project.path_32, Some(BStr::new(b"")));
        assert_eq!(project.name, Some(BStr::new(b"Project1")));
        assert_eq!(project.help_context_id, Some(BStr::new(b"0")));
        assert_eq!(
            project.compatibility_mode,
            CompatibilityMode::NoCompatibility
        );
        assert_eq!(project.version_info.major, 1);
        assert_eq!(project.version_info.minor, 0);
        assert_eq!(project.version_info.revision, 0);
        assert_eq!(project.version_info.auto_increment_revision, 0);
        assert_eq!(
            project.version_info.company_name,
            Some(BStr::new(b"Company Name"))
        );
        assert_eq!(
            project.version_info.file_description,
            Some(BStr::new(b"File Description"))
        );
        assert_eq!(
            project.version_info.trademark,
            Some(BStr::new(b"Trademark"))
        );
        assert_eq!(
            project.version_info.product_name,
            Some(BStr::new(b"Product Name"))
        );
        assert_eq!(project.version_info.comments, Some(BStr::new(b"Comments")));
        assert_eq!(
            project.server_support_files,
            ServerSupportFiles::Local,
            "server_support_files check"
        );
        assert_eq!(project.conditional_compile, Some(BStr::new(b"")));

        assert_eq!(
            project.compilation_type,
            CompilationType::NativeCode {
                optimization_type: OptimizationType::FavorFastCode,
                favor_pentium_pro: FavorPentiumPro::False,
                code_view_debug_info: CodeViewDebugInfo::NotCreated,
                aliasing: Aliasing::AssumeAliasing,
                bounds_check: BoundsCheck::CheckBounds,
                overflow_check: OverflowCheck::CheckOverflow,
                floating_point_check: FloatingPointErrorCheck::CheckFloatingPointError,
                pentium_fdiv_bug_check: PentiumFDivBugCheck::CheckPentiumFDivBug,
                unrounded_floating_point: UnroundedFloatingPoint::DoNotAllow,
            }
        );
        assert_eq!(project.start_mode, StartMode::StandAlone);
        assert_eq!(project.unattended, InteractionMode::Interactive);
        assert_eq!(project.retained, Retained::UnloadOnExit);
        assert_eq!(project.thread_per_object, Some(0));
        assert_eq!(project.max_number_of_threads, 1);
        assert_eq!(
            project.debug_startup_option,
            DebugStartupOption::WaitForComponentCreation,
            "debug_startup_option check"
        );
    }
}
