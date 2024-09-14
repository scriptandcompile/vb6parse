use std::collections::HashMap;

use bstr::{BStr, ByteSlice};

use serde::Serialize;
use uuid::Uuid;

use winnow::{
    ascii::{digit1, line_ending, space0},
    combinator::{alt, opt},
    error::ErrMode,
    token::{literal, take_until},
    Parser,
};

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    parsers::{header::object_parse, vb6stream::VB6Stream, VB6ObjectReference},
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
    pub optimization_type: OptimizationType,
    pub favor_pentium_pro: FavorPentiumPro,
    pub code_view_debug_info: CodeViewDebugInfo,
    pub aliasing: Aliasing,
    pub bounds_check: BoundsCheck,
    pub overflow_check: OverflowCheck,
    pub floating_point_check: FloatingPointErrorCheck,
    pub pentium_fdiv_bug_check: PentiumFDivBugCheck,
    pub unrounded_floating_point: UnroundedFloatingPoint,
    pub start_mode: StartMode,
    pub unattended: Unattended,
    pub retained: Retained,
    pub thread_per_object: Option<u16>,
    pub threading_model: ThreadingModel,
    pub max_number_of_threads: u16,
    pub debug_startup_option: DebugStartupOption,
    pub use_existing_browser: UseExistingBrowser,
    pub property_page: Option<&'a BStr>,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum Retained {
    #[default]
    UnloadOnExit = 0,
    RetainedInMemory = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum UseExistingBrowser {
    DoNotUse = 0,
    #[default]
    Use = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum StartMode {
    #[default]
    StandAlone = 0,
    Automation = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum Unattended {
    #[default]
    False = 0,
    True = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum ServerSupportFiles {
    #[default]
    Local = 0,
    Remote = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum UnroundedFloatingPoint {
    #[default]
    DoNotAllow = 0,
    Allow = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum UpgradeControls {
    #[default]
    Upgrade = 0,
    NoUpgrade = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum UnusedControlInfo {
    Retain = 0,
    #[default]
    Remove = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum Aliasing {
    #[default]
    AssumeAliasing = 0,
    AssumeNoAliasing = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum PentiumFDivBugCheck {
    CheckPentiumFDivBug = 0,
    #[default]
    NoPentiumFDivBugCheck = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum BoundsCheck {
    #[default]
    CheckBounds = 0,
    NoBoundsCheck = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum OverflowCheck {
    #[default]
    CheckOverflow = 0,
    NoOverflowCheck = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum FloatingPointErrorCheck {
    #[default]
    CheckFloatingPointError = 0,
    NoFloatingPointErrorCheck = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum CodeViewDebugInfo {
    #[default]
    NotCreated = 0,
    Created = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum FavorPentiumPro {
    #[default]
    False = 0,
    True = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum CompatibilityMode {
    NoCompatibility = 0,
    #[default]
    Project = 1,
    CompatibleExe = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum CompilationType {
    PCode = -1,
    #[default]
    NativeCode = 0,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum DebugStartupOption {
    #[default]
    WaitForComponentCreation = 0,
    StartComponent = 1,
    StartProgram = 2,
    StartBrowser = 3,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum OptimizationType {
    #[default]
    FavorFastCode = 0,
    FavorSmallCode = 1,
    NoOptimization = 2,
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

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
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
    Project {
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
            VB6ProjectReference::Project { path } => {
                let mut state = serializer.serialize_struct("ProjectReference", 1)?;

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
        let mut unattended = Unattended::False;
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
                "[",
                take_until(0.., "]"),
                "]",
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
            };

            // Looks like we are no longer parsing the standard VB6 property section
            // Now we are parsing some third party properties.
            if other_property_group.is_some() {
                let property_name = match (opt(space0), take_until(0.., "=")).parse_next(&mut input)
                {
                    Ok((_, property_name)) => property_name,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (opt(space0::<_, VB6Error>), "=")
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                }

                let property_value = match take_until::<_, _, VB6ErrorKind>(0.., ("\r", "\n"))
                    .parse_next(&mut input)
                {
                    Ok(property_value) => property_value,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                other_properties
                    .get_mut(other_property_group.unwrap())
                    .unwrap()
                    .insert(property_name, property_value);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            let _: VB6Result<_> = space0.parse_next(&mut input);

            // Type
            if literal::<_, _, VB6Error>("Type")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                project_type = match project_type_parse.parse_next(&mut input) {
                    Ok(project_type) => Some(project_type),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Designer")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let Ok(designer) = take_until_line_ending.parse_next(&mut input) else {
                    return Err(input.error(VB6ErrorKind::DesignerLineUnparseable));
                };

                designers.push(designer.as_bstr());

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Reference")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let reference = match reference_parse.parse_next(&mut input) {
                    Ok(reference) => reference,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                references.push(reference);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Object")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let object = match object_parse.parse_next(&mut input) {
                    Ok(object) => object,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                objects.push(object);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Module")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let module = match module_parse.parse_next(&mut input) {
                    Ok(module) => module,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                modules.push(module);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Class")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let class = match class_parse.parse_next(&mut input) {
                    Ok(class) => class,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                classes.push(class);

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("RelatedDoc")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

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
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

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
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

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
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

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
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                res_file_32_path = match qouted_value("\"").parse_next(&mut input) {
                    Ok(res_file_32_path) => Some(res_file_32_path),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("IconForm")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                icon_form = match qouted_value("\"").parse_next(&mut input) {
                    Ok(icon_form) => Some(icon_form),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Startup")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                // if the project lacks a startup object/function/etc it will be !(None)! or !! in the file
                // which is distinct from the double qouted way of targeting a specific object.
                startup = match alt((qouted_value("\""), qouted_value("!"))).parse_next(&mut input)
                {
                    Ok(startup) => {
                        // if we have !(None)! or !! then we have no startup object.
                        if startup == "(None)" || startup == "" {
                            None
                        } else {
                            Some(startup)
                        }
                    }
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("HelpFile")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                help_file_path = match qouted_value("\"").parse_next(&mut input) {
                    Ok(help_file_path) => Some(help_file_path),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Title")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                // it's perfectly possible to use '"' within the title string.
                // VB6 being the language it is, there is no escape sequence for
                // this. Instead, the title is wrapped in quotes and the quotes
                // are just simply included in the text. This means we can't use
                // the qouted_value parser here.
                title = match title_parse.parse_next(&mut input) {
                    Ok(title) => Some(title),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("ExeName32")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                exe_32_file_name = match qouted_value("\"").parse_next(&mut input) {
                    Ok(exe_32_file_name) => Some(exe_32_file_name),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Path32")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                path_32 = match qouted_value("\"").parse_next(&mut input) {
                    Ok(path_32) => Some(path_32),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Command32")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                // if the project lacks a commandline it will be !(None)! or !! or "" in the file
                // which is distinct from the double qouted way of targeting a specific object.
                command_line_arguments =
                    match alt((qouted_value("\""), qouted_value("!"))).parse_next(&mut input) {
                        Ok(command_line_arguments) => {
                            // if we have !(None)! or !! then we have no command32 line.
                            if command_line_arguments == "(None)" || command_line_arguments == "" {
                                None
                            } else {
                                Some(command_line_arguments)
                            }
                        }
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Name")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                // if the project lacks a name it will be !(None)! or !! or "" in the file..
                name = match alt((qouted_value("\""), qouted_value("!"))).parse_next(&mut input) {
                    Ok(name) => {
                        // if we have !(None)! or !! then we have no command32 line.
                        if name == "(None)" || name == "" {
                            None
                        } else {
                            Some(name)
                        }
                    }
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("Description")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                description = match qouted_value("\"").parse_next(&mut input) {
                    Ok(description) => Some(description),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("HelpContextID")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                // if the project lacks a help_context_id it will be !(None)! or !! or "" in the file..
                help_context_id =
                    match alt((qouted_value("\""), qouted_value("!"))).parse_next(&mut input) {
                        Ok(help_context_id) => {
                            // if we have !(None)! or !! then we have no command32 line.
                            if help_context_id == "(None)" || help_context_id == "" {
                                None
                            } else {
                                Some(help_context_id)
                            }
                        }
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("CompatibleMode")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                compatibility_mode =
                    match alt((qouted_value("\""), qouted_value("!"))).parse_next(&mut input) {
                        Ok(compatible_mode) => match compatible_mode.as_bytes() {
                            b"0" => CompatibilityMode::NoCompatibility,
                            b"1" => CompatibilityMode::Project,
                            b"2" => CompatibilityMode::CompatibleExe,
                            _ => {
                                return Err(input.error(VB6ErrorKind::CompatibleModeUnparseable));
                            }
                        },
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("VersionCompatible32")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                version_32_compatibility = match qouted_value("\"").parse_next(&mut input) {
                    Ok(version_32_compatibility) => Some(version_32_compatibility),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("CompatibleEXE32")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                exe_32_compatible = match qouted_value("\"").parse_next(&mut input) {
                    Ok(exe_32_compatible) => Some(exe_32_compatible),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("DllBaseAddress")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let Ok(base_address_hex_text): VB6Result<_> =
                    take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::DllBaseAddressUnparseable));
                };

                dll_base_address = match u32::from_str_radix(
                    base_address_hex_text.to_string().trim_start_matches("&H"),
                    16,
                ) {
                    Ok(dll_base_address) => dll_base_address,
                    Err(_) => {
                        return Err(input.error(VB6ErrorKind::DllBaseAddressUnparseable));
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

            if literal::<_, _, VB6Error>("RemoveUnusedControlInfo")
                .parse_next(&mut input)
                .is_ok()
            {
                unused_control_info = match unused_control_info_parse.parse_next(&mut input) {
                    Ok(unused_control_info) => unused_control_info,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("MajorVer")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let Ok(major_ver): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::MajorVersionUnparseable));
                };

                major = match major_ver.to_string().as_str().parse::<u16>() {
                    Ok(major) => major,
                    Err(_) => {
                        return Err(input.error(VB6ErrorKind::MajorVersionUnparseable));
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

            if literal::<_, _, VB6Error>("MinorVer")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let Ok(minor_ver): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::MinorVersionUnparseable));
                };

                minor = match minor_ver.to_string().as_str().parse::<u16>() {
                    Ok(minor) => minor,
                    Err(_) => {
                        return Err(input.error(VB6ErrorKind::MinorVersionUnparseable));
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

            if literal::<_, _, VB6Error>("RevisionVer")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let Ok(revision_ver): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::RevisionVersionUnparseable));
                };

                revision = match revision_ver.to_string().as_str().parse::<u16>() {
                    Ok(revision) => revision,
                    Err(_) => {
                        return Err(input.error(VB6ErrorKind::RevisionVersionUnparseable));
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

            if literal::<_, _, VB6Error>("ThreadingModel")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let Ok(threading_model_text): VB6Result<_> =
                    take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::ThreadingModelUnparseable));
                };

                threading_model = match threading_model_text.to_string().trim().parse::<u16>() {
                    Ok(0) => ThreadingModel::SingleThreaded,
                    Ok(1) => ThreadingModel::ApartmentThreaded,
                    Ok(_) | Err(_) => {
                        return Err(input.error(VB6ErrorKind::ThreadingModelInvalid));
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

            if literal::<_, _, VB6Error>("AutoIncrementVer")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                let Ok(auto_increment): VB6Result<_> =
                    take_until_line_ending.parse_next(&mut input)
                else {
                    return Err(input.error(VB6ErrorKind::AutoIncrementUnparseable));
                };

                auto_increment_revision = match auto_increment.to_string().as_str().parse::<u16>() {
                    Ok(auto_increment_revision) => auto_increment_revision,
                    Err(_) => {
                        return Err(input.error(VB6ErrorKind::AutoIncrementUnparseable));
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

            if literal::<_, _, VB6Error>("PropertyPage")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

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
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

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
                upgrade_controls = match upgrade_controls_parse.parse_next(&mut input) {
                    Ok(upgrade_controls) => upgrade_controls,
                    Err(e) => {
                        return Err(input.error(e.into_inner().unwrap()));
                    }
                };

                continue;
            }

            if literal::<_, _, VB6Error>("ServerSupportFiles")
                .parse_next(&mut input)
                .is_ok()
            {
                server_support_files = match server_support_files_parse.parse_next(&mut input) {
                    Ok(server_support_files) => server_support_files,
                    Err(e) => {
                        return Err(input.error(e.into_inner().unwrap()));
                    }
                };

                continue;
            }

            if literal::<_, _, VB6Error>("VersionCompanyName")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                company_name = match qouted_value("\"").parse_next(&mut input) {
                    Ok(company_name) => Some(company_name),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("VersionFileDescription")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                file_description = match qouted_value("\"").parse_next(&mut input) {
                    Ok(file_description) => Some(file_description),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("VersionLegalCopyright")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                copyright = match qouted_value("\"").parse_next(&mut input) {
                    Ok(copyright) => Some(copyright),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("VersionLegalTrademarks")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                trademark = match qouted_value("\"").parse_next(&mut input) {
                    Ok(trademark) => Some(trademark),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("VersionProductName")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                product_name = match qouted_value("\"").parse_next(&mut input) {
                    Ok(product_name) => Some(product_name),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("VersionComments")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                comments = match qouted_value("\"").parse_next(&mut input) {
                    Ok(comments) => Some(comments),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("CondComp")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                conditional_compile = match qouted_value("\"").parse_next(&mut input) {
                    Ok(conditional_compile) => Some(conditional_compile),
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("CompilationType")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                compilation_type = match (opt("-"), digit1::<_, VB6ErrorKind>)
                    .take()
                    .parse_next(&mut input)
                {
                    Ok(compilation_type) => match compilation_type.as_bytes() {
                        b"-1" => CompilationType::PCode,
                        b"0" => CompilationType::NativeCode,
                        _ => return Err(input.error(VB6ErrorKind::CompilationTypeUnparseable)),
                    },
                    Err(_) => return Err(input.error(VB6ErrorKind::CompilationTypeUnparseable)),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("OptimizationType")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                optimization_type = match digit1::<_, VB6ErrorKind>.parse_next(&mut input) {
                    Ok(op) => match op.as_bytes() {
                        b"0" => OptimizationType::FavorFastCode,
                        b"1" => OptimizationType::FavorSmallCode,
                        b"2" => OptimizationType::NoOptimization,
                        _ => return Err(input.error(VB6ErrorKind::OptimizationTypeUnparseable)),
                    },
                    Err(_) => return Err(input.error(VB6ErrorKind::OptimizationTypeUnparseable)),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("FavorPentiumPro(tm)")
                .parse_next(&mut input)
                .is_ok()
            {
                favor_pentium_pro = match favor_pentium_pro_parse.parse_next(&mut input) {
                    Ok(favor_pentium_pro) => favor_pentium_pro,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("CodeViewDebugInfo")
                .parse_next(&mut input)
                .is_ok()
            {
                code_view_debug_info = match code_view_debug_info_parse.parse_next(&mut input) {
                    Ok(code_view_debug_info) => code_view_debug_info,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("NoAliasing")
                .parse_next(&mut input)
                .is_ok()
            {
                aliasing = match aliasing_parse.parse_next(&mut input) {
                    Ok(aliasing) => aliasing,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("BoundsCheck")
                .parse_next(&mut input)
                .is_ok()
            {
                bounds_check = match bounds_check_parse.parse_next(&mut input) {
                    Ok(bounds_check) => bounds_check,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("OverflowCheck")
                .parse_next(&mut input)
                .is_ok()
            {
                overflow_check = match overflow_check_parse.parse_next(&mut input) {
                    Ok(overflow_check) => overflow_check,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("FlPointCheck")
                .parse_next(&mut input)
                .is_ok()
            {
                floating_point_check = match floating_point_error_check_parse.parse_next(&mut input)
                {
                    Ok(floating_point_check) => floating_point_check,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("FDIVCheck")
                .parse_next(&mut input)
                .is_ok()
            {
                pentium_fdiv_bug_check = match pentium_fdiv_bug_check_parse.parse_next(&mut input) {
                    Ok(pentium_fdiv_bug_check) => pentium_fdiv_bug_check,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("UnroundedFP")
                .parse_next(&mut input)
                .is_ok()
            {
                unrounded_floating_point =
                    match unrounded_floating_point_parse.parse_next(&mut input) {
                        Ok(unrounded_floating_point) => unrounded_floating_point,
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };

                continue;
            }

            if literal::<_, _, VB6Error>("StartMode")
                .parse_next(&mut input)
                .is_ok()
            {
                start_mode = match start_mode_parse.parse_next(&mut input) {
                    Ok(start_mode) => start_mode,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("Unattended")
                .parse_next(&mut input)
                .is_ok()
            {
                unattended = match unattended_parse.parse_next(&mut input) {
                    Ok(unattended) => unattended,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("Retained")
                .parse_next(&mut input)
                .is_ok()
            {
                retained = match retained_parse.parse_next(&mut input) {
                    Ok(retained) => retained,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

                continue;
            }

            if literal::<_, _, VB6Error>("ThreadPerObject")
                .parse_next(&mut input)
                .is_ok()
            {
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

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
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

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
                if (space0::<_, VB6Error>, "=", space0)
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoEqualSplit));
                };

                debug_startup_option = match take_until_line_ending.parse_next(&mut input) {
                    Ok(debug_startup_option) => match debug_startup_option.as_bytes() {
                        b"0" => DebugStartupOption::WaitForComponentCreation,
                        b"1" => DebugStartupOption::StartComponent,
                        b"2" => DebugStartupOption::StartProgram,
                        b"3" => DebugStartupOption::StartBrowser,
                        _ => return Err(input.error(VB6ErrorKind::DebugStartupOptionUnparseable)),
                    },
                    Err(_) => return Err(input.error(VB6ErrorKind::DebugStartupOptionUnparseable)),
                };

                if (space0, alt((line_ending, line_comment_parse)))
                    .parse_next(&mut input)
                    .is_err()
                {
                    return Err(input.error(VB6ErrorKind::NoLineEnding));
                }

                continue;
            }

            if literal::<_, _, VB6Error>("UseExistingBrowser")
                .parse_next(&mut input)
                .is_ok()
            {
                use_existing_browser = match use_existing_browser_parse.parse_next(&mut input) {
                    Ok(use_existing_browser) => use_existing_browser,
                    Err(e) => return Err(input.error(e.into_inner().unwrap())),
                };

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
            optimization_type,
            favor_pentium_pro,
            code_view_debug_info,
            aliasing,
            bounds_check,
            overflow_check,
            floating_point_check,
            pentium_fdiv_bug_check,
            unrounded_floating_point,
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
    pub fn get_project_references(&self) -> Vec<&VB6ProjectReference> {
        self.references
            .iter()
            .filter(|reference| matches!(reference, VB6ProjectReference::Project { .. }))
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

fn use_existing_browser_parse(input: &mut VB6Stream<'_>) -> VB6Result<UseExistingBrowser> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(UseExistingBrowser::DoNotUse),
        "-1".value(UseExistingBrowser::Use),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::UseExistingBrowserUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn retained_parse(input: &mut VB6Stream<'_>) -> VB6Result<Retained> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(Retained::UnloadOnExit),
        "1".value(Retained::RetainedInMemory),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::RetainedUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn unattended_parse(input: &mut VB6Stream<'_>) -> VB6Result<Unattended> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(Unattended::False),
        "-1".value(Unattended::True),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::UnattendedUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn unrounded_floating_point_parse(input: &mut VB6Stream<'_>) -> VB6Result<UnroundedFloatingPoint> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(UnroundedFloatingPoint::DoNotAllow),
        "-1".value(UnroundedFloatingPoint::Allow),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::NoAliasingUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn start_mode_parse(input: &mut VB6Stream<'_>) -> VB6Result<StartMode> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(StartMode::StandAlone),
        "1".value(StartMode::Automation),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::StartModeUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn aliasing_parse(input: &mut VB6Stream<'_>) -> VB6Result<Aliasing> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(Aliasing::AssumeAliasing),
        "-1".value(Aliasing::AssumeNoAliasing),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::NoAliasingUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn upgrade_controls_parse(input: &mut VB6Stream<'_>) -> VB6Result<UpgradeControls> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(UpgradeControls::Upgrade),
        "1".value(UpgradeControls::NoUpgrade),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::NoControlUpgradeUnparsable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn unused_control_info_parse(input: &mut VB6Stream<'_>) -> VB6Result<UnusedControlInfo> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(UnusedControlInfo::Retain),
        "1".value(UnusedControlInfo::Remove),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::UnusedControlInfoUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn overflow_check_parse(input: &mut VB6Stream<'_>) -> VB6Result<OverflowCheck> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(OverflowCheck::CheckOverflow),
        "-1".value(OverflowCheck::NoOverflowCheck),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::OverflowCheckUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn pentium_fdiv_bug_check_parse(input: &mut VB6Stream<'_>) -> VB6Result<PentiumFDivBugCheck> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(PentiumFDivBugCheck::CheckPentiumFDivBug),
        "-1".value(PentiumFDivBugCheck::NoPentiumFDivBugCheck),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::FDIVCheckUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn floating_point_error_check_parse(
    input: &mut VB6Stream<'_>,
) -> VB6Result<FloatingPointErrorCheck> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(FloatingPointErrorCheck::CheckFloatingPointError),
        "-1".value(FloatingPointErrorCheck::NoFloatingPointErrorCheck),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::FlPointCheckUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn bounds_check_parse(input: &mut VB6Stream<'_>) -> VB6Result<BoundsCheck> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(BoundsCheck::CheckBounds),
        "-1".value(BoundsCheck::NoBoundsCheck),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::BoundsCheckUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn code_view_debug_info_parse(input: &mut VB6Stream<'_>) -> VB6Result<CodeViewDebugInfo> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(CodeViewDebugInfo::NotCreated),
        "-1".value(CodeViewDebugInfo::Created),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::CodeViewDebugInfoUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn favor_pentium_pro_parse(input: &mut VB6Stream<'_>) -> VB6Result<FavorPentiumPro> {
    if (space0::<_, VB6ErrorKind>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let result = match alt::<_, _, VB6ErrorKind, _>((
        "0".value(FavorPentiumPro::False),
        "-1".value(FavorPentiumPro::True),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::FavorPentiumProUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    result
}

fn server_support_files_parse(input: &mut VB6Stream<'_>) -> VB6Result<ServerSupportFiles> {
    if (space0::<_, VB6Error>, "=", space0)
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoEqualSplit));
    };

    let value = match alt::<_, _, VB6ErrorKind, _>((
        '0'.value(ServerSupportFiles::Local),
        '1'.value(ServerSupportFiles::Remote),
    ))
    .parse_next(input)
    {
        Ok(result) => Ok(result),
        Err(_) => Err(ErrMode::Cut(VB6ErrorKind::ServerSupportFilesUnparseable)),
    };

    if (space0, alt((line_ending, line_comment_parse)))
        .parse_next(input)
        .is_err()
    {
        return Err(ErrMode::Cut(VB6ErrorKind::NoLineEnding));
    }

    value
}

fn title_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    // it's perfectly possible to use '"' within the title string.
    // VB6 being the language it is, there is no escape sequence for
    // this. Instead, the title is wrapped in quotes and the quotes
    // are just simply included in the text. This means we can't use
    // the qouted_value parser here.

    let _: VB6Result<_> = (space0, "\"").parse_next(input);

    let Ok(title): VB6Result<_> =
        alt((take_until(1.., "\"\r\n"), take_until(1.., "\"\n"))).parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ErrorKind::TitleUnparseable));
    };

    // We need to skip the closing quote.
    // But we also need to make sure we don't skip the line ending.
    // This is a bit odd, but all the other one off line parsers don't read
    // the line ending, so we need to make sure this one doesn't either.
    let _: VB6Result<_> = "\"".parse_next(input);

    Ok(title)
}

fn qouted_value<'a>(qoute_char: &'a str) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<&'a BStr> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<&'a BStr> {
        literal(qoute_char).parse_next(input)?;
        let qouted_value = take_until(0.., qoute_char).parse_next(input)?;
        literal(qoute_char).parse_next(input)?;

        Ok(qouted_value)
    }
}

fn module_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ProjectModule<'a>> {
    let (name, path) = semicolon_space_split_parse.parse_next(input)?;

    let name = name.as_bstr();
    let path = path.as_bstr();

    let module = VB6ProjectModule { name, path };

    Ok(module)
}

fn class_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ProjectClass<'a>> {
    let (name, path) = semicolon_space_split_parse.parse_next(input)?;

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

    let reference = VB6ProjectReference::Project { path };

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
    alt((project_reference_parse, compiled_reference_parse)).parse_next(input)
}

fn project_type_parse(input: &mut VB6Stream<'_>) -> VB6Result<CompileTargetType> {
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

    let project_type = alt::<_, CompileTargetType, VB6ErrorKind, _>((
        "Exe".value(CompileTargetType::Exe),
        "Control".value(CompileTargetType::Control),
        "OleExe".value(CompileTargetType::OleExe),
        "OleDll".value(CompileTargetType::OleDll),
    ))
    .parse_next(input)?;

    Ok(project_type)
}

#[cfg(test)]
mod tests {
    use winnow::stream::StreamIsPartial;

    use super::*;

    #[test]
    fn project_type_is_exe() {
        let mut input = VB6Stream::new("", b"Type=Exe");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Type=".parse_next(&mut input);

        let result = project_type_parse.parse_next(&mut input).unwrap();

        assert_eq!(result, CompileTargetType::Exe);
    }

    #[test]
    fn project_type_is_oledll() {
        let mut input = VB6Stream::new("", b"Type=OleDll");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Type=".parse_next(&mut input);

        let result = project_type_parse.parse_next(&mut input).unwrap();
        assert_eq!(result, CompileTargetType::OleDll);
    }

    #[test]
    fn project_type_is_unknown_type() {
        let mut input = VB6Stream::new("", b"Type=blah");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Type=".parse_next(&mut input);

        let result = project_type_parse.parse_next(&mut input);
        assert!(result.is_err());
    }

    #[test]
    fn reference_line_valid() {
        let mut input = VB6Stream::new("", b"Reference=*\\G{000440D8-E9ED-4435-A9A2-06B05387BB16}#c.0#0#..\\DBCommon\\Libs\\VbIntellisenseFix.dll#VbIntellisenseFix\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Reference=".parse_next(&mut input);

        let result = reference_parse.parse_next(&mut input).unwrap();

        let expected_uuid = Uuid::parse_str("000440D8-E9ED-4435-A9A2-06B05387BB16").unwrap();

        // we don't consume the line ending, so we should have 2 bytes left.
        assert_eq!(input.complete(), 2);
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
    fn module_line_valid() {
        let mut input = VB6Stream::new("", b"Module=modDBAssist; ..\\DBCommon\\DBAssist.bas\r\n");

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Module=".parse_next(&mut input);
        let result = module_parse.parse_next(&mut input).unwrap();

        // we don't consume the line ending, so we should have 2 bytes left.
        assert_eq!(input.complete(), 2);
        assert_eq!(result.name, "modDBAssist");
        assert_eq!(result.path, "..\\DBCommon\\DBAssist.bas");
    }

    #[test]
    fn class_line_valid() {
        let mut input = VB6Stream::new(
            "",
            b"Class=CStatusBarClass; ..\\DBCommon\\CStatusBarClass.cls\r\n",
        );

        let _: Result<&BStr, ErrMode<VB6ErrorKind>> = "Class=".parse_next(&mut input);
        let result = class_parse.parse_next(&mut input).unwrap();

        // we don't consume the line ending, so we should have 2 bytes left.
        assert_eq!(input.complete(), 2);
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
        assert_eq!(project.compilation_type, CompilationType::NativeCode);
        assert_eq!(project.optimization_type, OptimizationType::FavorFastCode);
        assert_eq!(project.favor_pentium_pro, FavorPentiumPro::False);
        assert_eq!(project.code_view_debug_info, CodeViewDebugInfo::NotCreated,);
        assert_eq!(project.aliasing, Aliasing::AssumeAliasing);
        assert_eq!(project.bounds_check, BoundsCheck::CheckBounds);
        assert_eq!(project.overflow_check, OverflowCheck::CheckOverflow);
        assert_eq!(
            project.floating_point_check,
            FloatingPointErrorCheck::CheckFloatingPointError
        );
        assert_eq!(
            project.pentium_fdiv_bug_check,
            PentiumFDivBugCheck::CheckPentiumFDivBug
        );
        assert_eq!(
            project.unrounded_floating_point,
            UnroundedFloatingPoint::DoNotAllow
        );
        assert_eq!(project.start_mode, StartMode::StandAlone);
        assert_eq!(project.unattended, Unattended::False);
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
        assert_eq!(project.compilation_type, CompilationType::NativeCode);
        assert_eq!(project.optimization_type, OptimizationType::FavorFastCode);
        assert_eq!(project.favor_pentium_pro, FavorPentiumPro::False);
        assert_eq!(project.code_view_debug_info, CodeViewDebugInfo::NotCreated,);
        assert_eq!(project.aliasing, Aliasing::AssumeAliasing);
        assert_eq!(project.bounds_check, BoundsCheck::CheckBounds);
        assert_eq!(project.overflow_check, OverflowCheck::CheckOverflow);
        assert_eq!(
            project.floating_point_check,
            FloatingPointErrorCheck::CheckFloatingPointError
        );
        assert_eq!(
            project.pentium_fdiv_bug_check,
            PentiumFDivBugCheck::CheckPentiumFDivBug
        );
        assert_eq!(
            project.unrounded_floating_point,
            UnroundedFloatingPoint::DoNotAllow
        );
        assert_eq!(project.start_mode, StartMode::StandAlone);
        assert_eq!(project.unattended, Unattended::False);
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
        assert_eq!(project.compilation_type, CompilationType::NativeCode);
        assert_eq!(project.optimization_type, OptimizationType::FavorFastCode);
        assert_eq!(project.favor_pentium_pro, FavorPentiumPro::False);
        assert_eq!(project.code_view_debug_info, CodeViewDebugInfo::NotCreated,);
        assert_eq!(project.aliasing, Aliasing::AssumeAliasing);
        assert_eq!(project.bounds_check, BoundsCheck::CheckBounds);
        assert_eq!(project.overflow_check, OverflowCheck::CheckOverflow);
        assert_eq!(
            project.floating_point_check,
            FloatingPointErrorCheck::CheckFloatingPointError
        );
        assert_eq!(
            project.pentium_fdiv_bug_check,
            PentiumFDivBugCheck::CheckPentiumFDivBug
        );
        assert_eq!(
            project.unrounded_floating_point,
            UnroundedFloatingPoint::DoNotAllow
        );
        assert_eq!(project.start_mode, StartMode::StandAlone);
        assert_eq!(project.unattended, Unattended::False);
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
