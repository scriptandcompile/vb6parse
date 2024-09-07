use bstr::{BStr, ByteSlice};

use uuid::Uuid;

use winnow::{
    ascii::{line_ending, space0},
    combinator::alt,
    error::ErrMode,
    token::{literal, take_until},
    Parser,
};

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    parsers::{header::object_parse, vb6stream::VB6Stream, VB6ObjectReference},
    vb6::{line_comment_parse, take_until_line_ending, VB6Result},
};

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Project<'a> {
    pub project_type: CompileTargetType,
    pub references: Vec<VB6ProjectReference<'a>>,
    pub objects: Vec<VB6ObjectReference<'a>>,
    pub modules: Vec<VB6ProjectModule<'a>>,
    pub classes: Vec<VB6ProjectClass<'a>>,
    pub designers: Vec<&'a BStr>,
    pub forms: Vec<&'a BStr>,
    pub user_controls: Vec<&'a BStr>,
    pub user_documents: Vec<&'a BStr>,
    pub upgrade_activex_controls: bool,
    pub res_file_32_path: Option<&'a BStr>,
    pub icon_form: Option<&'a BStr>,
    pub startup: Option<&'a BStr>,
    pub help_file_path: Option<&'a BStr>,
    pub title: Option<&'a BStr>,
    pub exe_32_file_name: Option<&'a BStr>,
    pub path_32: Option<&'a BStr>,
    pub command_line_arguments: Option<&'a BStr>,
    pub name: Option<&'a BStr>,
    pub description: Option<&'a BStr>,
    // May need to be switched to a u32. Not sure yet.
    pub help_context_id: Option<&'a BStr>,
    pub compatible_mode: bool,
    pub version_info: VersionInformation<'a>,
    pub server_support_files: bool,
    pub conditional_compile: Option<&'a BStr>,
    pub compilation_type: bool,
    pub optimization_type: bool,
    pub favor_pentium_pro: bool,
    pub code_view_debug_info: bool,
    pub aliasing: bool,
    pub bounds_check: bool,
    pub overflow_check: bool,
    pub floating_point_check: bool,
    pub pentium_fdiv_bug_check: bool,
    pub unrounded_floating_point: bool,
    pub start_mode: bool,
    pub unattended: bool,
    pub retained: bool,
    pub thread_per_object: u16,
    pub max_number_of_threads: u16,
    pub debug_startup_option: bool,
    pub auto_refresh: bool,
}

#[derive(Debug, PartialEq, Eq, Clone)]
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

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum CompileTargetType {
    Exe,
    Control,
    OleExe,
    OleDll,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ProjectReference<'a> {
    pub uuid: Uuid,
    pub unknown1: &'a BStr,
    pub unknown2: &'a BStr,
    pub path: &'a BStr,
    pub description: &'a BStr,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ProjectModule<'a> {
    pub name: &'a BStr,
    pub path: &'a BStr,
}

#[derive(Debug, PartialEq, Eq, Clone)]
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
    /// AutoRefresh=0
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
    /// assert_eq!(project.upgrade_activex_controls, true, "NoControlUpgrade is inverted fromt the file as upgrade_activex_controls");
    /// assert_eq!(project.res_file_32_path, Some(BStr::new(b"")));
    /// assert_eq!(project.icon_form, Some(BStr::new(b"")));
    /// assert_eq!(project.startup, Some(BStr::new(b"Form1")));
    /// assert_eq!(project.help_file_path, Some(BStr::new(b"")));
    /// assert_eq!(project.title, Some(BStr::new(b"Project1")));
    /// assert_eq!(project.exe_32_file_name, Some(BStr::new(b"Project1.exe")));
    /// assert_eq!(project.command_line_arguments, Some(BStr::new(b"")));
    /// assert_eq!(project.path_32, Some(BStr::new(b"")));
    /// assert_eq!(project.name, Some(BStr::new(b"Project1")));
    /// assert_eq!(project.help_context_id, Some(BStr::new(b"0")));
    /// assert_eq!(project.compatible_mode, false, "compatible_mode check");
    /// assert_eq!(project.version_info.major, 1);
    /// assert_eq!(project.version_info.minor, 0);
    /// assert_eq!(project.version_info.revision, 0);
    /// assert_eq!(project.version_info.auto_increment_revision, 0);
    /// assert_eq!(project.version_info.company_name, Some(BStr::new(b"Company Name")));
    /// assert_eq!(project.version_info.file_description, Some(BStr::new(b"File Description")));
    /// assert_eq!(project.version_info.trademark, Some(BStr::new(b"Trademark")));
    /// assert_eq!(project.version_info.product_name, Some(BStr::new(b"Product Name")));
    /// assert_eq!(project.version_info.comments, Some(BStr::new(b"Comments")));
    /// assert_eq!(project.server_support_files, false, "server_support_files check");
    /// assert_eq!(project.conditional_compile, Some(BStr::new(b"")));
    /// assert_eq!(project.compilation_type, false, "compilation_type check");
    /// assert_eq!(project.optimization_type, false, "optimization_type check");
    /// assert_eq!(project.favor_pentium_pro, false, "favor_pentium_pro check");
    /// assert_eq!(project.code_view_debug_info, false, "code_view_debug_info check");
    /// assert_eq!(project.aliasing, true, "NoAliasing is inverted from the file as aliasing");
    /// assert_eq!(project.bounds_check, false, "bounds_check check");
    /// assert_eq!(project.overflow_check, false, "overflow_check check");
    /// assert_eq!(project.floating_point_check, false, "floating_point_check check");
    /// assert_eq!(project.pentium_fdiv_bug_check, false, "pentium_fdiv_bug_check check");
    /// assert_eq!(project.unrounded_floating_point, false, "unrounded_floating_point check");
    /// assert_eq!(project.start_mode, false, "start_mode check");
    /// assert_eq!(project.unattended, false, "unattended check");
    /// assert_eq!(project.retained, false, "retained check");
    /// assert_eq!(project.thread_per_object, 0);
    /// assert_eq!(project.max_number_of_threads, 1);
    /// assert_eq!(project.debug_startup_option, false, "debug_startup_option check");
    /// assert_eq!(project.auto_refresh, false, "auto_refresh check");
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

        let mut project_type: Option<CompileTargetType> = None;
        let mut res_file_32_path = Some(BStr::new(b""));
        let mut icon_form = Some(BStr::new(b""));
        let mut startup = Some(BStr::new(b""));
        let mut help_file_path = Some(BStr::new(b""));
        let mut title = Some(BStr::new(b""));
        let mut exe_32_file_name = Some(BStr::new(b""));
        let mut path_32 = Some(BStr::new(b""));
        let mut command_line_arguments = Some(BStr::new(b""));
        let mut name = Some(BStr::new(b""));
        let mut description = Some(BStr::new(b""));
        let mut help_context_id = Some(BStr::new(b""));
        let mut compatible_mode = false;
        let mut upgrade_activex_controls = true; // True is the default.
        let mut server_support_files = false;
        let mut conditional_compile = Some(BStr::new(b""));
        let mut compilation_type = false;
        let mut optimization_type = false;
        let mut favor_pentium_pro = false;
        let mut code_view_debug_info = false;
        let mut aliasing = false;
        let mut bounds_check = false;
        let mut overflow_check = false;
        let mut floating_point_check = false;
        let mut pentium_fdiv_bug_check = false;
        let mut unrounded_floating_point = false;
        let mut start_mode = false;
        let mut unattended = false;
        let mut retained = false;
        let mut thread_per_object = 0;
        let mut max_number_of_threads = 1;
        let mut debug_startup_option = false;
        let mut auto_refresh = false;

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

        while !input.is_empty() {
            // We also want to skip any '[MS Transaction Server]' header lines.
            // There should only be one in the file since it's only used once,
            // but we want to be flexible in what we accept so we skip any of
            // these kinds of header lines.
            if line_ending::<_, VB6Error>.parse_next(&mut input).is_ok() {
                continue;
            };

            // We also want to skip any '[MS Transaction Server]' or other section header lines.
            if ("[", take_until(0.., "]"), "]", line_ending::<_, VB6Error>)
                .parse_next(&mut input)
                .is_ok()
            {
                continue;
            };

            let _: VB6Result<_> = space0.parse_next(&mut input);

            let Ok(key): VB6Result<_> = take_until(1.., "=").parse_next(&mut input) else {
                return Err(input.error(VB6ErrorKind::NoEqualSplit));
            };

            let _: VB6Result<_> = ("=", space0).parse_next(&mut input);

            match key.to_str() {
                Ok("Type") => {
                    project_type = match project_type_parse.parse_next(&mut input) {
                        Ok(project_type) => Some(project_type),
                        Err(_) => {
                            return Err(input.error(VB6ErrorKind::ProjectTypeUnknown));
                        }
                    };
                }
                Ok("Designer") => {
                    let Ok(designer): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                    else {
                        return Err(input.error(VB6ErrorKind::DesignerLineUnparseable));
                    };

                    designers.push(designer.as_bstr());
                }
                Ok("Reference") => {
                    let reference = match reference_parse.parse_next(&mut input) {
                        Ok(reference) => reference,
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };

                    references.push(reference);
                }
                Ok("Object") => {
                    let object = match object_parse.parse_next(&mut input) {
                        Ok(object) => object,
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };

                    objects.push(object);
                }
                Ok("Module") => {
                    let module = match module_parse.parse_next(&mut input) {
                        Ok(module) => module,
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };

                    modules.push(module);
                }
                Ok("Class") => {
                    let class = match class_parse.parse_next(&mut input) {
                        Ok(class) => class,
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };

                    classes.push(class);
                }
                Ok("Form") => {
                    let Ok(form): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                    else {
                        return Err(input.error(VB6ErrorKind::FormLineUnparseable));
                    };

                    forms.push(form);
                }
                Ok("UserControl") => {
                    let Ok(user_control): VB6Result<_> =
                        take_until_line_ending.parse_next(&mut input)
                    else {
                        return Err(input.error(VB6ErrorKind::UserControlLineUnparseable));
                    };

                    user_controls.push(user_control);
                }
                Ok("UserDocument") => {
                    let Ok(user_document): VB6Result<_> =
                        take_until_line_ending.parse_next(&mut input)
                    else {
                        return Err(input.error(VB6ErrorKind::UserDocumentLineUnparseable));
                    };

                    user_documents.push(user_document);
                }
                Ok("ResFile32") => {
                    res_file_32_path = match qouted_value("\"").parse_next(&mut input) {
                        Ok(res_file_32_path) => Some(res_file_32_path),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("IconForm") => {
                    icon_form = match qouted_value("\"").parse_next(&mut input) {
                        Ok(icon_form) => Some(icon_form),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("Startup") => {
                    startup = match qouted_value("\"").parse_next(&mut input) {
                        Ok(startup) => Some(startup),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("HelpFile") => {
                    help_file_path = match qouted_value("\"").parse_next(&mut input) {
                        Ok(help_file_path) => Some(help_file_path),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("Title") => {
                    // it's perfectly possible to use '"' within the title string.
                    // VB6 being the language it is, there is no escape sequence for
                    // this. Instead, the title is wrapped in quotes and the quotes
                    // are just simply included in the text. This means we can't use
                    // the qouted_value parser here.
                    title = match title_parse.parse_next(&mut input) {
                        Ok(title) => Some(title),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("ExeName32") => {
                    exe_32_file_name = match qouted_value("\"").parse_next(&mut input) {
                        Ok(exe_32_file_name) => Some(exe_32_file_name),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("Path32") => {
                    path_32 = match qouted_value("\"").parse_next(&mut input) {
                        Ok(path_32) => Some(path_32),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("Command32") => {
                    command_line_arguments = match qouted_value("\"").parse_next(&mut input) {
                        Ok(command_line_arguments) => Some(command_line_arguments),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("Name") => {
                    name = match qouted_value("\"").parse_next(&mut input) {
                        Ok(name) => Some(name),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("Description") => {
                    description = match qouted_value("\"").parse_next(&mut input) {
                        Ok(description) => Some(description),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("HelpContextID") => {
                    help_context_id = match qouted_value("\"").parse_next(&mut input) {
                        Ok(help_context_id) => Some(help_context_id),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("CompatibleMode") => {
                    compatible_mode = match qouted_true_false_parse.parse_next(&mut input) {
                        Ok(compatible_mode) => compatible_mode,
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("MajorVer") => {
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
                }
                Ok("MinorVer") => {
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
                }
                Ok("RevisionVer") => {
                    let Ok(revision_ver): VB6Result<_> =
                        take_until_line_ending.parse_next(&mut input)
                    else {
                        return Err(input.error(VB6ErrorKind::RevisionVersionUnparseable));
                    };

                    revision = match revision_ver.to_string().as_str().parse::<u16>() {
                        Ok(revision) => revision,
                        Err(_) => {
                            return Err(input.error(VB6ErrorKind::RevisionVersionUnparseable));
                        }
                    };
                }
                Ok("AutoIncrementVer") => {
                    let Ok(auto_increment): VB6Result<_> =
                        take_until_line_ending.parse_next(&mut input)
                    else {
                        return Err(input.error(VB6ErrorKind::AutoIncrementUnparseable));
                    };

                    auto_increment_revision =
                        match auto_increment.to_string().as_str().parse::<u16>() {
                            Ok(auto_increment_revision) => auto_increment_revision,
                            Err(_) => {
                                return Err(input.error(VB6ErrorKind::AutoIncrementUnparseable));
                            }
                        };
                }
                Ok("NoControlUpgrade") => {
                    // Invert answer since we inverted the name.
                    // This defaults to true, and is the most common value.
                    upgrade_activex_controls = match true_false_parse.parse_next(&mut input) {
                        Ok(inv_upgrade_activex_controls) => !inv_upgrade_activex_controls,
                        Err(_) => {
                            return Err(input.error(VB6ErrorKind::NoControlUpgradeUnparsable));
                        }
                    };
                }
                Ok("ServerSupportFiles") => {
                    server_support_files = match true_false_parse.parse_next(&mut input) {
                        Ok(server_support_files) => server_support_files,
                        Err(_) => {
                            return Err(input.error(VB6ErrorKind::ServerSupportFilesUnparseable));
                        }
                    }
                }
                Ok("VersionCompanyName") => {
                    company_name = match qouted_value("\"").parse_next(&mut input) {
                        Ok(company_name) => Some(company_name),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("VersionFileDescription") => {
                    file_description = match qouted_value("\"").parse_next(&mut input) {
                        Ok(file_description) => Some(file_description),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("VersionLegalCopyright") => {
                    copyright = match qouted_value("\"").parse_next(&mut input) {
                        Ok(copyright) => Some(copyright),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("VersionLegalTrademarks") => {
                    trademark = match qouted_value("\"").parse_next(&mut input) {
                        Ok(trademark) => Some(trademark),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("VersionProductName") => {
                    product_name = match qouted_value("\"").parse_next(&mut input) {
                        Ok(product_name) => Some(product_name),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("VersionComments") => {
                    comments = match qouted_value("\"").parse_next(&mut input) {
                        Ok(comments) => Some(comments),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("CondComp") => {
                    conditional_compile = match qouted_value("\"").parse_next(&mut input) {
                        Ok(conditional_compile) => Some(conditional_compile),
                        Err(e) => return Err(input.error(e.into_inner().unwrap())),
                    };
                }
                Ok("CompilationType") => {
                    compilation_type = match true_false_parse.parse_next(&mut input) {
                        Ok(compilation_type) => compilation_type,
                        Err(_) => return Err(input.error(VB6ErrorKind::CompilationTypeUnparseable)),
                    };
                }
                Ok("OptimizationType") => {
                    optimization_type = match true_false_parse.parse_next(&mut input) {
                        Ok(optimization_type) => optimization_type,
                        Err(_) => {
                            return Err(input.error(VB6ErrorKind::OptimizationTypeUnparseable))
                        }
                    };
                }
                Ok("FavorPentiumPro(tm)") => {
                    favor_pentium_pro = match true_false_parse.parse_next(&mut input) {
                        Ok(favor_pentium_pro) => favor_pentium_pro,
                        Err(_) => return Err(input.error(VB6ErrorKind::FavorPentiumProUnparseable)),
                    };
                }
                Ok("CodeViewDebugInfo") => {
                    code_view_debug_info = match true_false_parse.parse_next(&mut input) {
                        Ok(code_view_debug_info) => code_view_debug_info,
                        Err(_) => {
                            return Err(input.error(VB6ErrorKind::CodeViewDebugInfoUnparseable))
                        }
                    };
                }
                Ok("NoAliasing") => {
                    // Invert answer since we inverted the name.
                    aliasing = match true_false_parse.parse_next(&mut input) {
                        Ok(inv_aliasing) => !inv_aliasing,
                        Err(_) => return Err(input.error(VB6ErrorKind::NoAliasingUnparseable)),
                    };
                }
                Ok("BoundsCheck") => {
                    bounds_check = match true_false_parse.parse_next(&mut input) {
                        Ok(bounds_check) => bounds_check,
                        Err(_) => return Err(input.error(VB6ErrorKind::BoundsCheckUnparseable)),
                    };
                }
                Ok("OverflowCheck") => {
                    overflow_check = match true_false_parse.parse_next(&mut input) {
                        Ok(overflow_check) => overflow_check,
                        Err(_) => return Err(input.error(VB6ErrorKind::OverflowCheckUnparseable)),
                    };
                }
                Ok("FlPointCheck") => {
                    floating_point_check = match true_false_parse.parse_next(&mut input) {
                        Ok(floating_point_check) => floating_point_check,
                        Err(_) => return Err(input.error(VB6ErrorKind::FlPointCheckUnparseable)),
                    };
                }
                Ok("FDIVCheck") => {
                    pentium_fdiv_bug_check = match true_false_parse.parse_next(&mut input) {
                        Ok(pentium_fdiv_bug_check) => pentium_fdiv_bug_check,
                        Err(_) => return Err(input.error(VB6ErrorKind::FDIVCheckUnparseable)),
                    };
                }
                Ok("UnroundedFP") => {
                    unrounded_floating_point = match true_false_parse.parse_next(&mut input) {
                        Ok(unrounded_floating_point) => unrounded_floating_point,
                        Err(_) => return Err(input.error(VB6ErrorKind::UnroundedFPUnparseable)),
                    };
                }
                Ok("StartMode") => {
                    start_mode = match true_false_parse.parse_next(&mut input) {
                        Ok(start_mode) => start_mode,
                        Err(_) => return Err(input.error(VB6ErrorKind::StartModeUnparseable)),
                    };
                }
                Ok("Unattended") => {
                    unattended = match true_false_parse.parse_next(&mut input) {
                        Ok(unattended) => unattended,
                        Err(_) => return Err(input.error(VB6ErrorKind::UnattendedUnparseable)),
                    };
                }
                Ok("Retained") => {
                    retained = match true_false_parse.parse_next(&mut input) {
                        Ok(retained) => retained,
                        Err(_) => return Err(input.error(VB6ErrorKind::RetainedUnparseable)),
                    };
                }
                Ok("ThreadPerObject") => {
                    let Ok(threads): VB6Result<_> = take_until_line_ending.parse_next(&mut input)
                    else {
                        return Err(input.error(VB6ErrorKind::ThreadPerObjectUnparseable));
                    };

                    thread_per_object = match threads.to_string().as_str().parse::<u16>() {
                        Ok(thread_per_object) => thread_per_object,
                        Err(_) => return Err(input.error(VB6ErrorKind::ThreadPerObjectUnparseable)),
                    }
                }
                Ok("MaxNumberOfThreads") => {
                    let Ok(max_threads): VB6Result<_> =
                        take_until_line_ending.parse_next(&mut input)
                    else {
                        return Err(input.error(VB6ErrorKind::MaxThreadsUnparseable));
                    };

                    max_number_of_threads = match max_threads.to_string().as_str().parse::<u16>() {
                        Ok(max_number_of_threads) => max_number_of_threads,
                        Err(_) => return Err(input.error(VB6ErrorKind::MaxThreadsUnparseable)),
                    };
                }
                Ok("DebugStartupOption") => {
                    debug_startup_option = match true_false_parse.parse_next(&mut input) {
                        Ok(debug_startup_option) => debug_startup_option,
                        Err(_) => {
                            return Err(input.error(VB6ErrorKind::DebugStartupOptionUnparseable))
                        }
                    };
                }
                Ok("AutoRefresh") => {
                    auto_refresh = match auto_refresh_parse.parse_next(&mut input) {
                        Ok(auto_refresh) => auto_refresh,
                        Err(_) => return Err(input.error(VB6ErrorKind::AutoRefreshUnparseable)),
                    };
                }
                _ => {
                    return Err(input.error(VB6ErrorKind::LineTypeUnknown));
                }
            }

            if (space0, alt((line_ending, line_comment_parse)))
                .parse_next(&mut input)
                .is_err()
            {
                return Err(input.error(VB6ErrorKind::NoLineEnding));
            }
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
            upgrade_activex_controls,
            res_file_32_path,
            icon_form,
            startup,
            help_file_path,
            title,
            exe_32_file_name,
            path_32,
            command_line_arguments,
            name,
            description,
            help_context_id,
            compatible_mode,
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
            debug_startup_option,
            auto_refresh,
        };

        Ok(project)
    }
}

fn true_false_parse(input: &mut VB6Stream<'_>) -> VB6Result<bool> {
    // 0 is false...and -1 is true.
    // Why vb6? What are you like this? Who hurt you?
    let result = alt(('0'.value(false), "-1".value(true), "1".value(true))).parse_next(input)?;

    Ok(result)
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

fn auto_refresh_parse(input: &mut VB6Stream<'_>) -> VB6Result<bool> {
    // 0 is false...and 1 is true.
    // Of course, VB6 being VB6, this is the only entry that does something different.
    // le sigh.
    let result = alt(('0'.value(false), "1".value(true))).parse_next(input)?;

    Ok(result)
}

fn qouted_true_false_parse(input: &mut VB6Stream<'_>) -> VB6Result<bool> {
    let qoute = qouted_value("\"").parse_next(input)?;

    // 0 is false...and -1 is true.
    // Why vb6? What are you like this? Who hurt you?
    if qoute == "0" {
        Ok(false)
    } else if qoute == "-1" {
        Ok(true)
    } else if qoute == "1" {
        Ok(true)
    } else {
        Err(ErrMode::Cut(
            VB6ErrorKind::TrueFalseOneZeroNegOneUnparseable,
        ))
    }
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

fn reference_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6ProjectReference<'a>> {
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

    let reference = VB6ProjectReference {
        uuid,
        unknown1,
        unknown2,
        path,
        description,
    };

    Ok(reference)
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
    let Ok(project_type) = alt::<_, CompileTargetType, VB6ErrorKind, _>((
        "Exe".value(CompileTargetType::Exe),
        "Control".value(CompileTargetType::Control),
        "OleExe".value(CompileTargetType::OleExe),
        "OleDll".value(CompileTargetType::OleDll),
    ))
    .parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::ProjectTypeUnknown));
    };

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
        assert_eq!(result.uuid, expected_uuid);
        assert_eq!(result.unknown1, "c.0");
        assert_eq!(result.unknown2, "0");
        assert_eq!(result.path, r"..\DBCommon\Libs\VbIntellisenseFix.dll");
        assert_eq!(result.description, r"VbIntellisenseFix");
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
}
