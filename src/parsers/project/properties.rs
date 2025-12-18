//! Defines the `ProjectProperties` struct and related enums for VB6 project properties.
//! Contains settings such as startup mode, threading model, version information, and more.
//!

use num_enum::TryFromPrimitive;
use serde::Serialize;
use strum_macros::{EnumIter, EnumMessage};

use crate::parsers::project::compilesettings::CompilationType;

/// Represents the properties of a VB6 project.
///
/// # Examples
/// ```rust
/// use vb6parse::parsers::project::properties::{ProjectProperties, StartMode, InteractionMode};
/// let project_props = ProjectProperties {
///     unused_control_info: Default::default(),
///     upgrade_controls: Default::default(),
///     res_file_32_path: "path/to/resource.res",
///     icon_form: "Form1",
///     startup: "Sub Main",
///     help_file_path: "helpfile.chm",
///     title: "My VB6 Project",
///     exe_32_file_name: "MyApp.exe",
///     exe_32_compatible: "Yes",
///     dll_base_address: 4194304,
///     path_32: "C:\\MyProject",
///     command_line_arguments: "",
///     name: "MyProject",
///     description: "A sample VB6 project",
///     debug_startup_component: "Component1",
///     help_context_id: "1000",
///     compatibility_mode: Default::default(),
///     version_32_compatibility: "Yes",
///     version_info: Default::default(),
///     server_support_files: Default::default(),
///     conditional_compile: "",
///     compilation_type: Default::default(),
///     start_mode: StartMode::StandAlone,
///     unattended: InteractionMode::Interactive,
///     retained: Default::default(),
///     thread_per_object: 1,
///     threading_model: Default::default(),
///     max_number_of_threads: 1,
///     debug_startup_option: Default::default(),
///     use_existing_browser: Default::default(),
/// };
///
/// assert_eq!(project_props.name, "MyProject");
/// assert_eq!(project_props.start_mode, StartMode::StandAlone);
/// ```
#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, Default)]
pub struct ProjectProperties<'a> {
    /// Determines whether to retain or remove licensing information for unused `ActiveX` Controls.
    pub unused_control_info: UnusedControlInfo,
    /// Determines if `ActiveX` controls should be upgraded in the project.
    pub upgrade_controls: UpgradeControls,
    /// Path to the 32-bit resource file.
    pub res_file_32_path: &'a str,
    /// The form used for the project icon.
    pub icon_form: &'a str,
    /// The startup procedure or object of the project.
    pub startup: &'a str,
    /// The help file path for the project.
    pub help_file_path: &'a str,
    /// The title of the project.
    pub title: &'a str,
    /// The name of the executable file for 32-bit projects.
    pub exe_32_file_name: &'a str,
    /// Indicates if the executable is compatible with 32-bit systems.
    pub exe_32_compatible: &'a str,
    /// The base address for the DLL in 32-bit projects.
    pub dll_base_address: u32,
    /// The path for 32-bit projects.
    pub path_32: &'a str,
    /// Command line arguments for the project.
    pub command_line_arguments: &'a str,
    /// The name of the project.
    pub name: &'a str,
    /// The description of the project.
    pub description: &'a str,
    /// The startup component for debugging.
    pub debug_startup_component: &'a str,
    /// The help context ID for the project.
    ///
    /// Note: This may need to be changed to a u32 in the future.
    pub help_context_id: &'a str,
    /// The compatibility mode of the project.
    pub compatibility_mode: CompatibilityMode,
    /// Indicates if the project is compatible with 32-bit systems.
    pub version_32_compatibility: &'a str,
    /// The version information of the project.
    pub version_info: VersionInformation<'a>,
    /// Indicates if the project will produce server support files.
    pub server_support_files: ServerSupportFiles,
    /// The conditional compile settings for the project.
    pub conditional_compile: &'a str,
    /// The compilation type settings for the project.
    pub compilation_type: CompilationType,
    /// The start mode of the project.
    pub start_mode: StartMode,
    /// The interaction mode of the project.
    pub unattended: InteractionMode,
    /// The retained mode of the project.
    pub retained: Retained,
    /// The number of threads per object.
    pub thread_per_object: u16,
    /// The threading model of the project.
    pub threading_model: ThreadingModel,
    /// The maximum number of threads for the project.
    pub max_number_of_threads: u16,
    /// The debug startup option for the project.
    pub debug_startup_option: DebugStartupOption,
    /// Indicates whether to use an existing browser instance.
    pub use_existing_browser: ExistingBrowser,
}

/// Retained mode of the VB6 project.
///
/// Hints to the loading program whether the project DLL should be retained in
/// memory or unloaded when no longer in use.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum Retained {
    /// The DLL is unloaded when no longer in use.
    #[default]
    #[strum(message = "Unload the DLL when no longer in use")]
    UnloadOnExit = 0,
    /// `RetainedInMemory` only indicates to the loading program that the DLL
    /// should be retained in memory, it does not guarantee that the DLL will be
    /// retained in memory. Retaining a DLL in memory comes with a memory and
    /// performance cost that the host program may not wish to sustain.
    #[strum(message = "Retain the DLL in memory")]
    RetainedInMemory = 1,
}

impl TryFrom<&str> for Retained {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(Retained::UnloadOnExit),
            b"1" => Ok(Retained::RetainedInMemory),
            _ => Err(format!("Unknown Retained value: '{value}'")),
        }
    }
}

/// Indicates whether to use an existing browser instance.
///
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum ExistingBrowser {
    /// Do not use an existing browser instance.
    #[strum(message = "Do not use an existing browser instance")]
    DoNotUse = 0,
    /// If Internet Explorer is already running, use the existing instance.
    /// Otherwise, launch a new instance.
    #[default]
    #[strum(message = "Use an existing browser instance if available")]
    Use = -1,
}

impl TryFrom<&str> for ExistingBrowser {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(ExistingBrowser::DoNotUse),
            b"-1" => Ok(ExistingBrowser::Use),
            _ => Err(format!("Unknown ExistingBrowser value: '{value}'")),
        }
    }
}

/// Start mode of the VB6 project.
///
/// Indicates whether the project is a stand-alone application or an `ActiveX`
/// component.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum StartMode {
    /// The project is a stand-alone application. Essentially, a normal EXE.
    #[default]
    #[strum(message = "Stand-alone Application")]
    StandAlone = 0,
    /// The project is an `ActiveX` component.
    #[strum(message = "ActiveX Automation Component")]
    Automation = 1,
}

impl TryFrom<&str> for StartMode {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(StartMode::StandAlone),
            b"1" => Ok(StartMode::Automation),
            _ => Err(format!("Unknown StartMode value: '{value}'")),
        }
    }
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
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum InteractionMode {
    /// The program can show dialogs and interacts with the user.
    #[default]
    #[strum(message = "The program can show dialogs and interacts with the user.")]
    Interactive = 0,
    /// The program cannot show dialogs and will not interact with the user.
    #[strum(message = "The program cannot show dialogs and will not interact with the user.")]
    Unattended = -1,
}

impl TryFrom<&str> for InteractionMode {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(InteractionMode::Interactive),
            b"-1" => Ok(InteractionMode::Unattended),
            _ => Err(format!("Unknown InteractionMode value: '{value}'")),
        }
    }
}

/// Indicates if the project will produce the server support VBR & TLB files.
///
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum ServerSupportFiles {
    /// The project will not produce the VBR and TLB files since this is a
    /// local only project.
    #[default]
    #[strum(message = "Do not produce VBR and TLB files")]
    Local = 0,
    /// The project will produce the VBR and TLB files needed for packaging
    /// the client applications which use this server project.
    #[strum(message = "Produce VBR and TLB files for packaging")]
    Remote = 1,
}

impl TryFrom<&str> for ServerSupportFiles {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(ServerSupportFiles::Local),
            b"1" => Ok(ServerSupportFiles::Remote),
            _ => Err(format!("Unknown ServerSupportFiles value: '{value}'")),
        }
    }
}

/// If the `ActiveX` control has been updated in windows since the last time
/// the project was opened this setting determines if the project should
/// be updated to use the new control or not.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum UpgradeControls {
    /// The project should be updated to use the new control.
    #[strum(message = "Update project to use upgraded control")]
    #[default]
    Upgrade = 0,
    /// The project should not be updated to use the new control.
    #[strum(message = "Leave project untouched with older control")]
    NoUpgrade = 1,
}

impl TryFrom<&str> for UpgradeControls {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(UpgradeControls::Upgrade),
            b"1" => Ok(UpgradeControls::NoUpgrade),
            _ => Err(format!("Unknown UpgradeControls value: '{value}'")),
        }
    }
}

/// Determines if licensing information for `ActiveX` Controls unused, but
/// referenced within the project, should be retained or removed.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum UnusedControlInfo {
    /// The licensing information for `ActiveX` Controls unused, but referenced
    /// within the project, should be retained.
    #[strum(message = "Retain License Information")]
    Retain = 0,
    /// The licensing information for `ActiveX` Controls unused, but referenced
    /// within the project, should be removed.
    #[default]
    #[strum(message = "Remove License Information")]
    Remove = 1,
}

impl TryFrom<&str> for UnusedControlInfo {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(UnusedControlInfo::Retain),
            b"1" => Ok(UnusedControlInfo::Remove),
            _ => Err(format!("Unknown UnusedControlInfo value: '{value}'")),
        }
    }
}

/// Determines the level of compatibility required for each compile of the project.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum CompatibilityMode {
    /// Each time the component is compiled, new type library information is
    /// generated, including new class ID's and new interface ID's. There is no
    /// relation between versions of a component, and programs compiled to use
    /// one version cannot use another subsequent version.
    #[strum(message = "No Compatibility")]
    NoCompatibility = 0,
    /// Each time the component is compiled, the type library identifier is kept,
    /// so that projects can maintain their references to the component. All class
    /// ID's from the previous version are maintained. Interface Id's are changed
    /// only for classes that are no longer binary-compatible with their earlier
    /// counterparts.
    #[default]
    #[strum(message = "Project Compatibility")]
    Project = 1,
    /// When the component is compiled, if any binary-incompatible changes are
    /// detected, the IDE will present a warning dialog. If accepted, the component
    /// will retain the type library identifier and the class ID's from the previous
    /// version. The interface ID's are changed only for classes that are no longer
    /// binary-compatible with their earlier counterparts. Otherwise, the component
    /// will maintain the type library identifier and the class ID's from the previous
    /// version regardless of whether the changes are binary-compatible or not.
    #[strum(message = "Compatible Exe Mode")]
    CompatibleExe = 2,
}

impl TryFrom<&str> for CompatibilityMode {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(CompatibilityMode::NoCompatibility),
            b"1" => Ok(CompatibilityMode::Project),
            b"2" => Ok(CompatibilityMode::CompatibleExe),
            _ => Err(format!("Unknown CompatibilityMode value: '{value}'")),
        }
    }
}

/// When debugging the VB6 project, this option determines how the
/// debugging session will start.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum DebugStartupOption {
    /// When debugging, the IDE will wait for the component to be created before
    /// attaching the debugging session.
    #[default]
    #[strum(message = "Wait for Component Creation")]
    WaitForComponentCreation = 0,
    /// Lets the component determine what happens. The types of components include
    /// `ActiveX` Designers like the DHTML Designer and the `WebClass` Designer, and
    /// also User Controls and User Documents. If a User Control or User Document
    /// is selected, Visual Basic will launch the browser and display a dummy test
    /// page that contains the component. The component can tell Visual Basic to
    /// either launch the browser with a URL or start another program.
    ///
    /// Selecting a startup component on the Debugging tab does not effect the
    /// Startup Object specified on the General tab.
    ///
    /// For example: an ActiveX.dll project could specify 'Startup Object=Sub Main'
    /// and 'Start Component=DHTMLPage1'.
    ///
    /// When the project runs, Visual Basic would register the `DHTMLPage1` component
    /// as well as other components, execute and then launch Internet Explorer,
    /// and navigate to a URL that create an instance of `DHTMLPage1`.
    #[strum(message = "A Start Component is specified")]
    StartComponent = 1,
    /// Specifies an executable program to be used.
    #[strum(message = "A Start Program is specified")]
    StartProgram = 2,
    /// Specifies which URL the browser should navigate to.
    #[strum(message = "Start Browser is specified")]
    StartBrowser = 3,
}

impl TryFrom<&str> for DebugStartupOption {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value.as_bytes() {
            b"0" => Ok(DebugStartupOption::WaitForComponentCreation),
            b"1" => Ok(DebugStartupOption::StartComponent),
            b"2" => Ok(DebugStartupOption::StartProgram),
            b"3" => Ok(DebugStartupOption::StartBrowser),
            _ => Err(format!("Unknown DebugStartupOption value: '{value}'")),
        }
    }
}

/// Determines the version information of the VB6 project.
#[derive(Debug, PartialEq, Eq, Default, Copy, Clone, Serialize)]
pub struct VersionInformation<'a> {
    /// The major version number of the project.
    pub major: u16,
    /// The minor version number of the project.
    pub minor: u16,
    /// The revision version number of the project.
    pub revision: u16,
    /// How much the revision number should be incremented per compile.
    pub auto_increment_revision: u16,
    /// The name of the company that created the project.
    pub company_name: &'a str,
    /// The description of the project.
    pub file_description: &'a str,
    /// The copyright information of the project.
    pub copyright: &'a str,
    /// The trademark information of the project.
    pub trademark: &'a str,
    /// The product name of the project.
    pub product_name: &'a str,
    /// Additional comments about the project.
    pub comments: &'a str,
}

/// Determines the type of compile target for the VB6 project.
#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, EnumIter, EnumMessage, Default)]
pub enum CompileTargetType {
    /// The project is a standard EXE.
    #[strum(message = "A Standard Exe")]
    #[default]
    Exe,
    /// The project is a `UserControl`.
    #[strum(message = "A UserControl")]
    Control,
    /// The project is a OLE Executable.
    #[strum(message = "Ole Exe")]
    OleExe,
    /// The project is an OLE DLL.
    #[strum(message = "Ole Dll")]
    OleDll,
}

impl TryFrom<&str> for CompileTargetType {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "Exe" => Ok(CompileTargetType::Exe),
            "Control" => Ok(CompileTargetType::Control),
            "OleExe" => Ok(CompileTargetType::OleExe),
            "OleDll" => Ok(CompileTargetType::OleDll),
            _ => Err(format!("Unknown CompileTargetType value: '{value}'")),
        }
    }
}

/// Determines the threading model for the VB6 project.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum ThreadingModel {
    /// Single-threaded.
    #[strum(message = "")]
    SingleThreaded = 0,
    /// Apartment-threaded.
    #[default]
    #[strum(message = "")]
    ApartmentThreaded = 1,
}

impl TryFrom<&str> for ThreadingModel {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(ThreadingModel::SingleThreaded),
            "1" => Ok(ThreadingModel::ApartmentThreaded),
            _ => Err(format!("Unknown ThreadingModel value: '{value}'")),
        }
    }
}
