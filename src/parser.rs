#![warn(clippy::pedantic)]

use std::str;

use nom::{
    branch::alt,
    bytes::complete::{tag_no_case, take_until, take_until1},
    character::complete::{line_ending, not_line_ending},
    combinator::{eof, not, peek, value},
    error::{ErrorKind, ParseError},
    multi::many0,
    sequence::{pair, preceded, tuple},
    IResult,
};

use uuid::Uuid;

// These constants are used as text to tag capture against. Sadly, because vb6
// predates UTF we have to read the project file in as a byte slice since
// it can contain non-ascii text elements, especially the tm character, the
// copyright character and other such 'special' characters which are often found
// in the VersionLegalCopyright, VersionLegalDescription, etc fields.
const REFERENCE: &[u8] = b"Reference";
const OBJECT: &[u8] = b"Object";
const MODULE: &[u8] = b"Module";
const DESIGNER: &[u8] = b"Designer";
const USERDOCUMENT: &[u8] = b"UserDocument";
const CLASS: &[u8] = b"Class";
const FORM: &[u8] = b"Form";
const USERCONTROL: &[u8] = b"UserControl";
const RESFILE32: &[u8] = b"ResFile32";
const ICONFORM: &[u8] = b"IconForm";
const STARTUP: &[u8] = b"Startup";
const HELPFILE: &[u8] = b"HelpFile";
const TITLE: &[u8] = b"Title";
const EXENAME32: &[u8] = b"ExeName32";
const COMMAND32: &[u8] = b"Command32";
const NAME: &[u8] = b"Name";
const HELPCONTEXTID: &[u8] = b"HelpContextID";
const COMPATIBLEMODE: &[u8] = b"CompatibleMode";
const NOCONTROLUPGRADE: &[u8] = b"NoControlUpgrade";
const MAJORVER: &[u8] = b"MajorVer";
const MINORVER: &[u8] = b"MinorVer";
const REVISIONVER: &[u8] = b"RevisionVer";
const AUTOINCREMENTVER: &[u8] = b"AutoIncrementVer";
const SERVERSUPPORTFILES: &[u8] = b"ServerSupportFiles";
const VERSIONCOMPANYNAME: &[u8] = b"VersionCompanyName";
const VERSIONFILEDESCRIPTION: &[u8] = b"VersionFileDescription";
const VERSIONLEGALCOPYRIGHT: &[u8] = b"VersionLegalCopyright";
const VERSIONLEGALTRADEMARKS: &[u8] = b"VersionLegalTrademarks";
const VERSIONPRODUCTNAME: &[u8] = b"VersionProductName";
const CONDCOMP: &[u8] = b"CondComp";
const COMPILATIONTYPE: &[u8] = b"CompilationType";
const OPTIMIZATIONTYPE: &[u8] = b"OptimizationType";
const NOALIASING: &[u8] = b"NoAliasing";
const CODEVIEWDEBUGINFO: &[u8] = b"CodeViewDebugInfo";
// In the vbp file this is FavorPentiumPro(tm)
const FAVORPENTIUMPROTM: &[u8] = b"FavorPentiumPro(tm)";
const BOUNDSCHECK: &[u8] = b"BoundsCheck";
const OVERFLOWCHECK: &[u8] = b"OverflowCheck";
const FLPOINTCHECK: &[u8] = b"FlPointCheck";
const FDIVCHECK: &[u8] = b"FDIVCheck";
const UNROUNDEDFP: &[u8] = b"UnroundedFP";
const STARTMODE: &[u8] = b"StartMode";
const UNATTENDED: &[u8] = b"Unattended";
const RETAINED: &[u8] = b"Retained";
const THREADPEROBJECT: &[u8] = b"ThreadPerObject";
const MAXNUMBEROFTHREADS: &[u8] = b"MaxNumberOfThreads";
const DEBUGSTARTOPTION: &[u8] = b"DebugStartOption";
const AUTOREFRESH: &[u8] = b"AutoRefresh";

const EMPTY: &[u8] = b"";

#[derive(thiserror::Error, Debug, PartialEq)]
pub enum ProjectParseError {
    #[error("Line type is unknown.")]
    LineTypeUnknown,
    #[error("Project type is not Exe or OleDll")]
    ProjectTypeUnknown,
    #[error("Project line entry is not ended with a recognized line ending.")]
    NoLineEnding,
    #[error("Unable to parse the Uuid")]
    UnableToParseUuid,
    #[error("Unable to find a semicolon ';' in this line.")]
    NoSemicolonSplit,
    #[error("Unable to find an equal '=' in this line.")]
    NoEqualSplit,
    #[error("Unknown parser error")]
    Unparseable,
}

impl<I> ParseError<I> for ProjectParseError {
    fn from_error_kind(_input: I, _kind: ErrorKind) -> Self {
        ProjectParseError::Unparseable
    }

    fn append(_: I, _: ErrorKind, other: Self) -> Self {
        other
    }
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Project {
    pub project_type: ProjectType,
    pub references: Vec<VB6ProjectReference>,
    pub objects: Vec<VB6ProjectObject>,
    pub modules: Vec<VB6ProjectModule>,
    pub classes: Vec<VB6ProjectClass>,
    pub designers: Vec<String>,
    pub forms: Vec<String>,
    pub user_controls: Vec<String>,
    pub user_documents: Vec<String>,
    pub upgrade_activex_controls: bool,
    pub res_file_32_path: String,
    pub icon_form: Option<String>,
    pub startup: String,
    pub help_file_path: String,
    pub title: String,
    pub exe_32_file_name: String,
    pub command_line_arguments: String,
    pub name: String,
    // May need to be switched to a u32. Not sure yet.
    pub help_context_id: String,
    pub compatible_mode: String,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum ProjectType {
    Exe,
    Control,
    OleExe,
    OleDll,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ProjectReference {
    pub uuid: Uuid,
    pub unknown1: String,
    pub unknown2: String,
    pub path: String,
    pub description: String,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ProjectObject {
    pub uuid: Uuid,
    pub version: String,
    pub unknown1: String,
    pub file_name: String,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ProjectModule {
    pub name: String,
    pub path: String,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ProjectClass {
    pub name: String,
    pub path: String,
}

#[derive(Debug, PartialEq, Eq, Clone)]
enum LineType {
    Reference(VB6ProjectReference),
    UserDocument(String),
    Object(VB6ProjectObject),
    Module(VB6ProjectModule),
    Designer(String),
    Class(VB6ProjectClass),
    Form(String),
    UserControl(String),
    ResFile32(String),
    IconForm(String),
    Startup(String),
    HelpFile(String),
    Title(String),
    ExeName32(String),
    Command32(String),
    Name(String),
    HelpContextID(String),
    CompatibleMode(String),
    NoControlUpgrade(String), // 0 or line missing - false, 1 = 'Upgrade ActiveX Control'. default = true.
    MajorVer,                 // 0 - 9999, default 1.
    MinorVer,                 // 0 - 9999, default 0.
    RevisionVer,              // 0 - 9999, default 0.
    AutoIncrementVer,         // 0 - no increment, 1 - increment, default 0.
    ServerSupportFiles,
    VersionCompanyName,
    VersionFileDescription,
    VersionLegalCopyright,
    VersionLegalTrademarks,
    VersionProductName,
    CondComp,
    CompilationType,
    OptimizationType,
    /// In the file this is FavorPentiumPro(tm)
    FavorPentiumProTM,
    CodeViewDebugInfo,
    NoAliasing,
    BoundsCheck,
    FlPointCheck,
    FDIVCheck,
    UnroundedFP,
    StartMode,
    Unattended,
    Retained,
    ThreadPerObject,
    MaxNumberOfThreads,
    DebugStartOption,
    AutoRefresh,
    Empty,
}

impl VB6Project {
    pub fn parse(input: &[u8]) -> Result<Self, ProjectParseError> {
        let remainder = input;

        let Ok((remainder, project_type)) = project_type_parse(remainder) else {
            return Err(ProjectParseError::ProjectTypeUnknown);
        };

        let (_remainder, line_types) = many0(preceded(not(eof), line_type_parse))(remainder)
            .map_err(|_| ProjectParseError::NoLineEnding)?;

        let references = line_types
            .iter()
            .filter_map(|line| match line {
                LineType::Reference(reference) => Some(reference.clone()),
                _ => None,
            })
            .collect();

        let user_documents = line_types
            .iter()
            .filter_map(|line| match line {
                LineType::UserDocument(user_document) => Some(user_document.clone()),
                _ => None,
            })
            .collect();

        let objects = line_types
            .iter()
            .filter_map(|line| match line {
                LineType::Object(object) => Some(object.clone()),
                _ => None,
            })
            .collect();

        let modules = line_types
            .iter()
            .filter_map(|line| match line {
                LineType::Module(module) => Some(module.clone()),
                _ => None,
            })
            .collect();

        let classes = line_types
            .iter()
            .filter_map(|line| match line {
                LineType::Class(class) => Some(class.clone()),
                _ => None,
            })
            .collect();

        let designers = line_types
            .iter()
            .filter_map(|line| match line {
                LineType::Designer(designer) => Some(designer.clone()),
                _ => None,
            })
            .collect();

        let forms = line_types
            .iter()
            .filter_map(|line| match line {
                LineType::Form(form) => Some(form.clone()),
                _ => None,
            })
            .collect();

        let user_controls = line_types
            .iter()
            .filter_map(|line| match line {
                LineType::UserControl(user_control) => Some(user_control.clone()),
                _ => None,
            })
            .collect();

        // TODO:
        // All of these should have a default value that matches whatever
        // default VB6 uses whenever the item isn't in the VB6 project (*.vbp)
        // file. For now, I've just done an 'unwrap', this is not right, but
        // we should be able to come back to this later.
        let res_file_32_path = line_types
            .iter()
            .find_map(|line| match line {
                LineType::ResFile32(res_file_32_path) => Some(res_file_32_path.clone()),
                _ => None,
            })
            .unwrap();

        let icon_form = line_types.iter().find_map(|line| match line {
            LineType::IconForm(icon_form) => Some(icon_form.clone()),
            _ => None,
        });

        let startup = line_types
            .iter()
            .find_map(|line| match line {
                LineType::Startup(startup) => Some(startup.clone()),
                _ => None,
            })
            .unwrap();

        let help_file_path = line_types
            .iter()
            .find_map(|line| match line {
                LineType::HelpFile(help_file_path) => Some(help_file_path.clone()),
                _ => None,
            })
            .unwrap();

        let title = line_types
            .iter()
            .find_map(|line| match line {
                LineType::Title(title) => Some(title.clone()),
                _ => None,
            })
            .unwrap();

        let exe_32_file_name = line_types
            .iter()
            .find_map(|line| match line {
                LineType::ExeName32(exe_32_file_name) => Some(exe_32_file_name.clone()),
                _ => None,
            })
            .unwrap();

        let command_line_arguments = line_types
            .iter()
            .find_map(|line| match line {
                LineType::Command32(command_line_arguments) => Some(command_line_arguments.clone()),
                _ => None,
            })
            .unwrap();

        let name = line_types
            .iter()
            .find_map(|line| match line {
                LineType::Name(name) => Some(name.clone()),
                _ => None,
            })
            .unwrap();

        let help_context_id = line_types
            .iter()
            .find_map(|line| match line {
                LineType::HelpContextID(help_context_id) => Some(help_context_id.clone()),
                _ => None,
            })
            .unwrap();

        let compatible_mode = line_types
            .iter()
            .find_map(|line| match line {
                LineType::CompatibleMode(compatible_mode) => Some(compatible_mode.clone()),
                _ => None,
            })
            .unwrap();

        // If we don't have an NoControlUpgrade line then the default is
        // true. The text line itself is 1 or 0 which is translated to
        // true and false.
        let upgrade_activex_controls = line_types
            .iter()
            .find_map(|line| match line {
                LineType::NoControlUpgrade(control_upgrade) => Some(control_upgrade.clone()),
                _ => None,
            })
            .map_or(true, |value| matches!(value.as_str(), "1"));

        let mut project = VB6Project {
            project_type,
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
            command_line_arguments,
            name,
            help_context_id,
            compatible_mode,
        };

        project.validate();

        Ok(project)
    }

    fn validate(&mut self) {
        let icon = self.icon_form.clone();

        if (icon.is_none() || icon.unwrap().is_empty()) && !self.forms.is_empty() {
            let first_form_path = self.forms[0].clone();

            let form_name = file_name_without_extension(first_form_path);

            self.icon_form = form_name;
        };
    }
}

fn file_name_without_extension<P>(path: P) -> Option<String>
where
    P: AsRef<std::path::Path>,
{
    let file_name = path.as_ref().file_name()?;

    let file_limited_path = std::path::Path::new(file_name);

    // Using 'with_extension' this way is gross, but 'file_prefix' is currently
    // nightly only.
    let file_name_without_extension = file_limited_path.with_extension("");

    file_name_without_extension
        .into_os_string()
        .into_string()
        .ok()
}

fn take_line_remove_newline_parse(input: &[u8]) -> IResult<&[u8], &[u8], ProjectParseError> {
    // We specify the impl of the tag here since we want to catch a failure on
    // the alternation check and specifying it here makes the code easier to
    // read.
    let line_ending = line_ending::<&[u8], ProjectParseError>;

    let remainder = input;

    let (remainder, line) = alt((take_until("\n"), take_until("\r\n")))(remainder)?;
    let (remainder, _) = line_ending(remainder)?;

    Ok((remainder, line))
}

fn project_type_parse(input: &[u8]) -> IResult<&[u8], ProjectType, ProjectParseError> {
    // We specify the impl of the tag here since we want to catch a failure on
    // the alternation check and specifying it here makes the code easier to
    // read.
    let tag_no_case = tag_no_case::<&str, &[u8], ProjectParseError>;
    let line_ending = line_ending::<&[u8], ProjectParseError>;

    let remainder = input;

    // The first line of any VB6 project file (vbp) is a type line that
    // tells us what kind of project we have.
    // this should be in every project file, even an empty one, and it must
    // be one of these four options.
    // further, it should end with a "\r\n" to be conservative, we will accept
    // either an "\n" or an "\r\n"
    //
    // The project type line starts with a 'Type=' has either 'Exe' or 'OleDll'.
    let Ok((remainder, (_, project_type))) = pair(
        tag_no_case("Type="),
        alt((
            value(ProjectType::Exe, tag_no_case("Exe")),
            value(ProjectType::Control, tag_no_case("Control")),
            value(ProjectType::OleDll, tag_no_case("OleDll")),
            value(ProjectType::OleExe, tag_no_case("OleExe")),
        )),
    )(remainder) else {
        return Err(nom::Err::Failure(ProjectParseError::ProjectTypeUnknown));
    };

    // We split out the newline here so we can handle the difference between
    // a type line that ends in a newline and one without it.
    let Ok((remainder, _)) = line_ending(remainder) else {
        return Err(nom::Err::Failure(ProjectParseError::NoLineEnding));
    };

    Ok((remainder, project_type))
}

fn line_type_parse(input: &[u8]) -> IResult<&[u8], LineType, ProjectParseError> {
    // fully qualify the impl of these parsers here because it makes the following
    // code easier to read since we need to fully specify the parser function to
    // make error reporting easier.
    let take_until1 = take_until1::<&str, &[u8], ProjectParseError>;
    let line_ending = line_ending::<&[u8], ProjectParseError>;

    let remainder = input;

    let Ok((remainder, line_type_text)) = peek(alt((take_until1("="), line_ending)))(remainder)
    else {
        return Err(nom::Err::Failure(ProjectParseError::LineTypeUnknown));
    };

    let (remainder, line_type) = match line_type_text {
        REFERENCE => {
            let (remainder, reference) = reference_line_parse(remainder)?;

            (remainder, LineType::Reference(reference))
        }
        USERDOCUMENT => {
            let (remainder, (_key, user_document)) = key_value_pair_parse(remainder)?;

            (remainder, LineType::UserDocument(user_document))
        }
        OBJECT => {
            let (remainder, object) = object_line_parse(remainder)?;

            (remainder, LineType::Object(object))
        }
        MODULE => {
            let (remainder, module) = module_line_parse(remainder)?;

            (remainder, LineType::Module(module))
        }
        DESIGNER => {
            let (remainder, (_key, path)) = key_value_pair_parse(remainder)?;

            (remainder, LineType::Designer(path))
        }
        CLASS => {
            let (remainder, class) = class_line_parse(remainder)?;

            (remainder, LineType::Class(class))
        }
        FORM => {
            let (remainder, (_key, path)) = key_value_pair_parse(remainder)?;

            (remainder, LineType::Form(path))
        }
        USERCONTROL => {
            let (remainder, (_key, path)) = key_value_pair_parse(remainder)?;

            (remainder, LineType::UserControl(path))
        }
        RESFILE32 => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::ResFile32(value))
        }
        ICONFORM => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::IconForm(value))
        }
        STARTUP => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::Startup(value))
        }
        HELPFILE => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::HelpFile(value))
        }
        TITLE => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::Title(value))
        }
        EXENAME32 => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::ExeName32(value))
        }
        COMMAND32 => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::Command32(value))
        }
        NAME => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::Name(value))
        }
        HELPCONTEXTID => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::HelpContextID(value))
        }
        COMPATIBLEMODE => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::CompatibleMode(value))
        }
        NOCONTROLUPGRADE => {
            let (remainder, (_key, value)) = key_qouted_value_pair_parse(remainder)?;

            (remainder, LineType::NoControlUpgrade(value))
        }
        MAJORVER => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::MajorVer)
        }
        MINORVER => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::MinorVer)
        }
        REVISIONVER => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::RevisionVer)
        }
        AUTOINCREMENTVER => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::AutoIncrementVer)
        }
        SERVERSUPPORTFILES => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::ServerSupportFiles)
        }
        VERSIONCOMPANYNAME => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::VersionCompanyName)
        }
        VERSIONFILEDESCRIPTION => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::VersionFileDescription)
        }
        VERSIONLEGALCOPYRIGHT => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::VersionLegalCopyright)
        }
        VERSIONLEGALTRADEMARKS => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::VersionLegalTrademarks)
        }
        VERSIONPRODUCTNAME => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::VersionProductName)
        }
        CONDCOMP => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::CondComp)
        }
        COMPILATIONTYPE => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::CompilationType)
        }
        OPTIMIZATIONTYPE => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::OptimizationType)
        }
        NOALIASING => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::NoAliasing)
        }
        CODEVIEWDEBUGINFO => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::CodeViewDebugInfo)
        }
        // In the vbp file this is FavorPentiumPro(tm)
        FAVORPENTIUMPROTM => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::FavorPentiumProTM)
        }
        OVERFLOWCHECK => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::BoundsCheck)
        }
        BOUNDSCHECK => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::BoundsCheck)
        }
        FLPOINTCHECK => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::FlPointCheck)
        }
        FDIVCHECK => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::FDIVCheck)
        }
        UNROUNDEDFP => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::UnroundedFP)
        }
        STARTMODE => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::StartMode)
        }
        UNATTENDED => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::Unattended)
        }
        RETAINED => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::Retained)
        }
        THREADPEROBJECT => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::ThreadPerObject)
        }
        MAXNUMBEROFTHREADS => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::MaxNumberOfThreads)
        }
        DEBUGSTARTOPTION => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::DebugStartOption)
        }
        EMPTY => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::Empty)
        }
        AUTOREFRESH => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::AutoRefresh)
        }
        _ => {
            let (remainder, _) = take_line_remove_newline_parse(remainder)?;

            (remainder, LineType::Empty)
            //return Err(nom::Err::Failure(ProjectParseError::LineTypeUnknown));
        }
    };

    Ok((remainder, line_type))
}

fn key_qouted_value_pair_parse(input: &[u8]) -> IResult<&[u8], (&[u8], String), ProjectParseError> {
    let remainder = input;

    let (remainder, (key, value)) = key_value_pair_parse(remainder)?;

    // This variant uses a key/value variant has a double quoted value.
    let value = value.trim_matches('"').to_owned();

    Ok((remainder, (key, value)))
}

fn key_value_pair_parse(input: &[u8]) -> IResult<&[u8], (&[u8], String), ProjectParseError> {
    // Multiple lines are of the form 'key=value\r\n'
    // For example:
    // Form=..\DBCommon\frmSelectUser.frm\r\n
    // Designer=AllMfgStatus.Dsr\r\n
    //
    // This parser handles this by spliting on the equal and returning a
    // tuple of the two halves.

    // this parser reads right after the '=' to get just the name and path.

    // We specify the impl here because it makes the following code
    // easier to read since we need to fully specifiy the parser function to
    // make error reporting easier.
    let not_line_ending = not_line_ending::<&[u8], ProjectParseError>;
    let take_until = take_until::<&str, &[u8], ProjectParseError>;

    let remainder = input;

    let Ok((remainder, name)) = take_until(r"=")(remainder) else {
        return Err(nom::Err::Failure(ProjectParseError::NoEqualSplit));
    };

    //  We read up to the split, now we want to consume the splitter as well.
    let (remainder, _) = tag_no_case(r"=")(remainder)?;

    // Finally, we are grabbing the value.
    let Ok((remainder, value)) = not_line_ending(remainder) else {
        return Err(nom::Err::Failure(ProjectParseError::NoLineEnding));
    };

    // TODO:
    // this works on most lines and most systems, but there are some
    // key/value pairs where the value is a non-ascii and non-utf8 because
    // VB6 predates UTF. These escape sequences need to be learned and converted
    // into the correct UTF-8 format.
    let value = String::from_utf8(value.to_vec()).unwrap();

    // We snagged up to the line ending before, now we want to actually get
    // that line ending as well.
    let (remainder, _) = line_ending(remainder)?;

    Ok((remainder, (name, value)))
}

fn object_line_parse(input: &[u8]) -> IResult<&[u8], VB6ProjectObject, ProjectParseError> {
    // Object={C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0; crviewer.dll\r\n

    // a reference line starts with 'Object=' which is then followed by
    // a GUID which is:
    // S "{", followed by 8 hexadecimal digits, a "-", four hexadecimal digits,
    // a "-", another four hexadecimal digits, a "-", another twelve
    // hexadecimal digits, and finally a "}#". The '# is used to indicate the
    // start of the next section.
    // We then have a version number, another '#', an unknown value followed by
    // a semicolon and a single white space, then the file name (usually a .dll
    // or .ocx) and finally a newline.

    // We specify the impl here because it makes the following code
    // easier to read since we need to fully specifiy the parser function to
    // make error reporting easier.
    let not_line_ending = not_line_ending::<&[u8], ProjectParseError>;

    let remainder = input;

    let (remainder, (_, uuid_bytes)) =
        tuple((tag_no_case(b"Object={"), take_until("}#")))(remainder)?;

    let Ok(uuid_text) = str::from_utf8(uuid_bytes) else {
        return Err(nom::Err::Failure(ProjectParseError::UnableToParseUuid));
    };

    let Ok(uuid) = uuid::Uuid::parse_str(uuid_text) else {
        return Err(nom::Err::Failure(ProjectParseError::UnableToParseUuid));
    };

    let (remainder, _) = tag_no_case(b"}#")(remainder)?;

    // We again take until the first '#'.
    // This looks like, and almost certainly is, a version number of a
    // '#.#' form like 1.0, 2.1, 6.0, etc.
    let (remainder, version) = take_until("#")(remainder)?;

    let (remainder, _) = tag_no_case(b"#")(remainder)?;

    // We again take until the first '; '. It's not clear what this value is.
    // In every case, I've only seen '0'.
    let (remainder, unknown1) = take_until("; ")(remainder)?;

    let (remainder, _) = tag_no_case(b"; ")(remainder)?;

    // Finally, we are grabbing the file name, for this object.
    let Ok((remainder, file_name)) = not_line_ending(remainder) else {
        return Err(nom::Err::Failure(ProjectParseError::NoLineEnding));
    };

    // We snagged up to the line ending before, now we want to actually get
    // that line ending as well.
    let (remainder, _) = line_ending(remainder)?;

    let object = VB6ProjectObject {
        uuid,
        version: String::from_utf8(version.to_vec()).unwrap(),
        unknown1: String::from_utf8(unknown1.to_vec()).unwrap(),
        file_name: String::from_utf8(file_name.to_vec()).unwrap(),
    };

    Ok((remainder, object))
}

fn name_path_tuple_parse(input: &[u8]) -> IResult<&[u8], (&[u8], &[u8]), ProjectParseError> {
    // module and class lines both use a 'tag=filename; path\r\n' pattern.
    // examples:
    // Module=modDBAssist; ..\DBCommon\DBAssist.bas\r\n
    // Class=CDecodeVarsClass; ..\DBCommon\CDecodeVarsClass.cls\r\n

    // this parser reads right after the '=' to get just the name and path.

    // We specify the impl here because it makes the following code
    // easier to read since we need to fully specifiy the parser function to
    // make error reporting easier.
    let not_line_ending = not_line_ending::<&[u8], ProjectParseError>;
    let take_until = take_until::<&str, &[u8], ProjectParseError>;

    let remainder = input;

    let Ok((remainder, name)) = take_until(r"; ")(remainder) else {
        return Err(nom::Err::Failure(ProjectParseError::NoLineEnding));
    };

    //  We read up to the split, now we want to consume the splitter as well.
    let (remainder, _) = tag_no_case(r"; ")(remainder)?;

    // Finally, we are grabbing the path.
    let Ok((remainder, path)) = not_line_ending(remainder) else {
        return Err(nom::Err::Failure(ProjectParseError::NoLineEnding));
    };

    // We snagged up to the line ending before, now we want to actually get
    // that line ending as well.
    let (remainder, _) = line_ending(remainder)?;

    Ok((remainder, (name, path)))
}

fn module_line_parse(input: &[u8]) -> IResult<&[u8], VB6ProjectModule, ProjectParseError> {
    //Module=modDBAssist; ..\DBCommon\DBAssist.bas\r\n

    // a module line starts with 'Module=' which is then followed by
    // the modules name then seperated by a semicolon and a single white space,
    // then the file name (usually a .bas file) and finally a newline.

    let remainder = input;

    let (remainder, _) = tag_no_case(b"Module=")(remainder)?;

    let (remainder, (name, path)) = name_path_tuple_parse(remainder)?;

    let module = VB6ProjectModule {
        name: String::from_utf8(name.to_vec()).unwrap(),
        path: String::from_utf8(path.to_vec()).unwrap(),
    };

    Ok((remainder, module))
}

fn class_line_parse(input: &[u8]) -> IResult<&[u8], VB6ProjectClass, ProjectParseError> {
    // Class=CDecodeVarsClass; ..\DBCommon\CDecodeVarsClass.cls\r\n

    // a class line starts with 'Class=' which is then followed by
    // the class name then seperated by a semicolon and a single white space,
    // then the file name (usually a .cls file) and finally a newline.

    let remainder = input;

    let (remainder, _) = tag_no_case(b"Class=")(remainder)?;

    let (remainder, (name, path)) = name_path_tuple_parse(remainder)?;

    let class = VB6ProjectClass {
        name: String::from_utf8(name.to_vec()).unwrap(),
        path: String::from_utf8(path.to_vec()).unwrap(),
    };

    Ok((remainder, class))
}

fn reference_line_parse(input: &[u8]) -> IResult<&[u8], VB6ProjectReference, ProjectParseError> {
    // Reference=*\G{000440D8-E9ED-4435-A9A2-06B05387BB16}#c.0#0#..\DBCommon\Libs\VbIntellisenseFix.dll#VbIntellisenseFix\r\n

    // a reference line starts with 'Reference=' which is then followed by
    // "*\G" which indicates the next element will be
    // a GUID ie:
    // "{", followed 8 hexadecimal digits, a "-", four hexadecimal digits,
    // a "-", another four hexadecimal digits, a "-", another twelve
    // hexadecimal digits, and finally a "}#". The '# is used to indicate the
    // start of the next section.

    // We specify the impl here because it makes the following code
    // easier to read since we need to fully specifiy the parser function to
    // make error reporting easier.
    let tag_no_case = tag_no_case::<&[u8], &[u8], ProjectParseError>;
    let not_line_ending = not_line_ending::<&[u8], ProjectParseError>;
    let take_until = take_until::<&[u8], &[u8], ProjectParseError>;

    let remainder = input;

    let (remainder, (_, uuid_bytes)) =
        tuple((tag_no_case(b"Reference=*\\G{"), take_until(b"}#")))(remainder)?;

    let Ok(uuid_text) = str::from_utf8(uuid_bytes) else {
        return Err(nom::Err::Failure(ProjectParseError::UnableToParseUuid));
    };

    let Ok(uuid) = uuid::Uuid::parse_str(uuid_text) else {
        return Err(nom::Err::Failure(ProjectParseError::UnableToParseUuid));
    };

    let (remainder, _) = tag_no_case(b"}#")(remainder)?;

    // We again take until the first '#'. It's not clear what this value is.
    // I've seen values of 1.0, 2.0, c.0, and a few other 'something.something'
    // values.
    let (remainder, unknown1) = take_until(b"#")(remainder)?;

    let (remainder, _) = tag_no_case(b"#")(remainder)?;

    // We again take until the first '#'. It's not clear what this value is.
    // In every case, I've only seen '0'.
    let (remainder, unknown2) = take_until(b"#")(remainder)?;

    let (remainder, _) = tag_no_case(b"#")(remainder)?;

    // Another take until '#', this time we should have a path. This
    // path can be relative or absolute.
    let (remainder, path) = take_until(b"#")(remainder)?;

    let (remainder, _) = tag_no_case(b"#")(remainder)?;

    // Finally, we are grabbing the description, ie human readable, description
    // of this reference.
    let Ok((remainder, description)) = not_line_ending(remainder) else {
        return Err(nom::Err::Failure(ProjectParseError::NoLineEnding));
    };

    // We snagged up to the line ending before, now we want to actually get
    // that line ending as well.
    let (remainder, _) = line_ending(remainder)?;

    let reference = VB6ProjectReference {
        uuid,
        unknown1: String::from_utf8(unknown1.to_vec()).unwrap(),
        unknown2: String::from_utf8(unknown2.to_vec()).unwrap(),
        path: String::from_utf8(path.to_vec()).unwrap(),
        description: String::from_utf8(description.to_vec()).unwrap(),
    };

    Ok((remainder, reference))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn project_type_is_exe() {
        let project_type_line = b"Type=Exe\r\n";

        let (remainder, result) = project_type_parse(project_type_line).unwrap();

        assert_eq!(result, ProjectType::Exe);
        assert_eq!(remainder, b"");
    }

    #[test]
    fn project_type_is_oledll() {
        let project_type_line = b"Type=OleDll\r\n";

        let (remainder, result) = project_type_parse(project_type_line).unwrap();

        assert_eq!(result, ProjectType::OleDll);
        assert_eq!(remainder, b"");
    }

    #[test]
    fn project_type_is_unknown_type() {
        let project_type_line = b"Type=blah\r\n";

        let result = project_type_parse(project_type_line);

        assert!(result.is_err());
        assert_eq!(
            result.err().unwrap(),
            nom::Err::Failure(ProjectParseError::ProjectTypeUnknown)
        );
    }

    #[test]
    fn project_type_lacks_line_ending() {
        let project_type_line = b"Type=Exe";

        let result = project_type_parse(project_type_line);

        assert!(result.is_err());
        assert_eq!(
            result.err().unwrap(),
            nom::Err::Failure(ProjectParseError::NoLineEnding)
        );
    }

    #[test]
    fn reference_line_valid() {
        let reference_line = b"Reference=*\\G{000440D8-E9ED-4435-A9A2-06B05387BB16}#c.0#0#..\\DBCommon\\Libs\\VbIntellisenseFix.dll#VbIntellisenseFix\r\n";

        let (remainder, result) = reference_line_parse(reference_line).unwrap();

        let expected_uuid = Uuid::parse_str("000440D8-E9ED-4435-A9A2-06B05387BB16").unwrap();

        assert_eq!(remainder, []);
        assert_eq!(result.uuid, expected_uuid);
        assert_eq!(result.unknown1, "c.0");
        assert_eq!(result.unknown2, "0");
        assert_eq!(result.path, r"..\DBCommon\Libs\VbIntellisenseFix.dll");
        assert_eq!(result.description, r"VbIntellisenseFix");
    }

    #[test]
    fn object_line_valid() {
        let object_line = b"Object={C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0; crviewer.dll\r\n";

        let (remainder, result) = object_line_parse(object_line).unwrap();

        let expected_uuid = Uuid::parse_str("C4847593-972C-11D0-9567-00A0C9273C2A").unwrap();

        assert_eq!(remainder, []);
        assert_eq!(result.uuid, expected_uuid);
        assert_eq!(result.version, "8.0");
        assert_eq!(result.unknown1, "0");
        assert_eq!(result.file_name, "crviewer.dll");
    }

    #[test]
    fn module_line_valid() {
        let module_line = b"Module=modDBAssist; ..\\DBCommon\\DBAssist.bas\r\n";

        let (remainder, result) = module_line_parse(module_line).unwrap();

        assert_eq!(remainder, []);
        assert_eq!(result.name, "modDBAssist");
        assert_eq!(result.path, "..\\DBCommon\\DBAssist.bas");
    }

    #[test]
    fn class_line_valid() {
        let class_line = b"Class=CStatusBarClass; ..\\DBCommon\\CStatusBarClass.cls\r\n";

        let (remainder, result) = class_line_parse(class_line).unwrap();

        assert_eq!(remainder, []);
        assert_eq!(result.name, "CStatusBarClass");
        assert_eq!(result.path, "..\\DBCommon\\CStatusBarClass.cls");
    }

    #[test]
    fn key_value_line_valid() {
        let key_value_line = b"Designer=AllMfgStatus.Dsr\r\n";

        let (remainder, (key, value)) = key_value_pair_parse(key_value_line).unwrap();

        assert_eq!(remainder, []);
        assert_eq!(key, b"Designer");
        assert_eq!(value, "AllMfgStatus.Dsr");
    }

    #[test]
    fn key_qouted_value_line_valid() {
        let key_value_line = b"ResFile32=\"..\\DBCommon\\PSFC.RES\"\r\n";

        let (remainder, (key, value)) = key_qouted_value_pair_parse(key_value_line).unwrap();

        assert_eq!(remainder, []);
        assert_eq!(key, b"ResFile32");
        assert_eq!(value, "..\\DBCommon\\PSFC.RES");
    }
}
