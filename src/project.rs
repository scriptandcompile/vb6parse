#![warn(clippy::pedantic)]
use bstr::{BStr, ByteSlice};

use uuid::Uuid;

use winnow::{
    combinator::{alt, rest, separated_pair},
    error::{ErrMode, ParserError},
    token::{literal, take_until, take_while},
    PResult, Parser,
};

use crate::errors::VB6ProjectParseError;

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Project<'a> {
    pub project_type: CompileTargetType,
    pub references: Vec<VB6ProjectReference<'a>>,
    pub objects: Vec<VB6ProjectObject<'a>>,
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
    pub command_line_arguments: Option<&'a BStr>,
    pub name: Option<&'a BStr>,
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
pub struct VB6ProjectObject<'a> {
    pub uuid: Uuid,
    pub version: &'a BStr,
    pub unknown1: &'a BStr,
    pub file_name: &'a BStr,
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
    pub fn parse(input: &'a [u8]) -> Result<Self, VB6ProjectParseError<&'a [u8]>> {
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
        let mut command_line_arguments = Some(BStr::new(b""));
        let mut name = Some(BStr::new(b""));
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

        // First, bstr-ify the input and split it into lines.
        // The VB6 project format is basically a line-by-line format
        // so this is the easiest way to parse it.
        let input_bstr = input.as_bstr();
        let lines = input_bstr.lines();

        for mut line in lines {
            // Skip any empty lines.
            if line.is_empty() {
                continue;
            }

            // We also want to skip any '[MS Transaction Server]' header lines.
            // There should only be one in the file since it's only used once,
            // but we want to be flexible in what we accept so we skip any of
            // these kinds of header lines.
            if line.starts_with(b"[") {
                continue;
            }

            let (key, mut value) = match key_value_parse(&mut line, b"=") {
                Ok((key, value)) => (key, value),
                Err(e) => return Err(e.into_inner().unwrap()),
            };

            match key {
                b"Type" => {
                    project_type = match project_type_parse(&mut value) {
                        Ok(project) => Some(project),
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };
                }
                b"Designer" => {
                    designers.push(value.as_bstr());
                }
                b"Reference" => {
                    let reference = match reference_parse(&mut value) {
                        Ok(reference) => reference,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    references.push(reference);
                }
                b"Object" => {
                    let object = match object_parse(&mut value) {
                        Ok(object) => object,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    objects.push(object);
                }
                b"Module" => {
                    let module = match module_parse(&mut value) {
                        Ok(module) => module,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    modules.push(module);
                }
                b"Class" => {
                    let class = match class_parse(&mut value) {
                        Ok(class) => class,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    classes.push(class);
                }
                b"Form" => {
                    forms.push(value.into());
                }
                b"UserControl" => {
                    user_controls.push(value.into());
                }
                b"UserDocument" => {
                    user_documents.push(value.into());
                }
                b"ResFile32" => {
                    let res_file_32 = match qouted_value(&mut value, b"\"") {
                        Ok(res_file_32) => res_file_32,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    res_file_32_path = Some(res_file_32.as_bstr());
                }
                b"IconForm" => {
                    let form = match qouted_value(&mut value, b"\"") {
                        Ok(form) => form,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    icon_form = Some(form.as_bstr());
                }
                b"Startup" => {
                    let start_up_sub = match qouted_value(&mut value, b"\"") {
                        Ok(start_up_sub) => start_up_sub,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    startup = Some(start_up_sub.as_bstr());
                }
                b"HelpFile" => {
                    let help_file = match qouted_value(&mut value, b"\"") {
                        Ok(help_file) => help_file,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    help_file_path = Some(help_file.as_bstr());
                }
                b"Title" => {
                    let title_text = match qouted_value(&mut value, b"\"") {
                        Ok(title_text) => title_text,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    title = Some(title_text.as_bstr());
                }
                b"ExeName32" => {
                    let exe_file_name = match qouted_value(&mut value, b"\"") {
                        Ok(exe_file_name) => exe_file_name,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    exe_32_file_name = Some(exe_file_name.as_bstr());
                }
                b"Command32" => {
                    let command = match qouted_value(&mut value, b"\"") {
                        Ok(command) => command,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    command_line_arguments = Some(command.as_bstr());
                }
                b"Name" => {
                    let project_name = match qouted_value(&mut value, b"\"") {
                        Ok(project_name) => project_name,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    name = Some(project_name.as_bstr());
                }
                b"HelpContextID" => {
                    let help_context = match qouted_value(&mut value, b"\"") {
                        Ok(help_context) => help_context,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    help_context_id = Some(help_context.as_bstr());
                }
                b"CompatibleMode" => match qouted_true_false_parse(&mut value) {
                    Ok(val) => compatible_mode = val,
                    Err(_) => return Err(VB6ProjectParseError::CompatibilityModeUnparseable),
                },
                b"MajorVer" => {
                    let Ok(major_ver) = value.to_str() else {
                        return Err(VB6ProjectParseError::MajorVersionUnparseable);
                    };

                    major = major_ver.parse().unwrap();
                }
                b"MinorVer" => {
                    let Ok(minor_ver) = value.to_str() else {
                        return Err(VB6ProjectParseError::MinorVersionUnparseable);
                    };

                    minor = minor_ver.parse().unwrap();
                }
                b"RevisionVer" => {
                    let Ok(revision_ver) = value.to_str() else {
                        return Err(VB6ProjectParseError::MinorVersionUnparseable);
                    };

                    revision = revision_ver.parse().unwrap();
                }
                b"AutoIncrementVer" => {
                    let Ok(auto_increment) = value.to_str() else {
                        return Err(VB6ProjectParseError::AutoIncrementUnparseable);
                    };

                    auto_increment_revision = auto_increment.parse().unwrap();
                }
                b"NoControlUpgrade" => {
                    match true_false_parse(&mut value) {
                        // Invert answer since we inverted the name.
                        // This defaults to true, and is the most common value.
                        Ok(val) => upgrade_activex_controls = !val,
                        Err(_) => return Err(VB6ProjectParseError::NoControlUpgradeUnparsable),
                    }
                }
                b"ServerSupportFiles" => match true_false_parse(&mut value) {
                    Ok(val) => server_support_files = val,
                    Err(_) => return Err(VB6ProjectParseError::ServerSupportFilesUnparseable),
                },
                b"VersionCompanyName" => {
                    let company = match qouted_value(&mut value, b"\"") {
                        Ok(company) => company,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    company_name = Some(company.as_bstr());
                }
                b"VersionFileDescription" => {
                    let description = match qouted_value(&mut value, b"\"") {
                        Ok(description) => description,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    file_description = Some(description.as_bstr());
                }
                b"VersionLegalCopyright" => {
                    let legal_copyright = match qouted_value(&mut value, b"\"") {
                        Ok(legal_copyright) => legal_copyright,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    copyright = Some(legal_copyright.as_bstr());
                }
                b"VersionLegalTrademarks" => {
                    let legal_trademark = match qouted_value(&mut value, b"\"") {
                        Ok(legal_trademark) => legal_trademark,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    trademark = Some(legal_trademark.as_bstr());
                }
                b"VersionProductName" => {
                    let legal_product_name = match qouted_value(&mut value, b"\"") {
                        Ok(legal_product_name) => legal_product_name,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    product_name = Some(legal_product_name.as_bstr());
                }
                b"VersionComments" => {
                    let version_comments = match qouted_value(&mut value, b"\"") {
                        Ok(version_comments) => version_comments,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    comments = Some(version_comments.as_bstr());
                }
                b"CondComp" => {
                    let conditional = match qouted_value(&mut value, b"\"") {
                        Ok(conditional) => conditional,
                        Err(e) => return Err(e.into_inner().unwrap()),
                    };

                    conditional_compile = Some(conditional.as_bstr());
                }
                b"CompilationType" => match true_false_parse(&mut value) {
                    Ok(val) => compilation_type = val,
                    Err(_) => return Err(VB6ProjectParseError::CompilationTypeUnparseable),
                },
                b"OptimizationType" => match true_false_parse(&mut value) {
                    Ok(val) => optimization_type = val,
                    Err(_) => return Err(VB6ProjectParseError::OptimizationTypeUnparseable),
                },
                b"FavorPentiumPro(tm)" => match true_false_parse(&mut value) {
                    Ok(val) => favor_pentium_pro = val,
                    Err(_) => return Err(VB6ProjectParseError::FavorPentiumProUnparseable),
                },
                b"CodeViewDebugInfo" => match true_false_parse(&mut value) {
                    Ok(val) => code_view_debug_info = val,
                    Err(_) => return Err(VB6ProjectParseError::CodeViewDebugInfoUnparseable),
                },
                b"NoAliasing" => match true_false_parse(&mut value) {
                    // Invert the value since we inverted the name.
                    Ok(val) => aliasing = !val,
                    Err(_) => return Err(VB6ProjectParseError::NoAliasingUnparseable),
                },
                b"BoundsCheck" => match true_false_parse(&mut value) {
                    Ok(val) => bounds_check = val,
                    Err(_) => return Err(VB6ProjectParseError::BoundsCheckUnparseable),
                },
                b"OverflowCheck" => match true_false_parse(&mut value) {
                    Ok(val) => overflow_check = val,
                    Err(_) => return Err(VB6ProjectParseError::OverflowCheckUnparseable),
                },
                b"FlPointCheck" => match true_false_parse(&mut value) {
                    Ok(val) => floating_point_check = val,
                    Err(_) => return Err(VB6ProjectParseError::FlPointCheckUnparseable),
                },
                b"FDIVCheck" => match true_false_parse(&mut value) {
                    Ok(val) => pentium_fdiv_bug_check = val,
                    Err(_) => return Err(VB6ProjectParseError::FDIVCheckUnparseable),
                },
                b"UnroundedFP" => match true_false_parse(&mut value) {
                    Ok(val) => unrounded_floating_point = val,
                    Err(_) => return Err(VB6ProjectParseError::UnroundedFPUnparseable),
                },
                b"StartMode" => match true_false_parse(&mut value) {
                    Ok(val) => start_mode = val,
                    Err(_) => return Err(VB6ProjectParseError::StartModeUnparseable),
                },
                b"Unattended" => match true_false_parse(&mut value) {
                    Ok(val) => unattended = val,
                    Err(_) => return Err(VB6ProjectParseError::UnattendedUnparseable),
                },
                b"Retained" => match true_false_parse(&mut value) {
                    Ok(val) => retained = val,
                    Err(_) => return Err(VB6ProjectParseError::RetainedUnparseable),
                },
                b"ThreadPerObject" => {
                    let Ok(threads) = value.to_str() else {
                        return Err(VB6ProjectParseError::ThreadPerObjectUnparseable);
                    };

                    thread_per_object = threads.parse().unwrap();
                }
                b"MaxNumberOfThreads" => {
                    let Ok(max_threads) = value.to_str() else {
                        return Err(VB6ProjectParseError::MaxThreadsUnparseable);
                    };

                    max_number_of_threads = max_threads.parse().unwrap();
                }
                b"DebugStartupOption" => match true_false_parse(&mut value) {
                    Ok(val) => debug_startup_option = val,
                    Err(_) => return Err(VB6ProjectParseError::DebugStartupOptionUnparseable),
                },
                b"AutoRefresh" => match auto_refresh_parse(&mut value) {
                    Ok(val) => auto_refresh = val,
                    Err(_) => return Err(VB6ProjectParseError::AutoRefreshUnparseable),
                },
                _ => {
                    let line_type = key.as_bstr().to_string();
                    let val = value.as_bstr().to_string();
                    return Err(VB6ProjectParseError::LineTypeUnknown {
                        line_type,
                        value: val,
                    });
                }
            }
        }

        if project_type.is_none() {
            return Err(VB6ProjectParseError::FirstLineNotProject);
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
            command_line_arguments,
            name,
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

// fn file_name_without_extension<'a>(path: &'a BStr) -> Option<&'a BStr> {
//     let file_name = path.as_ref().file_name()?;

//     let file_limited_path = std::path::Path::new(file_name);

//     // Using 'with_extension' this way is gross, but 'file_prefix' is currently
//     // nightly only.
//     let file_name_without_extension = file_limited_path.with_extension("");

//     file_name_without_extension
//         .into_os_string()
//         .into_string()
//         .ok()
// }

type KVTuple<'a> = (&'a [u8], &'a [u8]);

fn key_value_parse<'a>(
    input: &mut &'a [u8],
    split: &[u8],
) -> PResult<KVTuple<'a>, VB6ProjectParseError<&'a [u8]>> {
    // Normally we would expect a key to be just an alphanumeric, but the
    // VB6 Project format includes this lovely gem of a line:
    //
    // ```FavorPentiumPro(tm)=0```
    //
    // So we have to include the '(' and the ')' characters in the key split.
    let take_key_valid_parse = take_while(1.., ('0'..='9', 'a'..='z', 'A'..='Z', '(', ')'));

    let (key, value) = separated_pair(take_key_valid_parse, split, rest).parse_next(input)?;

    Ok((key, value))
}

fn true_false_parse<'a>(input: &mut &'a [u8]) -> PResult<bool, VB6ProjectParseError<&'a [u8]>> {
    // 0 is false...and -1 is true.
    // Why vb6? What are you like this? Who hurt you?
    let result = alt(('0'.value(false), "-1".value(true))).parse_next(input)?;

    Ok(result)
}

fn auto_refresh_parse<'a>(input: &mut &'a [u8]) -> PResult<bool, VB6ProjectParseError<&'a [u8]>> {
    // 0 is false...and 1 is true.
    // Of course, VB6 being VB6, this is the only entry that does something different.
    // le sigh.
    let result = alt(('0'.value(false), "1".value(true))).parse_next(input)?;

    Ok(result)
}

fn qouted_true_false_parse<'a>(
    input: &mut &'a [u8],
) -> PResult<bool, VB6ProjectParseError<&'a [u8]>> {
    let mut qoute = qouted_value(input, b"\"")?;
    // 0 is false...and -1 is true.
    // Why vb6? What are you like this? Who hurt you?
    let result = alt(('0'.value(false), "-1".value(true))).parse_next(&mut qoute)?;

    Ok(result)
}

fn qouted_value<'a>(
    input: &mut &'a [u8],
    qoute_char: &[u8],
) -> PResult<&'a [u8], VB6ProjectParseError<&'a [u8]>> {
    literal(qoute_char).parse_next(input)?;
    let qouted_value = take_until(0.., qoute_char).parse_next(input)?;
    literal(qoute_char).parse_next(input)?;

    Ok(qouted_value)
}

fn object_parse<'a>(
    input: &mut &'a [u8],
) -> PResult<VB6ProjectObject<'a>, VB6ProjectParseError<&'a [u8]>> {
    literal('{').parse_next(input)?;

    let uuid_segment = take_until(1.., '}').parse_next(input)?;

    let Ok(uuid) = Uuid::parse_str(uuid_segment.to_str().unwrap()) else {
        return Err(ErrMode::Cut(VB6ProjectParseError::UnableToParseUuid));
    };

    literal("}#").parse_next(input)?;

    // still not sure what this element or the next represents.
    let version = take_until(1.., '#').parse_next(input)?;
    let version = version.as_bstr();

    literal(b"#").parse_next(input)?;

    let unknown1 = take_until(1.., ';').parse_next(input)?;
    let unknown1 = unknown1.as_bstr();

    // the file name is preceded by a semi-colon then a space. not sure why the
    // space is there, but it is. this strips it and the semi-colon out.
    literal(b"; ").parse_next(input)?;

    // the filename is the rest of the input.
    let file_name = input.as_bstr();

    let project_object = VB6ProjectObject {
        uuid,
        version,
        unknown1,
        file_name,
    };

    Ok(project_object)
}

fn module_parse<'a>(
    input: &mut &'a [u8],
) -> PResult<VB6ProjectModule<'a>, VB6ProjectParseError<&'a [u8]>> {
    let (name, path) = semicolon_space_split_parse.parse_next(input)?;

    let name = name.as_bstr();
    let path = path.as_bstr();

    let module = VB6ProjectModule { name, path };

    Ok(module)
}

fn class_parse<'a>(
    input: &mut &'a [u8],
) -> PResult<VB6ProjectClass<'a>, VB6ProjectParseError<&'a [u8]>> {
    let (name, path) = semicolon_space_split_parse.parse_next(input)?;

    let name = name.as_bstr();
    let path = path.as_bstr();

    let module = VB6ProjectClass { name, path };

    Ok(module)
}

fn semicolon_space_split_parse<'a>(
    input: &mut &'a [u8],
) -> PResult<(&'a [u8], &'a [u8]), VB6ProjectParseError<&'a [u8]>> {
    let left = take_until(1.., "; ").parse_next(input)?;
    literal("; ").parse_next(input)?;

    let right = input;

    Ok((left, right))
}

fn reference_parse<'a>(
    input: &mut &'a [u8],
) -> PResult<VB6ProjectReference<'a>, VB6ProjectParseError<&'a [u8]>> {
    literal(b"*\\G{").parse_next(input)?;

    // This is not the cleanest way to handle this but we need to replace the
    // first instance of "*\\G{" from the start of the segment. Notice the '\\'
    // escape sequence which is just a single slash in the file itself.
    // Then remove
    let uuid_segment = take_until(1.., "}#").parse_next(input)?;

    let Ok(uuid) = Uuid::parse_str(uuid_segment.to_str().unwrap()) else {
        return Err(ErrMode::Cut(VB6ProjectParseError::UnableToParseUuid));
    };

    literal("}#").parse_next(input)?;

    // still not sure what this element or the next represents.
    let Ok(unknown1): PResult<&[u8], ErrMode<VB6ProjectParseError<&'a [u8]>>> =
        take_until(1.., "#").parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ProjectParseError::ReferenceMissingSections));
    };

    literal("#").parse_next(input)?;
    let unknown1 = unknown1.as_bstr();

    let Ok(unknown2): PResult<&[u8], ErrMode<VB6ProjectParseError<&'a [u8]>>> =
        take_until(1.., "#").parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ProjectParseError::ReferenceMissingSections));
    };

    literal("#").parse_next(input)?;
    let unknown2 = unknown2.as_bstr();

    let Ok(path): PResult<&[u8], ErrMode<VB6ProjectParseError<&'a [u8]>>> =
        take_until(1.., "#").parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ProjectParseError::ReferenceMissingSections));
    };

    literal("#").parse_next(input)?;
    let path = path.as_bstr();

    let description = input;

    if description.contains(&b'#') {
        return Err(ErrMode::Cut(VB6ProjectParseError::ReferenceExtraSections));
    }

    let description = description.as_bstr();

    let reference = VB6ProjectReference {
        uuid,
        unknown1,
        unknown2,
        path,
        description,
    };

    Ok(reference)
}

fn project_type_parse<'a>(
    input: &mut &'a [u8],
) -> PResult<CompileTargetType, VB6ProjectParseError<&'a [u8]>> {
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
    let project_type = match alt((
        b"Exe".value(CompileTargetType::Exe),
        b"Control".value(CompileTargetType::Control),
        b"OleExe".value(CompileTargetType::OleExe),
        b"OleDll".value(CompileTargetType::OleDll),
    ))
    .parse_next(input)
    {
        Ok(type_project) => type_project,
        Err(e) => {
            let inner = e.or(ErrMode::Cut(VB6ProjectParseError::ProjectTypeUnknown));
            return Err(inner);
        }
    };

    Ok(project_type)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn project_type_is_exe() {
        let mut project_type_line = b"Type=Exe".as_slice();

        let (key, mut value) = key_value_parse(&mut project_type_line, b"=").unwrap();

        let result = project_type_parse(&mut value).unwrap();

        assert_eq!(key, b"Type");
        assert_eq!(result, CompileTargetType::Exe);
    }

    #[test]
    fn project_type_is_oledll() {
        let mut project_type_line = b"Type=OleDll".as_slice();

        let (key, mut value) = key_value_parse(&mut project_type_line, b"=").unwrap();

        let result = project_type_parse(&mut value).unwrap();

        assert_eq!(key, b"Type");
        assert_eq!(result, CompileTargetType::OleDll);
    }

    #[test]
    fn project_type_is_unknown_type() {
        let mut project_type_line = b"Type=blah".as_slice();

        let (key, mut value) = key_value_parse(&mut project_type_line, b"=").unwrap();
        let result = project_type_parse(&mut value);

        assert_eq!(key, b"Type");
        assert!(result.is_err());
        assert_eq!(
            result.unwrap_err(),
            ErrMode::Cut(VB6ProjectParseError::ProjectTypeUnknown)
        );
    }

    #[test]
    fn reference_line_valid() {
        let mut reference_line = b"Reference=*\\G{000440D8-E9ED-4435-A9A2-06B05387BB16}#c.0#0#..\\DBCommon\\Libs\\VbIntellisenseFix.dll#VbIntellisenseFix".as_slice();

        let (key, mut value) = key_value_parse(&mut reference_line, b"=").unwrap();
        let result = reference_parse(&mut value).unwrap();

        let expected_uuid = Uuid::parse_str("000440D8-E9ED-4435-A9A2-06B05387BB16").unwrap();

        assert_eq!(reference_line.len(), 0);
        assert_eq!(key, b"Reference");
        assert_eq!(result.uuid, expected_uuid);
        assert_eq!(result.unknown1, "c.0");
        assert_eq!(result.unknown2, "0");
        assert_eq!(result.path, r"..\DBCommon\Libs\VbIntellisenseFix.dll");
        assert_eq!(result.description, r"VbIntellisenseFix");
    }

    #[test]
    fn object_line_valid() {
        let mut object_line =
            b"Object={C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0; crviewer.dll".as_slice();

        let (key, mut value) = key_value_parse(&mut object_line, b"=").unwrap();
        let result = object_parse(&mut value).unwrap();

        let expected_uuid = Uuid::parse_str("C4847593-972C-11D0-9567-00A0C9273C2A").unwrap();

        assert_eq!(object_line.len(), 0);
        assert_eq!(key, b"Object");
        assert_eq!(result.uuid, expected_uuid);
        assert_eq!(result.version, "8.0");
        assert_eq!(result.unknown1, "0");
        assert_eq!(result.file_name, "crviewer.dll");
    }

    #[test]
    fn module_line_valid() {
        let mut module_line = b"Module=modDBAssist; ..\\DBCommon\\DBAssist.bas".as_slice();

        let (key, mut value) = key_value_parse(&mut module_line, b"=").unwrap();
        let result = module_parse(&mut value).unwrap();

        assert_eq!(module_line.len(), 0);
        assert_eq!(key, b"Module");
        assert_eq!(result.name, "modDBAssist");
        assert_eq!(result.path, "..\\DBCommon\\DBAssist.bas");
    }

    #[test]
    fn class_line_valid() {
        let mut class_line = b"Class=CStatusBarClass; ..\\DBCommon\\CStatusBarClass.cls".as_slice();

        let (key, mut value) = key_value_parse(&mut class_line, b"=").unwrap();
        let result = class_parse(&mut value).unwrap();

        assert_eq!(class_line.len(), 0);
        assert_eq!(key, b"Class");
        assert_eq!(result.name, "CStatusBarClass");
        assert_eq!(result.path, "..\\DBCommon\\CStatusBarClass.cls");
    }
}
