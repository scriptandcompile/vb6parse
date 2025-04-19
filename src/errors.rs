use std::fmt::{Debug, Display, Formatter};
use std::path::Path;

use winnow::{
    error::{ContextError, ParseError, ParserError},
    stream::Stream,
};

use ariadne::{Label, Report, ReportKind, Source};

use crate::parsers::VB6Stream;

#[derive(thiserror::Error, PartialEq, Eq, Debug)]
pub enum PropertyError {
    #[error("Appearance can only be a 0 (Flat) or a 1 (ThreeD)")]
    AppearanceInvalid,

    #[error("BorderStyle can only be a 0 (None) or 1 (FixedSingle)")]
    BorderStyleInvalid,

    #[error("ClipControls can only be a 0 (false) or a 1 (true)")]
    ClipControlsInvalid,

    #[error("DragMode can only be 0 (Manual) or 1 (Automatic)")]
    DragModeInvalid,

    #[error("Enabled can only be 0 (false) or a 1 (true)")]
    EnabledInvalid,

    #[error("MousePointer can only be 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom)")]
    MousePointerInvalid,

    #[error("OLEDropMode can only be 0 (None), or 1 (Manual)")]
    OLEDropModeInvalid,

    #[error("RightToLeft can only be 0 (false) or a 1 (true)")]
    RightToLeftInvalid,

    #[error("Visible can only be 0 (false) or a 1 (true)")]
    VisibleInvalid,

    #[error("Unknown property in header file")]
    UnknownProperty,

    #[error("Invalid property value. Only 0 or -1 are valid for this property")]
    InvalidPropertyValueZeroNegOne,

    #[error("Unable to parse the property name")]
    NameUnparsable,

    #[error("Unable to parse the resource file name")]
    ResourceFileNameUnparsable,

    #[error("Unable to parse the offset into the resource file for property")]
    OffsetUnparsable,

    #[error("Invalid property value. Only True or False are valid for this property")]
    InvalidPropertyValueTrueFalse,
}

#[derive(thiserror::Error, Debug)]
pub enum VB6ErrorKind {
    #[error("Property parsing error")]
    Property(#[from] PropertyError),

    #[error("Resource file parsing error")]
    ResourceFile(#[from] std::io::Error),

    #[error("Error reading the source file")]
    SourceFileError(std::io::Error),

    #[error("The file contains more than a significant number of non-ASCII characters. This file was likely saved in a non-English character set. The vb6parse crate currently does not support non-english vb6 files.")]
    LikelyNonEnglishCharacterSet,

    #[error("The reference line has too many elements")]
    ReferenceExtraSections,

    #[error("The reference line has too few elements")]
    ReferenceMissingSections,

    #[error("The first line of a VB6 project file must be a project 'Type' entry.")]
    FirstLineNotProject,

    #[error("Line type is unknown.")]
    LineTypeUnknown,

    #[error("Project type is not Exe, OleDll, Control, or OleExe")]
    ProjectTypeUnknown,

    #[error("Project line entry is not ended with a recognized line ending.")]
    NoLineEnding,

    #[error("Unable to parse the Uuid")]
    UnableToParseUuid,

    #[error("Unable to find a semicolon ';' in this line.")]
    NoSemicolonSplit,

    #[error("Unable to find an equal '=' in this line.")]
    NoEqualSplit,

    #[error("While trying to parse the offset into the resource file, no colon ':' was found.")]
    NoColonForOffsetSplit,

    #[error("No key value divider found in the line.")]
    NoKeyValueDividerFound,

    #[error("Unknown parser error")]
    Unparseable,

    #[error("Major version is not a number.")]
    MajorVersionUnparseable,

    #[error("Unable to parse hex address from DllBaseAddress key")]
    DllBaseAddressUnparseable,

    #[error("The Startup object is not a valid parameter. Must be a qouted startup method/object, \"(None)\", !(None)!, \"\", or \"!!\"")]
    StartupUnparseable,

    #[error("The Name parameter is invalid. Must be a qouted name, \"(None)\", !(None)!, \"\", or \"!!\"")]
    NameUnparseable,

    #[error("The CommandLine parameter is invalid. Must be a qouted command line, \"(None)\", !(None)!, \"\", or \"!!\"")]
    CommandLineUnparseable,

    #[error("The HelpContextId parameter is not a valid parameter line. Must be a qouted help context id, \"(None)\", !(None)!, \"\", or \"!!\"")]
    HelpContextIdUnparseable,

    #[error("Minor version is not a number.")]
    MinorVersionUnparseable,

    #[error("Revision version is not a number.")]
    RevisionVersionUnparseable,

    #[error("Unable to parse the value after ThreadingModel key")]
    ThreadingModelUnparseable,

    #[error("ThreadingModel can only be 0 (Apartment Threaded), or 1 (Single Threaded)")]
    ThreadingModelInvalid,

    #[error("No property name found after BeginProperty keyword.")]
    NoPropertyName,

    #[error("Unable to parse the RelatedDoc property line.")]
    RelatedDocLineUnparseable,

    #[error("AutoIncrement can only be a 0 (false) or a -1 (true)")]
    AutoIncrementUnparseable,

    #[error("CompatibilityMode can only be a 0 (CompatibilityMode::NoCompatibility), 1 (CompatibilityMode::Project), or 2 (CompatibilityMode::CompatibleExe)")]
    CompatibilityModeUnparseable,

    #[error("NoControlUpgrade can only be a 0 (UpgradeControls::Upgrade) or a 1 (UpgradeControls::NoUpgrade)")]
    NoControlUpgradeUnparsable,

    #[error("ServerSupportFiles can only be a 0 (false) or a -1 (true)")]
    ServerSupportFilesUnparseable,

    #[error("Comment line was unparsable")]
    CommentUnparseable,

    #[error("PropertyPage line was unparsable")]
    PropertyPageUnparseable,

    #[error("CompilationType can only be a 0 (false) or a -1 (true)")]
    CompilationTypeUnparseable,

    #[error("OptimizationType can only be a 0 (FastCode) or 1 (SmallCode), or 2 (NoOptimization)")]
    OptimizationTypeUnparseable,

    #[error("FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)")]
    FavorPentiumProUnparseable,

    #[error("Designer line is unparsable")]
    DesignerLineUnparseable,

    #[error("Form line is unparsable")]
    FormLineUnparseable,

    #[error("UserControl line is unparsable")]
    UserControlLineUnparseable,

    #[error("UserDocument line is unparsable")]
    UserDocumentLineUnparseable,

    #[error("Period expected in version number")]
    PeriodExpectedInVersionNumber,

    #[error("CodeViewDebugInfo can only be a 0 (false) or a -1 (true)")]
    CodeViewDebugInfoUnparseable,

    #[error("NoAliasing can only be a 0 (false) or a -1 (true)")]
    NoAliasingUnparseable,

    #[error("RemoveUnusedControlInfo can only be 0 (UnusedControlInfo::Retain) or -1 (UnusedControlInfo::Remove)")]
    UnusedControlInfoUnparseable,

    #[error("oundsCheck can only be a 0 (false) or a -1 (true)")]
    BoundsCheckUnparseable,

    #[error("OverflowCheck can only be a 0 (false) or a -1 (true)")]
    OverflowCheckUnparseable,

    #[error("FlPointCheck can only be a 0 (false) or a -1 (true)")]
    FlPointCheckUnparseable,

    #[error("FDIVCheck can only be a 0 (PentiumFDivBugCheck::CheckPentiumFDivBug) or a -1 (PentiumFDivBugCheck::NoPentiumFDivBugCheck)")]
    FDIVCheckUnparseable,

    #[error("UnroundedFP can only be a 0 (UnroundedFloatingPoint::DoNotAllow) or a -1 (UnroundedFloatingPoint::Allow)")]
    UnroundedFPUnparseable,

    #[error("StartMode can only be a 0 (StartMode::StandAlone) or a 1 (StartMode::Automation)")]
    StartModeUnparseable,

    #[error("Unattended can only be a 0 (Unattended::False) or a -1 (Unattended::True)")]
    UnattendedUnparseable,

    #[error(
        "Retained can only be a 0 (Retained::UnloadOnExit) or a 1 (Retained::RetainedInMemory)"
    )]
    RetainedUnparseable,

    #[error("Unable to parse the ShurtCut property.")]
    ShortCutUnparseable,

    #[error("DebugStartup can only be a 0 (false) or a -1 (true)")]
    DebugStartupOptionUnparseable,

    #[error("UseExistingBrowser can only be a 0 (UseExistingBrowser::DoNotUse) or a -1 (UseExistingBrowser::Use)")]
    UseExistingBrowserUnparseable,

    #[error("AutoRefresh can only be a 0 (false) or a -1 (true)")]
    AutoRefreshUnparseable,

    #[error("Data control Connection type is not valid.")]
    ConnectionTypeUnparseable,

    #[error("Thread Per Object is not a number.")]
    ThreadPerObjectUnparseable,

    #[error("Unknown attribute in class header file. Must be one of: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_Ext_KEY")]
    UnknownAttribute,

    #[error("Error parsing header")]
    Header,

    #[error("No name in the attribute section of the VB6 file")]
    MissingNameAttribute,

    #[error("Keyword not found")]
    KeywordNotFound,

    #[error("Error parsing true/false from header. Must be a 0 (false), -1 (true), or 1 (true)")]
    TrueFalseOneZeroNegOneUnparseable,

    #[error("Error parsing the VB6 file contents")]
    FileContent,

    #[error("Max Threads is not a number.")]
    MaxThreadsUnparseable,

    #[error("No EndProperty found after BeginProperty")]
    NoEndProperty,

    #[error("No line ending after EndProperty")]
    NoLineEndingAfterEndProperty,

    #[error("Expected namespace after Begin keyword")]
    NoNamespaceAfterBegin,

    #[error("No dot found after namespace")]
    NoDotAfterNamespace,

    #[error("No User Control name found after namespace and '.'")]
    NoUserControlNameAfterDot,

    #[error("No space after Control kind")]
    NoSpaceAfterControlKind,

    #[error("No control name found after Control kind")]
    NoControlNameAfterControlKind,

    #[error("No line ending after Control name")]
    NoLineEndingAfterControlName,

    #[error("Unknown token")]
    UnknownToken,

    #[error("Title text was unparsable")]
    TitleUnparseable,

    #[error("Unable to parse hex color value")]
    HexColorParseError,

    #[error("Unknown control in control list")]
    UnknownControlKind,

    #[error("Property name is not a valid ASCII string")]
    PropertyNameAsciiConversionError,

    #[error("String is unterminated")]
    UnterminatedString,

    #[error("Unable to parse VB6 string.")]
    StringParseError,

    #[error("Property value is not a valid ASCII string")]
    PropertyValueAsciiConversionError,

    #[error("Key value pair format is incorrect")]
    KeyValueParseError,

    #[error("Namespace is not a valid ASCII string")]
    NamespaceAsciiConversionError,

    #[error("Control kind is not a valid ASCII string")]
    ControlKindAsciiConversionError,

    #[error("Qualified control name is not a valid ASCII string")]
    QualifiedControlNameAsciiConversionError,

    #[error("Variable names must be less than 255 characters in VB6.")]
    VariableNameTooLong,

    #[error("Internal Parser Error - please report this issue to the developers.")]
    InternalParseError,
}

#[derive(Debug, thiserror::Error)]
pub struct VB6Error {
    pub file_name: String,

    pub source_code: String,

    pub source_offset: usize,

    pub column: usize,

    pub line_number: usize,

    pub kind: VB6ErrorKind,
}

impl VB6Error {
    #[must_use]
    pub fn new(input: &VB6Stream, kind: VB6ErrorKind) -> Self {
        // Get the file name from the file path in the input stream.

        // If the file path is empty, use "unknown" as a placeholder.
        // This is useful for errors that occur in the input stream
        // that do not have a file path associated with them.
        let file_name = Path::new(&input.file_path)
            .file_name()
            .map(|name| name.to_string_lossy().to_string())
            .unwrap_or_else(|| "unknown".to_string());

        let source_code = input.stream.to_string();
        let source_offset = input.index;
        let column = input.column;
        let line_number = input.line_number;

        Self {
            file_name,
            source_code,
            source_offset,
            column,
            line_number,
            kind,
        }
    }

    #[must_use]
    pub fn new_without_stream(kind: VB6ErrorKind) -> Self {
        VB6Error {
            file_name: "unknown".to_string(),
            source_code: String::new(),
            source_offset: 0,
            column: 0,
            line_number: 0,
            kind,
        }
    }
}

impl Display for VB6Error {
    fn fmt(&self, _: &mut Formatter) -> Result<(), std::fmt::Error> {
        let source = Source::from(self.source_code.clone());

        let kind_label = Label::new((
            self.file_name.clone(),
            // TODO: Use RangeInclusive variant when new
            // version of ariadne is released.
            self.source_offset..self.source_offset + 1,
        ))
        .with_message(self.kind.to_string());

        Report::build(
            ReportKind::Error,
            self.file_name.clone(),
            self.source_offset,
        )
        .with_message("Parsing error")
        .with_label(kind_label)
        .finish()
        .eprint((self.file_name.clone(), source))
        .unwrap();

        Ok(())
    }
}

impl<'a> ParserError<VB6Stream<'a>> for VB6Error {
    type Inner = VB6ErrorKind;

    fn from_input(input: &VB6Stream<'a>) -> Self {
        VB6Error::new(input, VB6ErrorKind::InternalParseError)
    }

    fn into_inner(self) -> winnow::Result<Self::Inner, Self> {
        Ok(self.kind)
    }

    fn append(self, _: &VB6Stream, _: &<VB6Stream as Stream>::Checkpoint) -> Self {
        self
    }
}

impl<'a> From<ParseError<VB6Stream<'a>, ContextError>> for VB6Error {
    fn from(err: ParseError<VB6Stream<'a>, ContextError>) -> Self {
        let input = err.input();
        VB6Error::new(input, VB6ErrorKind::InternalParseError)
    }
}

impl<'a> ParserError<VB6Stream<'a>> for VB6ErrorKind {
    type Inner = VB6ErrorKind;

    fn into_inner(self) -> winnow::Result<Self::Inner, Self> {
        Ok(self)
    }

    fn from_input(_: &VB6Stream) -> Self {
        VB6ErrorKind::InternalParseError
    }

    fn append(self, _: &VB6Stream, _: &<VB6Stream as Stream>::Checkpoint) -> Self {
        self
    }
}

impl<'a> From<ParseError<VB6Stream<'a>, ContextError>> for VB6ErrorKind {
    fn from(_: ParseError<VB6Stream<'a>, ContextError>) -> Self {
        VB6ErrorKind::InternalParseError
    }
}
