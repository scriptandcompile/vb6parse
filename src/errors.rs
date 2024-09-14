use std::fmt::{Debug, Display, Formatter};

use winnow::{
    error::{ContextError, ErrorKind, ParseError, ParserError},
    stream::Stream,
};

use ariadne::{Label, Report, ReportKind, Source};

use thiserror::Error;

use crate::parsers::VB6Stream;

#[derive(Error, Debug)]
pub enum VB6ErrorKind {
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

    #[error("Unable to parse the property name")]
    PropertyNameUnparsable,

    #[error("Unable to find a semicolon ';' in this line.")]
    NoSemicolonSplit,

    #[error("Unable to find an equal '=' in this line.")]
    NoEqualSplit,

    #[error("Unable to parse the resource file name")]
    ResourceFileNameUnparsable,

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

    #[error("CompatibleMode Invalid. CompatibileMode can only be 0, 1, or 2.")]
    CompatibleModeUnparseable,

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

    #[error("CompatibilityMode can only be a 0 (false) or a -1 (true)")]
    CompatibilityModeUnparseable,

    #[error("NoControlUpgrade can only be a 0 (false) or a 1 (true)")]
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

    #[error("oundsCheck can only be a 0 (false) or a -1 (true)")]
    BoundsCheckUnparseable,

    #[error("OverflowCheck can only be a 0 (false) or a -1 (true)")]
    OverflowCheckUnparseable,

    #[error("FlPointCheck can only be a 0 (false) or a -1 (true)")]
    FlPointCheckUnparseable,

    #[error("FDIVCheck can only be a 0 (false) or a -1 (true)")]
    FDIVCheckUnparseable,

    #[error("UnroundedFP can only be a 0 (false) or a -1 (true)")]
    UnroundedFPUnparseable,

    #[error("StartMode can only be a 0 (false) or a -1 (true)")]
    StartModeUnparseable,

    #[error("Unattended can only be a 0 (false) or a -1 (true)")]
    UnattendedUnparseable,

    #[error("Retained can only be a 0 (false) or a -1 (true)")]
    RetainedUnparseable,

    #[error("DebugStartup can only be a 0 (false) or a -1 (true)")]
    DebugStartupOptionUnparseable,

    #[error("UseExistingBrowser can only be a 0 (false) or a -1 (true)")]
    UseExistingBrowserUnparseable,

    #[error("AutoRefresh can only be a 0 (false) or a -1 (true)")]
    AutoRefreshUnparseable,

    #[error("Thread Per Object is not a number.")]
    ThreadPerObjectUnparseable,

    #[error("Unknown attribute in class header file. Must be one of: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_Ext_KEY")]
    UnknownAttribute,

    #[error("Error parsing header")]
    Header,

    #[error("No class name in the class file")]
    MissingClassName,

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

    #[error("Unknown property in header file")]
    UnknownProperty,

    #[error("Invalid property value. Only 0 or -1 are valid for this property")]
    InvalidPropertyValueZeroNegOne,

    #[error("Invalid property value. Only True or False are valid for this property")]
    InvalidPropertyValueTrueFalse,

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

    #[error("Appearance can only be a 0 (Flat) or a 1 (ThreeD)")]
    AppearancePropertyInvalid,

    #[error("BorderStyle can only be a 0 (None) or 1 (FixedSingle)")]
    BorderStylePropertyInvalid,

    #[error("ClipControls can only be a 0 (false) or a 1 (true)")]
    ClipControlsPropertyInvalid,

    #[error("DragMode can only be 0 (Manual) or 1 (Automatic)")]
    DragModePropertyInvalid,

    #[error("Enabled can only be 0 (false) or a 1 (true)")]
    EnabledPropertyInvalid,

    #[error("MousePointer can only be 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom)")]
    MousePointerPropertyInvalid,

    #[error("OLEDropMode can only be 0 (None), or 1 (Manual)")]
    OLEDropModePropertyInvalid,

    #[error("RightToLeft can only be 0 (false) or a 1 (true)")]
    RightToLeftPropertyInvalid,

    #[error("Visible can only be 0 (false) or a 1 (true)")]
    VisiblePropertyInvalid,

    #[error("Variable names must be less than 255 characters in VB6.")]
    VariableNameTooLong,

    #[error("Internal Parser Error - please report this issue to the developers.")]
    WinnowParseError,
}

#[derive(Debug, Error)]
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
        let file_name = input.file_name.clone();
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
}

impl Display for VB6Error {
    fn fmt(&self, _: &mut Formatter) -> Result<(), std::fmt::Error> {
        let source = Source::from(self.source_code.clone());

        let kind_label = Label::new((
            self.file_name.clone(),
            self.source_offset..=self.source_offset,
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
    fn from_error_kind(input: &VB6Stream<'a>, _: ErrorKind) -> Self {
        VB6Error::new(input, VB6ErrorKind::WinnowParseError)
    }

    fn append(self, _: &VB6Stream, _: &<VB6Stream as Stream>::Checkpoint, _: ErrorKind) -> Self {
        self
    }
}

impl<'a> From<ParseError<VB6Stream<'a>, ContextError>> for VB6Error {
    fn from(err: ParseError<VB6Stream<'a>, ContextError>) -> Self {
        let input = err.input();
        VB6Error::new(input, VB6ErrorKind::WinnowParseError)
    }
}

impl<'a> ParserError<VB6Stream<'a>> for VB6ErrorKind {
    fn from_error_kind(_: &VB6Stream, _: ErrorKind) -> Self {
        VB6ErrorKind::WinnowParseError
    }

    fn append(self, _: &VB6Stream, _: &<VB6Stream as Stream>::Checkpoint, _: ErrorKind) -> Self {
        self
    }
}

impl<'a> From<ParseError<VB6Stream<'a>, ContextError>> for VB6ErrorKind {
    fn from(_: ParseError<VB6Stream<'a>, ContextError>) -> Self {
        VB6ErrorKind::WinnowParseError
    }
}
