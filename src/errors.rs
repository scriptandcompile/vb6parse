#![warn(clippy::pedantic)]

use std::fmt::{Debug, Display, Formatter};

use winnow::{
    error::{ContextError, ErrorKind, ParseError, ParserError},
    stream::Stream,
};

use ariadne::{Label, Report, ReportKind, Source};

use thiserror::Error;

use crate::vb6stream::VB6Stream;

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

    #[error("Unable to find a semicolon ';' in this line.")]
    NoSemicolonSplit,

    #[error("Unable to find an equal '=' in this line.")]
    NoEqualSplit,

    #[error("Unknown parser error")]
    Unparseable,

    #[error("Major version is not a number.")]
    MajorVersionUnparseable,

    #[error("Minor version is not a number.")]
    MinorVersionUnparseable,

    #[error("Revision version is not a number.")]
    RevisionVersionUnparseable,

    #[error("No property name found after BeginProperty keyword.")]
    NoPropertyName,

    #[error("AutoIncrement can only be a 0 (false) or a -1 (true)")]
    AutoIncrementUnparseable,

    #[error("CompatibilityMode can only be a 0 (false) or a -1 (true)")]
    CompatibilityModeUnparseable,

    #[error("NoControlUpgrade can only be a 0 (false) or a -1 (true)")]
    NoControlUpgradeUnparsable,

    #[error("ServerSupportFiles can only be a 0 (false) or a -1 (true)")]
    ServerSupportFilesUnparseable,

    #[error("CompilationType can only be a 0 (false) or a -1 (true)")]
    CompilationTypeUnparseable,

    #[error("OptimizationType can only be a 0 (false) or a -1 (true)")]
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

    #[error("AutoRefresh can only be a 0 (false) or a -1 (true)")]
    AutoRefreshUnparseable,

    #[error("Thread Per Object is not a number.")]
    ThreadPerObjectUnparseable,

    #[error("Error parsing header")]
    Header,

    #[error("No class name in the class file")]
    MissingClassName,

    #[error("Keyword not found")]
    KeywordNotFound,

    #[error("Error parsing true/false from header. Must be a 0 (false) or a -1 (true)")]
    TrueFalseZSeroNegOneUnparseable,

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

    #[error("Winnow Error")]
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
        Report::build(ReportKind::Error, (), 34)
            .with_message(self.kind.to_string())
            .with_label(
                Label::new(self.source_offset..self.source_offset + 1)
                    .with_message(self.kind.to_string()),
            )
            .finish()
            .print(Source::from(self.source_code.as_str()))
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
        VB6Error::new(&input, VB6ErrorKind::WinnowParseError)
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
