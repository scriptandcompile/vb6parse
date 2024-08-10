#![warn(clippy::pedantic)]

use std::fmt::Debug;

use winnow::{
    error::{ContextError, ErrorKind, ParseError, ParserError},
    stream::Stream,
};

use miette::{Diagnostic, NamedSource, SourceOffset, SourceSpan};
use thiserror::Error;

use crate::vb6stream::VB6Stream;

#[derive(Error, Debug, Diagnostic)]
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

    #[error("Winnow Error")]
    WinnowParseError,
}

#[derive(Debug, Error, Diagnostic)]
#[error("{kind}")]
#[diagnostic()]
pub struct VB6Error {
    #[source_code]
    pub src: NamedSource<String>,

    #[label]
    pub location: SourceSpan,

    #[diagnostic(transparent)]
    pub kind: VB6ErrorKind,
}

impl VB6Error {
    pub fn new(input: &VB6Stream, kind: VB6ErrorKind) -> Self {
        let code = input.stream.to_string();
        let src =
            NamedSource::new(input.file_name.clone(), code.clone()).with_language("VisualBasic 6");
        let len = code.len();
        Self {
            src,
            location: SourceSpan::new(
                SourceOffset::from_location(code, input.line_number, input.column),
                len,
            ),
            kind,
        }
    }
}

impl From<&VB6Error> for SourceSpan {
    fn from(info: &VB6Error) -> Self {
        info.location.clone()
    }
}

impl<'a> ParserError<VB6Stream<'a>> for VB6Error {
    fn from_error_kind(input: &VB6Stream, _: ErrorKind) -> Self {
        VB6Error::new(input, VB6ErrorKind::WinnowParseError)
    }

    fn append(self, _: &VB6Stream, _: &<VB6Stream as Stream>::Checkpoint, _: ErrorKind) -> Self {
        self
    }
}

impl<'a> From<winnow::error::ParseError<VB6Stream<'a>, ContextError>> for VB6Error {
    fn from(err: ParseError<VB6Stream<'a>, ContextError>) -> Self {
        let input = err.input();
        VB6Error::new(&input, VB6ErrorKind::WinnowParseError)
    }
}
