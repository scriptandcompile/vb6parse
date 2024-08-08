#![warn(clippy::pedantic)]

use std::fmt::Debug;

use winnow::{
    error::{ContextError, ErrorKind, ParseError, ParserError},
    stream::Stream,
};

use miette::{Diagnostic, NamedSource, SourceOffset, SourceSpan};
use thiserror::Error;

use crate::vb6stream::VB6Stream;

#[derive(Debug, Error, Diagnostic)]
#[error("A parsing error occured")]
pub struct ErrorInfo {
    #[source_code]
    pub src: NamedSource<String>,
    #[label("oh no")]
    pub location: SourceSpan,
}

impl ErrorInfo {
    pub fn new(input: &VB6Stream, len: usize) -> Self {
        let code = input.stream.to_string();
        Self {
            src: NamedSource::new(input.file_name.clone(), code.clone()),
            location: SourceSpan::new(
                SourceOffset::from_location(code, input.line_number, input.column),
                len,
            ),
        }
    }
}

impl From<&ErrorInfo> for SourceSpan {
    fn from(info: &ErrorInfo) -> Self {
        info.location.clone()
    }
}

#[derive(Error, Debug, Diagnostic)]
pub enum VB6ParseError<I>
where
    I: Stream + Clone,
{
    #[error("The reference line has too many elements")]
    ReferenceExtraSections {
        #[label = "The reference line has too many elements"]
        info: ErrorInfo,
    },

    #[error("The reference line has too few elements")]
    ReferenceMissingSections {
        #[label = "The reference line has too few elements"]
        info: ErrorInfo,
    },

    #[error("The first line of a VB6 project file must be a project 'Type' entry.")]
    FirstLineNotProject {
        #[label = "The first line of a VB6 project file must be a project 'Type' entry."]
        info: ErrorInfo,
    },

    #[error("Line type is unknown.")]
    LineTypeUnknown {
        #[label = "Line type is unknown"]
        info: ErrorInfo,
    },

    #[error("Project type is not Exe, OleDll, Control, or OleExe")]
    ProjectTypeUnknown {
        #[label = "Project type is not Exe, OleDll, Control, or OleExe"]
        info: ErrorInfo,
    },

    #[error("Project line entry is not ended with a recognized line ending.")]
    NoLineEnding {
        #[label = "Project line entry is not ended with a recognized line ending."]
        info: ErrorInfo,
    },

    #[error("Unable to parse the Uuid")]
    UnableToParseUuid {
        #[label = "Unable to parse the Uuid"]
        info: ErrorInfo,
    },

    #[error("Unable to find a semicolon ';' in this line.")]
    NoSemicolonSplit {
        #[label = "Unable to find a semicolon ';' in this line."]
        info: ErrorInfo,
    },

    #[error("Unable to find an equal '=' in this line.")]
    NoEqualSplit {
        #[label = "Unable to find an equal '=' in this line."]
        info: ErrorInfo,
    },

    #[error("Unknown parser error")]
    Unparseable {
        #[label = "Unknown parser error"]
        info: ErrorInfo,
    },

    #[error("Major version is not a number.")]
    MajorVersionUnparseable {
        #[label = "Major version is not a number."]
        info: ErrorInfo,
    },

    #[error("Minor version is not a number.")]
    MinorVersionUnparseable {
        #[label = "Minor version is not a number."]
        info: ErrorInfo,
    },

    #[error("Revision version is not a number.")]
    RevisionVersionUnparseable {
        #[label = "Revision version is not a number."]
        info: ErrorInfo,
    },

    #[error("No property name found after BeginProperty keyword.")]
    NoPropertyName {
        #[label = "No property name found after BeginProperty keyword."]
        info: ErrorInfo,
    },

    #[error("AutoIncrement can only be a 0 (false) or a -1 (true)")]
    AutoIncrementUnparseable {
        #[label = "AutoIncrement can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("CompatibilityMode can only be a 0 (false) or a -1 (true)")]
    CompatibilityModeUnparseable {
        #[label = "CompatibilityMode can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("NoControlUpgrade can only be a 0 (false) or a -1 (true)")]
    NoControlUpgradeUnparsable {
        #[label = "NoControlUpgrade can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("ServerSupportFiles can only be a 0 (false) or a -1 (true)")]
    ServerSupportFilesUnparseable {
        #[label = "ServerSupportFiles can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("CompilationType can only be a 0 (false) or a -1 (true)")]
    CompilationTypeUnparseable {
        #[label = "CompilationType can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("OptimizationType can only be a 0 (false) or a -1 (true)")]
    OptimizationTypeUnparseable {
        #[label = "OptimizationType can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)")]
    FavorPentiumProUnparseable {
        #[label = "FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("Designer line is unparsable")]
    DesignerLineUnparseable {
        #[label = "Designer line is unparsable"]
        info: ErrorInfo,
    },

    #[error("Form line is unparsable")]
    FormLineUnparseable {
        #[label = "Form line is unparsable"]
        info: ErrorInfo,
    },

    #[error("UserControl line is unparsable")]
    UserControlLineUnparseable {
        #[label = "UserControl line is unparsable"]
        info: ErrorInfo,
    },

    #[error("UserDocument line is unparsable")]
    UserDocumentLineUnparseable {
        #[label = "UserDocument line is unparsable"]
        info: ErrorInfo,
    },

    #[error("Period expected in version number")]
    PeriodExpectedInVersionNumber {
        #[label = "Period expected in version number"]
        info: ErrorInfo,
    },

    #[error("CodeViewDebugInfo can only be a 0 (false) or a -1 (true)")]
    CodeViewDebugInfoUnparseable {
        #[label = "CodeViewDebugInfo can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("NoAliasing can only be a 0 (false) or a -1 (true)")]
    NoAliasingUnparseable {
        #[label = "NoAliasing can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("oundsCheck can only be a 0 (false) or a -1 (true)")]
    BoundsCheckUnparseable {
        #[label = "BoundsCheck can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("OverflowCheck can only be a 0 (false) or a -1 (true)")]
    OverflowCheckUnparseable {
        #[label = "OverflowCheck can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("FlPointCheck can only be a 0 (false) or a -1 (true)")]
    FlPointCheckUnparseable {
        #[label = "FlPointCheck can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("FDIVCheck can only be a 0 (false) or a -1 (true)")]
    FDIVCheckUnparseable {
        #[label = "FDIVCheck can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("UnroundedFP can only be a 0 (false) or a -1 (true)")]
    UnroundedFPUnparseable {
        #[label = "UnroundedFP can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("StartMode can only be a 0 (false) or a -1 (true)")]
    StartModeUnparseable {
        #[label = "StartMode can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("Unattended can only be a 0 (false) or a -1 (true)")]
    UnattendedUnparseable {
        #[label = "Unattended can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("Retained can only be a 0 (false) or a -1 (true)")]
    RetainedUnparseable {
        #[label = "Retained can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("DebugStartup can only be a 0 (false) or a -1 (true)")]
    DebugStartupOptionUnparseable {
        #[label = "DebugStartup can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("AutoRefresh can only be a 0 (false) or a -1 (true)")]
    AutoRefreshUnparseable {
        #[label = "AutoRefresh can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("Thread Per Object is not a number.")]
    ThreadPerObjectUnparseable {
        #[label = "Thread Per Object is not a number."]
        info: ErrorInfo,
    },

    #[error("Error parsing header")]
    #[diagnostic(transparent)]
    Header {
        #[label = "A parsing error occured"]
        info: ErrorInfo,
    },

    #[error("No class name in the class file")]
    #[diagnostic(transparent)]
    MissingClassName {
        #[label = "No class name found in the class file"]
        info: ErrorInfo,
    },

    #[error("Keyword not found")]
    #[diagnostic(transparent)]
    KeywordNotFound {
        #[label = "Keyword not found"]
        info: ErrorInfo,
    },

    #[error("Error parsing true/false from header. Must be a 0 (false) or a -1 (true)")]
    #[diagnostic(transparent)]
    TrueFalseZSeroNegOneUnparseable {
        #[label = "True/False can only be a 0 (false) or a -1 (true)"]
        info: ErrorInfo,
    },

    #[error("Error parsing the VB6 file contents")]
    #[diagnostic(transparent)]
    FileContent {
        #[label = "A parsing error occured"]
        info: ErrorInfo,
    },

    #[error("Max Threads is not a number.")]
    #[diagnostic(transparent)]
    MaxThreadsUnparseable {
        #[label = "Max Threads is not a number."]
        info: ErrorInfo,
    },

    #[error("No EndProperty found after BeginProperty")]
    #[diagnostic(transparent)]
    NoEndProperty {
        #[label = "No EndProperty found after BeginProperty"]
        info: ErrorInfo,
    },

    #[error("No line ending after EndProperty")]
    #[diagnostic(transparent)]
    NoLineEndingAfterEndProperty {
        #[label = "No line ending after EndProperty"]
        info: ErrorInfo,
    },

    #[error("Expected namespace after Begin keyword")]
    #[diagnostic(transparent)]
    NoNamespaceAfterBegin {
        #[label = "Expected namespace after Begin keyword"]
        info: ErrorInfo,
    },

    #[error("No dot found after namespace")]
    #[diagnostic(transparent)]
    NoDotAfterNamespace {
        #[label = "No dot found after namespace"]
        info: ErrorInfo,
    },

    #[error("No User Control name found after namespace and '.'")]
    #[diagnostic(transparent)]
    NoUserControlNameAfterDot {
        #[label = "No User Control name found after namespace and '.'"]
        info: ErrorInfo,
    },

    #[error("No space after Control kind")]
    #[diagnostic(transparent)]
    NoSpaceAfterControlKind {
        #[label = "No space after Control kind"]
        info: ErrorInfo,
    },

    #[error("No control name found after Control kind")]
    #[diagnostic(transparent)]
    NoControlNameAfterControlKind {
        #[label = "No control name found after Control kind"]
        info: ErrorInfo,
    },

    #[error("No line ending after Control name")]
    #[diagnostic(transparent)]
    NoLineEndingAfterControlName {
        #[label = "No line ending after Control name"]
        info: ErrorInfo,
    },

    #[error("Unknown token")]
    #[diagnostic(transparent)]
    UnknownToken {
        #[label = "Unknown token"]
        info: ErrorInfo,
    },

    #[error("Title text was unparsable")]
    #[diagnostic(transparent)]
    TitleUnparseable {
        #[label = "Title text was unparsable"]
        info: ErrorInfo,
    },

    #[error("Winnow Error")]
    #[diagnostic(code(vb6parse::error::project::parse_error))]
    ParseError(I, ErrorKind),
}

impl<I> ParserError<I> for VB6ParseError<I>
where
    I: Stream + Clone,
{
    fn from_error_kind(input: &I, kind: ErrorKind) -> Self {
        VB6ParseError::ParseError(input.clone(), kind)
    }

    fn append(self, _: &I, _: &<I as Stream>::Checkpoint, _: ErrorKind) -> Self {
        self
    }
}

impl<'a> From<VB6ParseError<VB6Stream<'a>>> for ErrorInfo {
    fn from(err: VB6ParseError<VB6Stream>) -> Self {
        match err {
            VB6ParseError::ParseError(input, _) => ErrorInfo::new(&input, 0),
            _ => ErrorInfo::new(&VB6Stream::new("".to_string(), &[]), 0),
        }
    }
}

impl<'a> From<winnow::error::ParseError<VB6Stream<'a>, ContextError>>
    for VB6ParseError<VB6Stream<'a>>
{
    fn from(err: ParseError<VB6Stream<'a>, ContextError>) -> Self {
        let input = err.input();
        VB6ParseError::ParseError(
            VB6Stream::new(input.file_name.clone(), input.stream),
            ErrorKind::Fail,
        )
    }
}
