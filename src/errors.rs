#![warn(clippy::pedantic)]
use winnow::{
    error::{ErrorKind, ParserError},
    stream::Stream,
};

#[derive(thiserror::Error, Debug, PartialEq)]
pub enum VB6ProjectParseError<I> {
    #[error("The reference line has too many elements")]
    ReferenceExtraSections,
    #[error("The reference line has too few elements")]
    ReferenceMissingSections,
    #[error("The first line of a VB6 project file must be a project 'Type' entry.")]
    FirstLineNotProject,
    #[error("Line type is unknown.\r\n  Line Type: '{line_type}'\r\n  Value: '{value}'")]
    LineTypeUnknown { line_type: String, value: String },
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
    #[error("Max Threads is not a number.")]
    MaxThreadsUnparseable,

    #[error("Winnow Error")]
    ParseError(I, ErrorKind),
}

impl<I: Stream + Clone> ParserError<I> for VB6ProjectParseError<I> {
    fn from_error_kind(input: &I, kind: ErrorKind) -> Self {
        VB6ProjectParseError::ParseError(input.clone(), kind)
    }

    fn append(self, _: &I, _: &<I as Stream>::Checkpoint, _: ErrorKind) -> Self {
        self
    }
}
