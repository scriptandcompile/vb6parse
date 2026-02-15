//! Form file (.frm) parsing and validation errors.

use thiserror::Error;

/// Errors that can occur during form file parsing and validation.
#[derive(Error, Debug, Clone, PartialEq, Eq)]
pub enum FormError {
    /// The `ComboBox` style is invalid.
    #[error("The `ComboBox` style is invalid: '{value}'. Only 0 (Dropdown Combo), 1 (Simple Combo), or 2 (Dropdown List) are valid styles.")]
    InvalidComboBoxStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `CheckBox` value is invalid.
    #[error("The `CheckBox` value is invalid: '{value}'. Only 0 (Unchecked), 1 (Checked), or 2 (Grayed) are valid values.")]
    InvalidCheckBoxValue {
        /// The invalid value that was found.
        value: String,
    },

    /// The `BOFAction` property has an invalid value.
    #[error("The `BOFAction` value is invalid: '{value}'. Only 0 (MoveFirst), or 1 (BOF) are valid values.")]
    InvalidBOFAction {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ConnectionType` property has an invalid value.
    #[error("The `ConnectionType` value is invalid: '{value}'. Only 'Access', 'dBase III', 'dBase IV', 'dBase 5.0', 'Excel 3.0', 'Excel 4.0', 'Excel 5.0', 'Excel 8.0', 'FoxPro 2.0', 'FoxPro 2.5', 'FoxPro 2.6', 'FoxPro 3.0', 'Lotus WK1', 'Lotus WK3', 'Lotus WK4', 'Paradox 3.X', 'Paradox 4.X', 'Paradox 5.X', or 'Text' are valid values.")]
    InvalidConnectionType {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DefaultCursorType` property has an invalid value.
    #[error("The `DefaultCursorType` value is invalid: '{value}'. Only 0 (Default), 1 (Odbc), or 2 (ServerSide) are valid values.")]
    InvalidDefaultCursorType {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DatabaseDriverType` property has an invalid value.
    #[error("The `DatabaseDriverType` value is invalid: '{value}'. Only 1 (ODBC), or 2 (Jet) are valid values.")]
    InvalidDatabaseDriverType {
        /// The invalid value that was found.
        value: String,
    },

    /// The `EOFAction` property has an invalid value.
    #[error("The `EOFAction` value is invalid: '{value}'. Only 0 (MoveLast), 1 (EOF), or 2 (AddNew) are valid values.")]
    InvalidEOFAction {
        /// The invalid value that was found.
        value: String,
    },

    /// The record set type is invalid.
    #[error("The `RecordSetType` value is invalid: '{value}'. Only 0 (Table), 1 (Dynaset), or 2 (Snapshot) are valid values.")]
    InvalidRecordSetType {
        /// The invalid value that was found.
        value: String,
    },

    /// The archive attribute is invalid.
    #[error("The `ArchiveAttribute` value is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    InvalidArchiveAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The hidden attribute is invalid.
    #[error("The `Hidden` valud is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    InvalidHiddenAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ReadOnly` attribute is invalid.
    #[error("The `ReadOnly` value is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    InvalidReadOnlyAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The `System` attribute is invalid.
    #[error("The `System` value is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    InvalidSystemAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Normal` attribute is invalid.
    #[error("The `Normal` value is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    InvalidNormalAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The `FormBorderStyle` property has an invalid value.
    #[error("The `FormBorderStyle` value is invalid: '{value}'. Only 0 (None), 1 (FixedSingle), 2 (Sizable), 3 (FixedDialog), 4 (FixedToolWindow), or 5 (SizableToolWindow) are valid values.")]
    InvalidFormBorderStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ControlBox` property has an invalid value.
    #[error("The `ControlBox` value is invalid: '{value}'. Only 0 (Excluded) or -1 (Included) are valid values.")]
    InvalidControlBox {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MaxButton` property has an invalid value.
    #[error("The `MaxButton` value is invalid: '{value}'. Only 0 (Excluded) or -1 (Included) are valid values.")]
    InvalidMaxButton {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MinButton` property has an invalid value.
    #[error("The `MinButton` value is invalid: '{value}'. Only 0 (Excluded) or -1 (Included) are valid values.")]
    InvalidMinButton {
        /// The invalid value that was found.
        value: String,
    },

    /// The `PaletteMode` property has an invalid value.
    #[error("The `PaletteMode` value is invalid: '{value}'. Only 0 (HalfTone), 1 (UseZOrder), or 2 (Custom) are valid values.")]
    InvalidPaletteMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `WordWrap` property has an invalid value.
    #[error(
        "The `WordWrap` value is invalid: '{value}'. Only 0 (False) or -1 (True) are valid values."
    )]
    InvalidWordWrap {
        /// The invalid value that was found.
        value: String,
    },

    /// The `WhatsThisButton` value is invalid.
    #[error("The `WhatsThisButton` value is invalid: '{value}'. Only 0 (Excluded) or -1 (Included) are valid values.")]
    InvalidWhatsThisButton {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ShowInTaskbar` value is invalid.
    #[error("The `ShowInTaskbar` value is invalid: '{value}'. Only 0 (Hide) or -1 (Show) are valid values.")]
    InvalidShowInTaskbar {
        /// The invalid value that was found.
        value: String,
    },

    /// The `NegotiatePosition` value is invalid.
    #[error("The `NegotiatePosition` value is invalid: '{value}'. Only 0 (None), 1 (Left), 2 (Middle), or 3 (Right) are valid values.")]
    InvalidNegotiatePosition {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ListBoxStyle` value is invalid.
    #[error("The `ListBoxStyle` value is invalid: '{value}'. Only 0 (Standard) or 1 (Checkbox) are valid values.")]
    InvalidListBoxStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `AutoSize` value is invalid.
    #[error("The `AutoSize` value is invalid: '{value}'. Only 0 (Fixed) or -1 (Resize) are valid values.")]
    InvalidAutoSize {
        /// The invalid value that was found.
        value: String,
    },

    /// The `AutoRedraw` value is invalid.
    #[error("The `AutoRedraw` value is invalid: '{value}'. Only 0 (Manual) or -1 (Automatic) are valid values.")]
    InvalidAutoRedraw {
        /// The invalid value that was found.
        value: String,
    },

    /// The `TextDirection` value is invalid.
    #[error("The `TextAlign` value is invalid: '{value}'. Only 0 (LeftToRight) or -1 (RightToLeft) are valid values.")]
    InvalidTextDirection {
        /// The invalid value that was found.
        value: String,
    },

    /// The `TabStop` value is invalid.
    #[error("The `TabStop` value is invalid: '{value}'. Only 0 (ProgrammaticOnly) or -1 (Included) are valid values.")]
    InvalidTabStop {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Visibility` value is invalid.
    #[error("The `Visibility` value is invalid: '{value}'. Only 0 (Hidden) or -1 (Visible) are valid values.")]
    InvalidVisibility {
        /// The invalid value that was found.
        value: String,
    },

    /// The `HasDeviceContext` value is invalid.
    #[error("The `HasDeviceContext` value is invalid: '{value}'. Only 0 (No) or -1 (Yes) are valid values.")]
    InvalidHasDeviceContext {
        /// The invalid value that was found.
        value: String,
    },

    /// The `CausesValidation` value is invalid.
    #[error("The `CausesValidation` value is invalid: '{value}'. Only 0 (No) or -1 (Yes) are valid values.")]
    InvalidCausesValidation {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Movability` value is invalid.
    #[error("The `Movability` value is invalid: '{value}'. Only 0 (Fixed) or -1 (Movable) are valid values.")]
    InvalidMovability {
        /// The invalid value that was found.
        value: String,
    },

    /// The `FontTransparency` value is invalid.
    #[error("The `FontTransparency` value is invalid: '{value}'. Only 0 (Opaque) or -1 (Transparent) are valid values.")]
    InvalidFontTransparency {
        /// The invalid value that was found.
        value: String,
    },

    /// The `WhatsThisHelp` value is invalid.
    #[error("The `WhatsThisHelp` value is invalid: '{value}'. Only 0 (F1Help) or -1 (WhatsThisHelp) are valid values.")]
    InvalidWhatsThisHelp {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Activation` value is invalid.
    #[error("The `Activation` value is invalid: '{value}'. Only 0 (Disabled) or -1 (Enabled) are valid values.")]
    InvalidActivation {
        /// The invalid value that was found.
        value: String,
    },

    /// The `LinkMode` value is invalid (form-specific).
    #[error("The `LinkMode` value is invalid: '{value}'. Only 0 (None) or 1 (Source).")]
    InvalidFormLinkMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `WindowState` value is invalid.
    #[error("The `WindowState` value is invalid: '{value}'. Only 0 (Normal), 1 (Minimized), or 2 (Maximized) are valid values.")]
    InvalidWindowState {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Align` value is invalid.
    #[error("The `Align` value is invalid: '{value}'. Only 0 (None), 1 (Top), 2 (Bottom), 3 (Left), or 4 (Right) are valid values.")]
    InvalidAlign {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Appearance` value is invalid.
    #[error("The `Appearance` value is invalid: '{value}'. Only 0 (Flat) or 1 (ThreeD) are valid values.")]
    InvalidAppearance {
        /// The invalid value that was found.
        value: String,
    },

    /// The `JustifyAlignment` value is invalid.
    #[error("The `JustifyAlignment` value is invalid: '{value}'. Only 0 (Left), 1 (Right) are valid values.")]
    InvalidJustifyAlignment {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Alignment` value is invalid.
    #[error("The `Alignment` value is invalid: '{value}'. Only 0 (Left), 1 (Center), or 2 (Right) are valid values.")]
    InvalidAlignment {
        /// The invalid value that was found.
        value: String,
    },

    /// The `BackStyle` value is invalid.
    #[error("The `BackStyle` value is invalid: '{value}'. Only 0 (Transparent) or 1 (Opaque) are valid values.")]
    InvalidBackStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `BorderStyle` value is invalid.
    #[error("The `BorderStyle` value is invalid: '{value}'. Only 0 (None) or 1 (FixedSingle) are valid values.")]
    InvalidBorderStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DragMode` value is invalid.
    #[error("The `DragMode` value is invalid: '{value}'. Only 0 (Manual) or 1 (Automatic) are valid values.")]
    InvalidDragMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DrawMode` value is invalid.
    #[error("The `DrawMode` value is invalid: '{value}'. Only 1 (Blackness), 2 (NotMergePen), 3 (MaskNotPen), 4 (NotCopyPen), 5 (MaskPenNot), 6 (Invert), 7 (XorPen), 8 (NotMaskPen), 9 (MaskPen), 10 (NotXorPen), 11 (Nop), 12 (MergeNotPen), 13 (CopyPen), 14 (MergePenNot), 15 (Merge Pen), or 16 (Whiteness) are valid values.")]
    InvalidDrawMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DrawStyle` value is invalid.
    #[error("The `DrawStyle` value is invalid: '{value}'. Only 0 (Solid), 1 (Dash), 2 (Dot), 3 (DashDot), 4 (DashDotDot), 5 (Transparent), or 6 (InsideSolid) are valid values.")]
    InvalidDrawStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MousePointer` value is invalid.
    #[error("The `MousePointer` value is invalid: '{value}'. Only 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 4 (Icon), 5 (Size), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom) are valid values.")]
    InvalidMousePointer {
        /// The invalid value that was found.
        value: String,
    },

    /// The `OLEDragMode` value is invalid.
    #[error("The `OLEDragMode` value is invalid: '{value}'. Only 0 (Manual), or 1 (Automatic) are valid values.")]
    InvalidOLEDragMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `OLEDropMode` value is invalid.
    #[error("The `OLEDropMode` value is invalid: '{value}'. Only 0 (None), or 1 (Manual) are valid values.")]
    InvalidOLEDropMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ClipControls` value is invalid.
    #[error("The `ClipControls` value is invalid: '{value}'. Only 0 (Unbounded) or 1 (Clipped) are valid values.")]
    InvalidClipControls {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Style` value is invalid.
    #[error("The `Style` value is invalid: '{value}'. Only 0 (Standard) or 1 (Graphical) are valid values.")]
    InvalidStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `FillStyle` value is invalid.
    #[error("The `FillStyle` value is invalid: '{value}'. Only 0 (Solid), 1 (Transparent), 2 (HorizontalLine), 3 (VerticalLine), 4 (UpwardDiagonal), 5 (DownwardDiagonal), 6 (Cross), or 7 (DiagonalCross) are valid values.")]
    InvalidFillStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `LinkMode` value is invalid.
    #[error("The `LinkMode` value is invalid: '{value}'. Only 0 (None), 1 (Automatic), 2 (Manual), or 3 (Notify) are valid values.")]
    InvalidLinkMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MultiSelect` value is invalid.
    #[error("The `MultiSelect` value is invalid: '{value}'. Only 0 (None), 1 (Simple), or 2 (Extended) are valid values.")]
    InvalidMultiSelect {
        /// The invalid value that was found.
        value: String,
    },

    /// The `OLETypeAllowed` value is invalid.
    #[error("The `OLETypeAllowed` value is invalid: '{value}'. Only 0 (Link), 1 (Embedded), or 2 (Either) are valid values.")]
    InvalidOLETypeAllowed {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ScaleMode` value is invalid.
    #[error("The `ScaleMode` value is invalid: '{value}'. Only 0 (User), 1 (Twips), 2 (Points), 3 (Pixels), 4 (Characters), 5 (Inches), 6 (Millimeters), 7 (Centimeters), 8 (HiMetric), 9 (ContainerPosition), 10 (ContainerSize) are valid values.")]
    InvalidScaleMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `SizeMode` value is invalid.
    #[error("The `SizeMode` value is invalid: '{value}'. Only 0 (Clip), 1 (Stretch), 2 (AutoSize), or 3 (Zoom) are valid values.")]
    InvalidSizeMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `OptionButtonValue` value is invalid.
    #[error("The `OptionButtonValue` value is invalid: '{value}'. Only 0 (UnSelected), or 1 (Selected) are valid values.")]
    InvalidOptionButtonValue {
        /// The invalid value that was found.
        value: String,
    },

    /// The `UpdateOptions` value is invalid.
    #[error("The `UpdateOptions` value is invalid: '{value}'. Only 0 (Automatic), 1 (Frozen), or 2 (Manual) are valid values.")]
    InvalidUpdateOptions {
        /// The invalid value that was found.
        value: String,
    },

    /// The `AutoActivate` value is invalid.
    #[error("The `AutoActivate` value is invalid: '{value}'. Only 0 (Manual), 1 (GetFocus), 2 (DoubleClick), or 3 (Automatic) are valid values.")]
    InvalidAutoActivate {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DisplayType` value is invalid.
    #[error("The `DisplayType` value is invalid: '{value}'. Only 0 (Content) or 1 (Icon) are valid values.")]
    InvalidDisplayType {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ScrollBars` value is invalid.
    #[error("The `ScrollBars` value is invalid: '{value}'. Only 0 (None), 1 (Horizontal), 2 (Vertical), or 3 (Both) are valid values.")]
    InvalidScrollBars {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MultiLine` value is invalid.
    #[error("The `MultiLine` value is invalid: '{value}'. Only 0 (SingleLine) or -1 (MultiLine) are valid values.")]
    InvalidMultiLine {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Shape` value is invalid.
    #[error("The `Shape` value is invalid: '{value}'. Only 0 (Rectangle), 1 (Square), 2 (Oval), 3 (Circle), 4 (RoundedRectangle), or 5 (RoundSquare) are valid values.")]
    InvalidShape {
        /// The invalid value that was found.
        value: String,
    },

    /// The 'VERSION' keyword is missing from the form file header.
    #[error("The 'VERSION' keyword is missing from the form file header.")]
    VersionKeywordMissing,

    /// The 'Begin' keyword is missing from the form file header.
    #[error("The 'Begin' keyword is missing from the form file header.")]
    BeginKeywordMissing,

    /// The `Form` is missing from the form file header.
    #[error("The Form is missing from the form file header.")]
    Missing,

    /// `Property` parsing error.
    #[error("Property parsing error")]
    PropertyError,

    /// `Resource` file parsing error.
    #[error("Resource file parsing error: {message}")]
    ResourceFileError {
        /// Error message from resource file parsing.
        message: String,
    },

    /// Error reading the source file.
    #[error("Error reading the source file: {message}")]
    SourceFileError {
        /// Error message from source file reading.
        message: String,
    },

    /// The file contains non-English character set.
    #[error("The file contains more than a significant number of non-ASCII characters. This file was likely saved in a non-English character set. The vb6parse crate currently does not support non-english vb6 files.")]
    LikelyNonEnglishCharacterSet,

    /// The reference line has too many elements.
    #[error("The reference line has too many elements")]
    ReferenceExtraSections,

    /// The reference line has too few elements.
    #[error("The reference line has too few elements")]
    ReferenceMissingSections,

    /// The first line must be a project 'Type' entry.
    #[error("The first line of a VB6 project file must be a project 'Type' entry.")]
    FirstLineNotProject,

    /// Line type is unknown.
    #[error("Line type is unknown.")]
    LineTypeUnknown,

    /// Project type is unknown.
    #[error("Project type is not Exe, OleDll, Control, or OleExe")]
    ProjectTypeUnknown,

    /// Project lacks a version number.
    #[error("Project lacks a version number.")]
    NoVersion,

    /// Project parse error while processing an Object line.
    #[error("Project parse error while processing an Object line.")]
    NoObjects,

    /// Form parse error. No Form found in form file.
    #[error("Form parse error. No Form found in form file.")]
    NoForm,

    /// Parse error while processing Form attributes.
    #[error("Parse error while processing Form attributes.")]
    AttributeParseError,

    /// Parse error while attempting to parse Form tokens.
    #[error("Parse error while attempting to parse Form tokens.")]
    TokenParseError,

    /// Project parse error, failure to find BEGIN element.
    #[error("Project parse error, failure to find BEGIN element.")]
    NoBegin,

    /// Project line entry is not ended with a recognized line ending.
    #[error("Project line entry is not ended with a recognized line ending.")]
    NoLineEnding,

    /// Unable to parse the Uuid.
    #[error("Unable to parse the Uuid")]
    UnableToParseUuid,

    /// Unable to find a semicolon ';' in this line.
    #[error("Unable to find a semicolon ';' in this line.")]
    NoSemicolonSplit,

    /// Unable to find an equal '=' in this line.
    #[error("Unable to find an equal '=' in this line.")]
    NoEqualSplit,

    /// While trying to parse the offset into the resource file, no colon ':' was found.
    #[error("While trying to parse the offset into the resource file, no colon ':' was found.")]
    NoColonForOffsetSplit,

    /// No key value divider found in the line.
    #[error("No key value divider found in the line.")]
    NoKeyValueDividerFound,

    /// Unknown parser error.
    #[error("Unknown parser error")]
    Unparsable,

    /// Major version is not a number.
    #[error("Major version is not a number.")]
    MajorVersionUnparsable,

    /// Unable to parse hex address from `DllBaseAddress` key.
    #[error("Unable to parse hex address from DllBaseAddress key")]
    DllBaseAddressUnparsable,

    /// The Startup object is not a valid parameter.
    #[error("The Startup object is not a valid parameter. Must be a quoted startup method/object, \"(None)\", !(None)!, \"\", or \"!!\"")]
    StartupUnparsable,

    /// The Name parameter is invalid.
    #[error("The Name parameter is invalid. Must be a quoted name, \"(None)\", !(None)!, \"\", or \"!!\"")]
    NameUnparsable,

    /// The `CommandLine` parameter is invalid.
    #[error("The CommandLine parameter is invalid. Must be a quoted command line, \"(None)\", !(None)!, \"\", or \"!!\"")]
    CommandLineUnparsable,

    /// The `HelpContextId` parameter is not a valid parameter line.
    #[error("The HelpContextId parameter is not a valid parameter line. Must be a quoted help context id, \"(None)\", !(None)!, \"\", or \"!!\"")]
    HelpContextIdUnparsable,

    /// Minor version is not a number.
    #[error("Minor version is not a number.")]
    MinorVersionUnparsable,

    /// Revision version is not a number.
    #[error("Revision version is not a number.")]
    RevisionVersionUnparsable,

    /// Unable to parse the value after `ThreadingModel` key.
    #[error("Unable to parse the value after ThreadingModel key")]
    ThreadingModelUnparsable,

    /// `ThreadingModel` can only be 0 or 1.
    #[error("ThreadingModel can only be 0 (Apartment Threaded), or 1 (Single Threaded)")]
    ThreadingModelInvalid,

    /// No property name found after `BeginProperty` keyword.
    #[error("No property name found after BeginProperty keyword.")]
    NoPropertyName,

    /// Unable to parse the `RelatedDoc` property line.
    #[error("Unable to parse the RelatedDoc property line.")]
    RelatedDocLineUnparsable,

    /// `AutoIncrement` can only be 0 or -1.
    #[error("AutoIncrement can only be a 0 (false) or a -1 (true)")]
    AutoIncrementUnparsable,

    /// `CompatibilityMode` value is invalid.
    #[error("CompatibilityMode can only be a 0 (CompatibilityMode::NoCompatibility), 1 (CompatibilityMode::Project), or 2 (CompatibilityMode::CompatibleExe)")]
    CompatibilityModeUnparsable,

    /// `NoControlUpgrade` value is invalid.
    #[error("NoControlUpgrade can only be a 0 (UpgradeControls::Upgrade) or a 1 (UpgradeControls::NoUpgrade)")]
    NoControlUpgradeUnparsable,

    /// `ServerSupportFiles` can only be 0 or -1.
    #[error("ServerSupportFiles can only be a 0 (false) or a -1 (true)")]
    ServerSupportFilesUnparsable,

    /// `Comment` line was unparsable.
    #[error("Comment line was unparsable")]
    CommentUnparsable,

    /// `PropertyPage` line was unparsable.
    #[error("PropertyPage line was unparsable")]
    PropertyPageUnparsable,

    /// `CompilationType` can only be 0 or -1.
    #[error("CompilationType can only be a 0 (false) or a -1 (true)")]
    CompilationTypeUnparsable,

    /// `OptimizationType` value is invalid.
    #[error("OptimizationType can only be a 0 (FastCode) or 1 (SmallCode), or 2 (NoOptimization)")]
    OptimizationTypeUnparsable,

    /// `FavorPentiumPro(tm)` can only be 0 or -1.
    #[error("FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)")]
    FavorPentiumProUnparsable,

    /// `Designer` line is unparsable.
    #[error("Designer line is unparsable")]
    DesignerLineUnparsable,

    /// Form line is unparsable.
    #[error("Form line is unparsable")]
    FormLineUnparsable,

    /// `UserControl` line is unparsable.
    #[error("UserControl line is unparsable")]
    UserControlLineUnparsable,

    /// `UserDocument` line is unparsable.
    #[error("UserDocument line is unparsable")]
    UserDocumentLineUnparsable,

    /// Period expected in version number.
    #[error("Period expected in version number")]
    PeriodExpectedInVersionNumber,

    /// `CodeViewDebugInfo` can only be 0 or -1.
    #[error("CodeViewDebugInfo can only be a 0 (false) or a -1 (true)")]
    CodeViewDebugInfoUnparsable,

    /// `NoAliasing` can only be 0 or -1.
    #[error("NoAliasing can only be a 0 (false) or a -1 (true)")]
    NoAliasingUnparsable,

    /// `RemoveUnusedControlInfo` value is invalid.
    #[error("RemoveUnusedControlInfo can only be 0 (UnusedControlInfo::Retain) or -1 (UnusedControlInfo::Remove)")]
    UnusedControlInfoUnparsable,

    /// `BoundsCheck` can only be 0 or -1.
    #[error("BoundsCheck can only be a 0 (false) or a -1 (true)")]
    BoundsCheckUnparsable,

    /// `OverflowCheck` can only be 0 or -1.
    #[error("OverflowCheck can only be a 0 (false) or a -1 (true)")]
    OverflowCheckUnparsable,

    /// `FlPointCheck` can only be 0 or -1.
    #[error("FlPointCheck can only be a 0 (false) or a -1 (true)")]
    FlPointCheckUnparsable,

    /// `FDIVCheck` value is invalid.
    #[error("FDIVCheck can only be a 0 (PentiumFDivBugCheck::CheckPentiumFDivBug) or a -1 (PentiumFDivBugCheck::NoPentiumFDivBugCheck)")]
    FDIVCheckUnparsable,

    /// `UnroundedFP` value is invalid.
    #[error("UnroundedFP can only be a 0 (UnroundedFloatingPoint::DoNotAllow) or a -1 (UnroundedFloatingPoint::Allow)")]
    UnroundedFPUnparsable,

    /// `StartMode` value is invalid.
    #[error("StartMode can only be a 0 (StartMode::StandAlone) or a 1 (StartMode::Automation)")]
    StartModeUnparsable,

    /// `Unattended` value is invalid.
    #[error("Unattended can only be a 0 (Unattended::False) or a -1 (Unattended::True)")]
    UnattendedUnparsable,

    /// `Retained` value is invalid.
    #[error(
        "Retained can only be a 0 (Retained::UnloadOnExit) or a 1 (Retained::RetainedInMemory)"
    )]
    RetainedUnparsable,

    /// Unable to parse the `ShortCut` property.
    #[error("Unable to parse the ShortCut property.")]
    ShortCutUnparsable,

    /// `DebugStartup` can only be 0 or -1.
    #[error("DebugStartup can only be a 0 (false) or a -1 (true)")]
    DebugStartupOptionUnparsable,

    /// `UseExistingBrowser` value is invalid.
    #[error("UseExistingBrowser can only be a 0 (UseExistingBrowser::DoNotUse) or a -1 (UseExistingBrowser::Use)")]
    UseExistingBrowserUnparsable,

    /// `AutoRefresh` can only be 0 or -1.
    #[error("AutoRefresh can only be a 0 (false) or a -1 (true)")]
    AutoRefreshUnparsable,

    /// `Thread Per Object` is not a number.
    #[error("Thread Per Object is not a number.")]
    ThreadPerObjectUnparsable,

    /// Unknown attribute in class header file.
    #[error("Unknown attribute in class header file. Must be one of: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_Ext_KEY")]
    UnknownAttribute,

    /// Error parsing header.
    #[error("Error parsing header")]
    Header,

    /// No name in the attribute section of the VB6 file.
    #[error("No name in the attribute section of the VB6 file")]
    MissingNameAttribute,

    /// Keyword not found.
    #[error("Keyword not found")]
    KeywordNotFound,

    /// Error parsing true/false from header.
    #[error("Error parsing true/false from header. Must be a 0 (false), -1 (true), or 1 (true)")]
    TrueFalseOneZeroNegOneUnparsable,

    /// Error parsing the VB6 file contents.
    #[error("Error parsing the VB6 file contents")]
    FileContent,

    /// Max Threads is not a number.
    #[error("Max Threads is not a number.")]
    MaxThreadsUnparsable,

    /// No `EndProperty` found after `BeginProperty`.
    #[error("No EndProperty found after BeginProperty")]
    NoEndProperty,

    /// No line ending after `EndProperty`.
    #[error("No line ending after EndProperty")]
    NoLineEndingAfterEndProperty,

    /// Expected namespace after `Begin` keyword.
    #[error("Expected namespace after Begin keyword")]
    NoNamespaceAfterBegin,

    /// No dot found after namespace.
    #[error("No dot found after namespace")]
    NoDotAfterNamespace,

    /// No User Control name found after namespace and '.'.
    #[error("No User Control name found after namespace and '.'")]
    NoUserControlNameAfterDot,

    /// No space after Control kind.
    #[error("No space after Control kind")]
    NoSpaceAfterControlKind,

    /// No control name found after Control kind.
    #[error("No control name found after Control kind")]
    NoControlNameAfterControlKind,

    /// No line ending after Control name.
    #[error("No line ending after Control name")]
    NoLineEndingAfterControlName,

    /// Unknown token in form parsing.
    #[error("Unknown token")]
    UnknownToken,

    /// Title text was unparsable.
    #[error("Title text was unparsable")]
    TitleUnparsable,

    /// Unable to parse hex color value.
    #[error("Unable to parse hex color value")]
    HexColorParseError,

    /// Unknown control in control list.
    #[error("Unknown control in control list")]
    UnknownControlKind,

    /// Property name is not a valid ASCII string.
    #[error("Property name is not a valid ASCII string")]
    PropertyNameAsciiConversionError,

    /// String is unterminated.
    #[error("String is unterminated")]
    UnterminatedString,

    /// Unable to parse VB6 string.
    #[error("Unable to parse VB6 string.")]
    StringParseError,

    /// Property value is not a valid ASCII string.
    #[error("Property value is not a valid ASCII string")]
    PropertyValueAsciiConversionError,

    /// Key value pair format is incorrect.
    #[error("Key value pair format is incorrect")]
    KeyValueParseError,

    /// Namespace is not a valid ASCII string.
    #[error("Namespace is not a valid ASCII string")]
    NamespaceAsciiConversionError,

    /// Control kind is not a valid ASCII string.
    #[error("Control kind is not a valid ASCII string")]
    ControlKindAsciiConversionError,

    /// Qualified control name is not a valid ASCII string.
    #[error("Qualified control name is not a valid ASCII string")]
    QualifiedControlNameAsciiConversionError,

    /// Variable names must be less than 255 characters in VB6.
    #[error("Variable names must be less than 255 characters in VB6.")]
    VariableNameTooLong,

    /// Invalid top-level control type.
    #[error("Invalid top-level control type: '{control_type}'. Form files must have either 'VB.Form' or 'VB.MDIForm' as the top-level element.")]
    InvalidTopLevelControl {
        /// The invalid control type that was found.
        control_type: String,
    },

    /// Internal Parser Error.
    #[error("Internal Parser Error - please report this issue to the developers.")]
    InternalParseError,
}
