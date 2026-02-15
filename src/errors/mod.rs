//! Module containing the error types used in the VB6 parser.
//!
//! This module is organized by the layer or file type that produces the error:
//! - [`decode`] - Source file decoding errors (Windows-1252, malformed files)
//! - [`tokenize`] - Tokenization and basic code parsing errors
//! - [`resource`] - Form resource file (FRX) parsing errors
//! - [`class`] - Class file (CLS) specific errors
//! - [`module`] - Module file (BAS) specific errors
//! - [`project`] - Project file (VBP) specific errors
//! - [`form`] - Form file (FRM) specific errors
//! - [`property`] - Property value validation errors
//!
//! The [`ErrorDetails`] type is the central error container that wraps any of these
//! error kinds along with source location information for diagnostic reporting.

use ariadne::{Label, Report, ReportKind, Source};
use core::convert::From;
use std::error::Error;
use std::fmt::{Debug, Display};

/// Unified error kind enum that replaces all file-type-specific error kinds.
///
/// This enum consolidates all parsing errors into a single type, eliminating
/// the need for generic type parameters on `ErrorDetails` and `ParseResult`.
/// It contains ~236 variants organized by the layer or file type that produces them:
///
/// - **Code/Tokenization** - Token parsing and basic syntax errors
/// - **Class Files** - Class file (.cls) specific parsing errors  
/// - **Module Files** - Module file (.bas) specific parsing errors
/// - **Form Files** - Form file (.frm) validation and parsing errors
/// - **Project Files** - Project file (.vbp) parsing errors
/// - **Resource Files** - Resource file (.frx) binary data errors
/// - **Source Decoding** - File encoding and decoding errors
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ErrorKind {
    // ========================================================================
    // CODE / TOKENIZATION ERRORS
    // ========================================================================
    /// Variable names in VB6 have a maximum length of 255 characters.
    #[error("Variable names in VB6 have a maximum length of 255 characters.")]
    VariableNameTooLong,

    /// Unknown token encountered during parsing.
    #[error("Unknown token '{token}' found.")]
    UnknownToken {
        /// The unknown token that was encountered.
        token: String,
    },

    /// Unexpected end of the code stream.
    #[error("Unexpected end of code stream.")]
    UnexpectedEndOfStream,

    // ========================================================================
    // CLASS FILE ERRORS
    // ========================================================================
    /// The 'VERSION' keyword is missing from the class file header.
    #[error("The 'VERSION' keyword is missing from the class file header.")]
    ClassVersionKeywordMissing,

    /// The 'BEGIN' keyword is missing from the class file header.
    #[error("The 'BEGIN' keyword is missing from the class file header.")]
    ClassBeginKeywordMissing,

    /// The 'Class' keyword is missing from the class file header.
    #[error("The 'Class' keyword is missing from the class file header.")]
    ClassKeywordMissing,

    /// Missing whitespace between 'VERSION' keyword and major version number.
    #[error(
        "After the 'VERSION' keyword there should be a space before the major version number."
    )]
    ClassWhitespaceMissingBetweenVersionAndMajorVersionNumber,

    /// The 'VERSION' keyword is not fully uppercase.
    #[error("The 'VERSION' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    ClassVersionKeywordNotFullyUppercase {
        /// The text of the 'VERSION' keyword as found in the source.
        version_text: String,
    },

    /// The 'CLASS' keyword is not fully uppercase.
    #[error("The 'CLASS' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    ClassKeywordNotFullyUppercase {
        /// The text of the 'CLASS' keyword as found in the source.
        class_text: String,
    },

    /// The 'BEGIN' keyword is not fully uppercase.
    #[error("The 'BEGIN' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    ClassBeginKeywordNotFullyUppercase {
        /// The text of the 'BEGIN' keyword as found in the source.
        begin_text: String,
    },

    /// The 'END' keyword is not fully uppercase.
    #[error(
        "The 'END' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE."
    )]
    ClassEndKeywordNotFullyUppercase {
        /// The text of the 'END' keyword as found in the source.
        end_text: String,
    },

    /// The 'BEGIN' keyword should be on its own line.
    #[error("The 'BEGIN' keyword should stand alone on its own line.")]
    ClassBeginKeywordShouldBeStandAlone,

    /// The 'END' keyword should be on its own line.
    #[error("The 'END' keyword should stand alone on its own line.")]
    ClassEndKeywordShouldBeStandAlone,

    /// Unable to parse the major version number.
    #[error("Unable to parse the major version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    ClassUnableToParseMajorVersionNumber,

    /// Unable to convert the major version text to a number.
    #[error("Unable to convert the major version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    ClassUnableToConvertMajorVersionNumber,

    /// Unable to parse the minor version number.
    #[error("Unable to parse the minor version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    ClassUnableToParseMinorVersionNumber,

    /// Unable to convert the minor version text to a number.
    #[error("Unable to convert the minor version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    ClassUnableToConvertMinorVersionNumber,

    /// The period divider between major and minor version digits is missing.
    #[error("The '.' divider between major and minor version digits is missing.")]
    ClassMissingPeriodDividerBetweenMajorAndMinorVersion,

    /// Missing whitespace between minor version digits and 'CLASS' keyword.
    #[error("Missing whitespace between minor version digits and 'CLASS' keyword. This may not be compliant with Microsoft's VB6 IDE.")]
    ClassMissingWhitespaceAfterMinorVersion,

    /// Incorrect whitespace between minor version digits and 'CLASS' keyword.
    #[error("Between the minor version digits and the 'CLASS' keyword should be a single ASCII space. This may not be compliant with Microsoft's VB6 IDE.")]
    ClassIncorrectWhitespaceAfterMinorVersion,

    /// Whitespace was used to divide between major and minor version numbers.
    #[error("Whitespace was used to divide between major and minor version information. This may not be compliant with Microsoft's VB6 IDE.")]
    ClassWhitespaceDividerBetweenMajorAndMinorVersionNumbers,

    /// CST parsing error in class file.
    #[error("CST parsing error: {message}")]
    ClassCSTError {
        /// Error message from CST parsing.
        message: String,
    },

    // ========================================================================
    // MODULE FILE ERRORS
    // ========================================================================
    /// The 'Attribute' keyword is missing from the module file header.
    #[error("The 'Attribute' keyword is missing from the module file header.")]
    ModuleAttributeKeywordMissing,

    /// Missing whitespace in module header.
    #[error("The 'Attribute' keyword and the 'VB_Name' attribute must be separated by at least one ASCII whitespace character.")]
    ModuleMissingWhitespaceInHeader,

    /// The `VB_Name` attribute is missing from the module file header.
    #[error("The 'VB_Name' attribute is missing from the module file header.")]
    ModuleVBNameAttributeMissing,

    /// The `VB_Name` attribute is missing the equal symbol.
    #[error("The 'VB_Name' attribute is missing the equal symbol from the module file header.")]
    ModuleEqualMissing,

    /// The `VB_Name` attribute value is unquoted.
    #[error("The 'VB_Name' attribute is unquoted.")]
    ModuleVBNameAttributeValueUnquoted,

    // ========================================================================
    // FORM FILE ERRORS (161 variants)
    // ========================================================================
    /// The `ComboBox` style is invalid.
    #[error("The `ComboBox` style is invalid: '{value}'. Only 0 (Dropdown Combo), 1 (Simple Combo), or 2 (Dropdown List) are valid styles.")]
    FormInvalidComboBoxStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `CheckBox` value is invalid.
    #[error("The `CheckBox` value is invalid: '{value}'. Only 0 (Unchecked), 1 (Checked), or 2 (Grayed) are valid values.")]
    FormInvalidCheckBoxValue {
        /// The invalid value that was found.
        value: String,
    },

    /// The `BOFAction` property has an invalid value.
    #[error("The `BOFAction` value is invalid: '{value}'. Only 0 (MoveFirst), or 1 (BOF) are valid values.")]
    FormInvalidBOFAction {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ConnectionType` property has an invalid value.
    #[error("The `ConnectionType` value is invalid: '{value}'. Only 'Access', 'dBase III', 'dBase IV', 'dBase 5.0', 'Excel 3.0', 'Excel 4.0', 'Excel 5.0', 'Excel 8.0', 'FoxPro 2.0', 'FoxPro 2.5', 'FoxPro 2.6', 'FoxPro 3.0', 'Lotus WK1', 'Lotus WK3', 'Lotus WK4', 'Paradox 3.X', 'Paradox 4.X', 'Paradox 5.X', or 'Text' are valid values.")]
    FormInvalidConnectionType {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DefaultCursorType` property has an invalid value.
    #[error("The `DefaultCursorType` value is invalid: '{value}'. Only 0 (Default), 1 (Odbc), or 2 (ServerSide) are valid values.")]
    FormInvalidDefaultCursorType {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DatabaseDriverType` property has an invalid value.
    #[error("The `DatabaseDriverType` value is invalid: '{value}'. Only 1 (ODBC), or 2 (Jet) are valid values.")]
    FormInvalidDatabaseDriverType {
        /// The invalid value that was found.
        value: String,
    },

    /// The `EOFAction` property has an invalid value.
    #[error("The `EOFAction` value is invalid: '{value}'. Only 0 (MoveLast), 1 (EOF), or 2 (AddNew) are valid values.")]
    FormInvalidEOFAction {
        /// The invalid value that was found.
        value: String,
    },

    /// The record set type is invalid.
    #[error("The `RecordSetType` value is invalid: '{value}'. Only 0 (Table), 1 (Dynaset), or 2 (Snapshot) are valid values.")]
    FormInvalidRecordSetType {
        /// The invalid value that was found.
        value: String,
    },

    /// The archive attribute is invalid.
    #[error("The `ArchiveAttribute` value is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    FormInvalidArchiveAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The hidden attribute is invalid.
    #[error("The `Hidden` valud is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    FormInvalidHiddenAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ReadOnly` attribute is invalid.
    #[error("The `ReadOnly` value is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    FormInvalidReadOnlyAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The `System` attribute is invalid.
    #[error("The `System` value is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    FormInvalidSystemAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Normal` attribute is invalid.
    #[error("The `Normal` value is invalid: '{value}'. Only 0 (Exclude) or -1 (Include) are valid values.")]
    FormInvalidNormalAttribute {
        /// The invalid value that was found.
        value: String,
    },

    /// The `FormBorderStyle` property has an invalid value.
    #[error("The `FormBorderStyle` value is invalid: '{value}'. Only 0 (None), 1 (FixedSingle), 2 (Sizable), 3 (FixedDialog), 4 (FixedToolWindow), or 5 (SizableToolWindow) are valid values.")]
    FormInvalidFormBorderStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ControlBox` property has an invalid value.
    #[error("The `ControlBox` value is invalid: '{value}'. Only 0 (Excluded) or -1 (Included) are valid values.")]
    FormInvalidControlBox {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MaxButton` property has an invalid value.
    #[error("The `MaxButton` value is invalid: '{value}'. Only 0 (Excluded) or -1 (Included) are valid values.")]
    FormInvalidMaxButton {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MinButton` property has an invalid value.
    #[error("The `MinButton` value is invalid: '{value}'. Only 0 (Excluded) or -1 (Included) are valid values.")]
    FormInvalidMinButton {
        /// The invalid value that was found.
        value: String,
    },

    /// The `PaletteMode` property has an invalid value.
    #[error("The `PaletteMode` value is invalid: '{value}'. Only 0 (HalfTone), 1 (UseZOrder), or 2 (Custom) are valid values.")]
    FormInvalidPaletteMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `WordWrap` property has an invalid value.
    #[error(
        "The `WordWrap` value is invalid: '{value}'. Only 0 (False) or -1 (True) are valid values."
    )]
    FormInvalidWordWrap {
        /// The invalid value that was found.
        value: String,
    },

    /// The `WhatsThisButton` value is invalid.
    #[error("The `WhatsThisButton` value is invalid: '{value}'. Only 0 (Excluded) or -1 (Included) are valid values.")]
    FormInvalidWhatsThisButton {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ShowInTaskbar` value is invalid.
    #[error("The `ShowInTaskbar` value is invalid: '{value}'. Only 0 (Hide) or -1 (Show) are valid values.")]
    FormInvalidShowInTaskbar {
        /// The invalid value that was found.
        value: String,
    },

    /// The `NegotiatePosition` value is invalid.
    #[error("The `NegotiatePosition` value is invalid: '{value}'. Only 0 (None), 1 (Left), 2 (Middle), or 3 (Right) are valid values.")]
    FormInvalidNegotiatePosition {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ListBoxStyle` value is invalid.
    #[error("The `ListBoxStyle` value is invalid: '{value}'. Only 0 (Standard) or 1 (Checkbox) are valid values.")]
    FormInvalidListBoxStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `AutoSize` value is invalid.
    #[error("The `AutoSize` value is invalid: '{value}'. Only 0 (Fixed) or -1 (Resize) are valid values.")]
    FormInvalidAutoSize {
        /// The invalid value that was found.
        value: String,
    },

    /// The `AutoRedraw` value is invalid.
    #[error("The `AutoRedraw` value is invalid: '{value}'. Only 0 (Manual) or -1 (Automatic) are valid values.")]
    FormInvalidAutoRedraw {
        /// The invalid value that was found.
        value: String,
    },

    /// The `TextDirection` value is invalid.
    #[error("The `TextAlign` value is invalid: '{value}'. Only 0 (LeftToRight) or -1 (RightToLeft) are valid values.")]
    FormInvalidTextDirection {
        /// The invalid value that was found.
        value: String,
    },

    /// The `TabStop` value is invalid.
    #[error("The `TabStop` value is invalid: '{value}'. Only 0 (ProgrammaticOnly) or -1 (Included) are valid values.")]
    FormInvalidTabStop {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Visibility` value is invalid.
    #[error("The `Visibility` value is invalid: '{value}'. Only 0 (Hidden) or -1 (Visible) are valid values.")]
    FormInvalidVisibility {
        /// The invalid value that was found.
        value: String,
    },

    /// The `HasDeviceContext` value is invalid.
    #[error("The `HasDeviceContext` value is invalid: '{value}'. Only 0 (No) or -1 (Yes) are valid values.")]
    FormInvalidHasDeviceContext {
        /// The invalid value that was found.
        value: String,
    },

    /// The `CausesValidation` value is invalid.
    #[error("The `CausesValidation` value is invalid: '{value}'. Only 0 (No) or -1 (Yes) are valid values.")]
    FormInvalidCausesValidation {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Movability` value is invalid.
    #[error("The `Movability` value is invalid: '{value}'. Only 0 (Fixed) or -1 (Movable) are valid values.")]
    FormInvalidMovability {
        /// The invalid value that was found.
        value: String,
    },

    /// The `FontTransparency` value is invalid.
    #[error("The `FontTransparency` value is invalid: '{value}'. Only 0 (Opaque) or -1 (Transparent) are valid values.")]
    FormInvalidFontTransparency {
        /// The invalid value that was found.
        value: String,
    },

    /// The `WhatsThisHelp` value is invalid.
    #[error("The `WhatsThisHelp` value is invalid: '{value}'. Only 0 (F1Help) or -1 (WhatsThisHelp) are valid values.")]
    FormInvalidWhatsThisHelp {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Activation` value is invalid.
    #[error("The `Activation` value is invalid: '{value}'. Only 0 (Disabled) or -1 (Enabled) are valid values.")]
    FormInvalidActivation {
        /// The invalid value that was found.
        value: String,
    },

    /// The `LinkMode` value is invalid (form-specific).
    #[error("The `LinkMode` value is invalid: '{value}'. Only 0 (None) or 1 (Source).")]
    FormInvalidFormLinkMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `WindowState` value is invalid.
    #[error("The `WindowState` value is invalid: '{value}'. Only 0 (Normal), 1 (Minimized), or 2 (Maximized) are valid values.")]
    FormInvalidWindowState {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Align` value is invalid.
    #[error("The `Align` value is invalid: '{value}'. Only 0 (None), 1 (Top), 2 (Bottom), 3 (Left), or 4 (Right) are valid values.")]
    FormInvalidAlign {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Appearance` value is invalid.
    #[error("The `Appearance` value is invalid: '{value}'. Only 0 (Flat) or 1 (ThreeD) are valid values.")]
    FormInvalidAppearance {
        /// The invalid value that was found.
        value: String,
    },

    /// The `JustifyAlignment` value is invalid.
    #[error("The `JustifyAlignment` value is invalid: '{value}'. Only 0 (Left), 1 (Right) are valid values.")]
    FormInvalidJustifyAlignment {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Alignment` value is invalid.
    #[error("The `Alignment` value is invalid: '{value}'. Only 0 (Left), 1 (Center), or 2 (Right) are valid values.")]
    FormInvalidAlignment {
        /// The invalid value that was found.
        value: String,
    },

    /// The `BackStyle` value is invalid.
    #[error("The `BackStyle` value is invalid: '{value}'. Only 0 (Transparent) or 1 (Opaque) are valid values.")]
    FormInvalidBackStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `BorderStyle` value is invalid.
    #[error("The `BorderStyle` value is invalid: '{value}'. Only 0 (None) or 1 (FixedSingle) are valid values.")]
    FormInvalidBorderStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DragMode` value is invalid.
    #[error("The `DragMode` value is invalid: '{value}'. Only 0 (Manual) or 1 (Automatic) are valid values.")]
    FormInvalidDragMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DrawMode` value is invalid.
    #[error("The `DrawMode` value is invalid: '{value}'. Only 1 (Blackness), 2 (NotMergePen), 3 (MaskNotPen), 4 (NotCopyPen), 5 (MaskPenNot), 6 (Invert), 7 (XorPen), 8 (NotMaskPen), 9 (MaskPen), 10 (NotXorPen), 11 (Nop), 12 (MergeNotPen), 13 (CopyPen), 14 (MergePenNot), 15 (Merge Pen), or 16 (Whiteness) are valid values.")]
    FormInvalidDrawMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DrawStyle` value is invalid.
    #[error("The `DrawStyle` value is invalid: '{value}'. Only 0 (Solid), 1 (Dash), 2 (Dot), 3 (DashDot), 4 (DashDotDot), 5 (Transparent), or 6 (InsideSolid) are valid values.")]
    FormInvalidDrawStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MousePointer` value is invalid.
    #[error("The `MousePointer` value is invalid: '{value}'. Only 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 4 (Icon), 5 (Size), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom) are valid values.")]
    FormInvalidMousePointer {
        /// The invalid value that was found.
        value: String,
    },

    /// The `OLEDragMode` value is invalid.
    #[error("The `OLEDragMode` value is invalid: '{value}'. Only 0 (Manual), or 1 (Automatic) are valid values.")]
    FormInvalidOLEDragMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `OLEDropMode` value is invalid.
    #[error("The `OLEDropMode` value is invalid: '{value}'. Only 0 (None), or 1 (Manual) are valid values.")]
    FormInvalidOLEDropMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ClipControls` value is invalid.
    #[error("The `ClipControls` value is invalid: '{value}'. Only 0 (Unbounded) or 1 (Clipped) are valid values.")]
    FormInvalidClipControls {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Style` value is invalid.
    #[error("The `Style` value is invalid: '{value}'. Only 0 (Standard) or 1 (Graphical) are valid values.")]
    FormInvalidStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `FillStyle` value is invalid.
    #[error("The `FillStyle` value is invalid: '{value}'. Only 0 (Solid), 1 (Transparent), 2 (HorizontalLine), 3 (VerticalLine), 4 (UpwardDiagonal), 5 (DownwardDiagonal), 6 (Cross), or 7 (DiagonalCross) are valid values.")]
    FormInvalidFillStyle {
        /// The invalid value that was found.
        value: String,
    },

    /// The `LinkMode` value is invalid.
    #[error("The `LinkMode` value is invalid: '{value}'. Only 0 (None), 1 (Automatic), 2 (Manual), or 3 (Notify) are valid values.")]
    FormInvalidLinkMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MultiSelect` value is invalid.
    #[error("The `MultiSelect` value is invalid: '{value}'. Only 0 (None), 1 (Simple), or 2 (Extended) are valid values.")]
    FormInvalidMultiSelect {
        /// The invalid value that was found.
        value: String,
    },

    /// The `OLETypeAllowed` value is invalid.
    #[error("The `OLETypeAllowed` value is invalid: '{value}'. Only 0 (Link), 1 (Embedded), or 2 (Either) are valid values.")]
    FormInvalidOLETypeAllowed {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ScaleMode` value is invalid.
    #[error("The `ScaleMode` value is invalid: '{value}'. Only 0 (User), 1 (Twips), 2 (Points), 3 (Pixels), 4 (Characters), 5 (Inches), 6 (Millimeters), 7 (Centimeters), 8 (HiMetric), 9 (ContainerPosition), 10 (ContainerSize) are valid values.")]
    FormInvalidScaleMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `SizeMode` value is invalid.
    #[error("The `SizeMode` value is invalid: '{value}'. Only 0 (Clip), 1 (Stretch), 2 (AutoSize), or 3 (Zoom) are valid values.")]
    FormInvalidSizeMode {
        /// The invalid value that was found.
        value: String,
    },

    /// The `OptionButtonValue` value is invalid.
    #[error("The `OptionButtonValue` value is invalid: '{value}'. Only 0 (UnSelected), or 1 (Selected) are valid values.")]
    FormInvalidOptionButtonValue {
        /// The invalid value that was found.
        value: String,
    },

    /// The `UpdateOptions` value is invalid.
    #[error("The `UpdateOptions` value is invalid: '{value}'. Only 0 (Automatic), 1 (Frozen), or 2 (Manual) are valid values.")]
    FormInvalidUpdateOptions {
        /// The invalid value that was found.
        value: String,
    },

    /// The `AutoActivate` value is invalid.
    #[error("The `AutoActivate` value is invalid: '{value}'. Only 0 (Manual), 1 (GetFocus), 2 (DoubleClick), or 3 (Automatic) are valid values.")]
    FormInvalidAutoActivate {
        /// The invalid value that was found.
        value: String,
    },

    /// The `DisplayType` value is invalid.
    #[error("The `DisplayType` value is invalid: '{value}'. Only 0 (Content) or 1 (Icon) are valid values.")]
    FormInvalidDisplayType {
        /// The invalid value that was found.
        value: String,
    },

    /// The `ScrollBars` value is invalid.
    #[error("The `ScrollBars` value is invalid: '{value}'. Only 0 (None), 1 (Horizontal), 2 (Vertical), or 3 (Both) are valid values.")]
    FormInvalidScrollBars {
        /// The invalid value that was found.
        value: String,
    },

    /// The `MultiLine` value is invalid.
    #[error("The `MultiLine` value is invalid: '{value}'. Only 0 (SingleLine) or -1 (MultiLine) are valid values.")]
    FormInvalidMultiLine {
        /// The invalid value that was found.
        value: String,
    },

    /// The `Shape` value is invalid.
    #[error("The `Shape` value is invalid: '{value}'. Only 0 (Rectangle), 1 (Square), 2 (Oval), 3 (Circle), 4 (RoundedRectangle), or 5 (RoundSquare) are valid values.")]
    FormInvalidShape {
        /// The invalid value that was found.
        value: String,
    },

    /// The 'VERSION' keyword is missing from the form file header.
    #[error("The 'VERSION' keyword is missing from the form file header.")]
    FormVersionKeywordMissing,

    /// The 'Begin' keyword is missing from the form file header.
    #[error("The 'Begin' keyword is missing from the form file header.")]
    FormBeginKeywordMissing,

    /// The `Form` is missing from the form file header.
    #[error("The Form is missing from the form file header.")]
    FormMissing,

    /// `Property` parsing error.
    #[error("Property parsing error")]
    FormPropertyError,

    /// `Resource` file parsing error.
    #[error("Resource file parsing error: {message}")]
    FormResourceFileError {
        /// Error message from resource file parsing.
        message: String,
    },

    /// Error reading the source file.
    #[error("Error reading the source file: {message}")]
    FormSourceFileError {
        /// Error message from source file reading.
        message: String,
    },

    /// The file contains non-English character set.
    #[error("The file contains more than a significant number of non-ASCII characters. This file was likely saved in a non-English character set. The vb6parse crate currently does not support non-english vb6 files.")]
    FormLikelyNonEnglishCharacterSet,

    /// The reference line has too many elements.
    #[error("The reference line has too many elements")]
    FormReferenceExtraSections,

    /// The reference line has too few elements.
    #[error("The reference line has too few elements")]
    FormReferenceMissingSections,

    /// The first line must be a project 'Type' entry.
    #[error("The first line of a VB6 project file must be a project 'Type' entry.")]
    FormFirstLineNotProject,

    /// Line type is unknown.
    #[error("Line type is unknown.")]
    FormLineTypeUnknown,

    /// Project type is unknown.
    #[error("Project type is not Exe, OleDll, Control, or OleExe")]
    FormProjectTypeUnknown,

    /// Project lacks a version number.
    #[error("Project lacks a version number.")]
    FormNoVersion,

    /// Project parse error while processing an Object line.
    #[error("Project parse error while processing an Object line.")]
    FormNoObjects,

    /// Form parse error. No Form found in form file.
    #[error("Form parse error. No Form found in form file.")]
    FormNoForm,

    /// Parse error while processing Form attributes.
    #[error("Parse error while processing Form attributes.")]
    FormAttributeParseError,

    /// Parse error while attempting to parse Form tokens.
    #[error("Parse error while attempting to parse Form tokens.")]
    FormTokenParseError,

    /// Project parse error, failure to find BEGIN element.
    #[error("Project parse error, failure to find BEGIN element.")]
    FormNoBegin,

    /// Project line entry is not ended with a recognized line ending.
    #[error("Project line entry is not ended with a recognized line ending.")]
    FormNoLineEnding,

    /// Unable to parse the Uuid.
    #[error("Unable to parse the Uuid")]
    FormUnableToParseUuid,

    /// Unable to find a semicolon ';' in this line.
    #[error("Unable to find a semicolon ';' in this line.")]
    FormNoSemicolonSplit,

    /// Unable to find an equal '=' in this line.
    #[error("Unable to find an equal '=' in this line.")]
    FormNoEqualSplit,

    /// While trying to parse the offset into the resource file, no colon ':' was found.
    #[error("While trying to parse the offset into the resource file, no colon ':' was found.")]
    FormNoColonForOffsetSplit,

    /// No key value divider found in the line.
    #[error("No key value divider found in the line.")]
    FormNoKeyValueDividerFound,

    /// Unknown parser error.
    #[error("Unknown parser error")]
    FormUnparsable,

    /// Major version is not a number.
    #[error("Major version is not a number.")]
    FormMajorVersionUnparsable,

    /// Unable to parse hex address from `DllBaseAddress` key.
    #[error("Unable to parse hex address from DllBaseAddress key")]
    FormDllBaseAddressUnparsable,

    /// The Startup object is not a valid parameter.
    #[error("The Startup object is not a valid parameter. Must be a quoted startup method/object, \"(None)\", !(None)!, \"\", or \"!!\"")]
    FormStartupUnparsable,

    /// The Name parameter is invalid.
    #[error("The Name parameter is invalid. Must be a quoted name, \"(None)\", !(None)!, \"\", or \"!!\"")]
    FormNameUnparsable,

    /// The `CommandLine` parameter is invalid.
    #[error("The CommandLine parameter is invalid. Must be a quoted command line, \"(None)\", !(None)!, \"\", or \"!!\"")]
    FormCommandLineUnparsable,

    /// The `HelpContextId` parameter is not a valid parameter line.
    #[error("The HelpContextId parameter is not a valid parameter line. Must be a quoted help context id, \"(None)\", !(None)!, \"\", or \"!!\"")]
    FormHelpContextIdUnparsable,

    /// Minor version is not a number.
    #[error("Minor version is not a number.")]
    FormMinorVersionUnparsable,

    /// Revision version is not a number.
    #[error("Revision version is not a number.")]
    FormRevisionVersionUnparsable,

    /// Unable to parse the value after `ThreadingModel` key.
    #[error("Unable to parse the value after ThreadingModel key")]
    FormThreadingModelUnparsable,

    /// `ThreadingModel` can only be 0 or 1.
    #[error("ThreadingModel can only be 0 (Apartment Threaded), or 1 (Single Threaded)")]
    FormThreadingModelInvalid,

    /// No property name found after `BeginProperty` keyword.
    #[error("No property name found after BeginProperty keyword.")]
    FormNoPropertyName,

    /// Unable to parse the `RelatedDoc` property line.
    #[error("Unable to parse the RelatedDoc property line.")]
    FormRelatedDocLineUnparsable,

    /// `AutoIncrement` can only be 0 or -1.
    #[error("AutoIncrement can only be a 0 (false) or a -1 (true)")]
    FormAutoIncrementUnparsable,

    /// `CompatibilityMode` value is invalid.
    #[error("CompatibilityMode can only be a 0 (CompatibilityMode::NoCompatibility), 1 (CompatibilityMode::Project), or 2 (CompatibilityMode::CompatibleExe)")]
    FormCompatibilityModeUnparsable,

    /// `NoControlUpgrade` value is invalid.
    #[error("NoControlUpgrade can only be a 0 (UpgradeControls::Upgrade) or a 1 (UpgradeControls::NoUpgrade)")]
    FormNoControlUpgradeUnparsable,

    /// `ServerSupportFiles` can only be 0 or -1.
    #[error("ServerSupportFiles can only be a 0 (false) or a -1 (true)")]
    FormServerSupportFilesUnparsable,

    /// `Comment` line was unparsable.
    #[error("Comment line was unparsable")]
    FormCommentUnparsable,

    /// `PropertyPage` line was unparsable.
    #[error("PropertyPage line was unparsable")]
    FormPropertyPageUnparsable,

    /// `CompilationType` can only be 0 or -1.
    #[error("CompilationType can only be a 0 (false) or a -1 (true)")]
    FormCompilationTypeUnparsable,

    /// `OptimizationType` value is invalid.
    #[error("OptimizationType can only be a 0 (FastCode) or 1 (SmallCode), or 2 (NoOptimization)")]
    FormOptimizationTypeUnparsable,

    /// `FavorPentiumPro(tm)` can only be 0 or -1.
    #[error("FavorPentiumPro(tm) can only be a 0 (false) or a -1 (true)")]
    FormFavorPentiumProUnparsable,

    /// `Designer` line is unparsable.
    #[error("Designer line is unparsable")]
    FormDesignerLineUnparsable,

    /// Form line is unparsable.
    #[error("Form line is unparsable")]
    FormFormLineUnparsable,

    /// `UserControl` line is unparsable.
    #[error("UserControl line is unparsable")]
    FormUserControlLineUnparsable,

    /// `UserDocument` line is unparsable.
    #[error("UserDocument line is unparsable")]
    FormUserDocumentLineUnparsable,

    /// Period expected in version number.
    #[error("Period expected in version number")]
    FormPeriodExpectedInVersionNumber,

    /// `CodeViewDebugInfo` can only be 0 or -1.
    #[error("CodeViewDebugInfo can only be a 0 (false) or a -1 (true)")]
    FormCodeViewDebugInfoUnparsable,

    /// `NoAliasing` can only be 0 or -1.
    #[error("NoAliasing can only be a 0 (false) or a -1 (true)")]
    FormNoAliasingUnparsable,

    /// `RemoveUnusedControlInfo` value is invalid.
    #[error("RemoveUnusedControlInfo can only be 0 (UnusedControlInfo::Retain) or -1 (UnusedControlInfo::Remove)")]
    FormUnusedControlInfoUnparsable,

    /// `BoundsCheck` can only be 0 or -1.
    #[error("BoundsCheck can only be a 0 (false) or a -1 (true)")]
    FormBoundsCheckUnparsable,

    /// `OverflowCheck` can only be 0 or -1.
    #[error("OverflowCheck can only be a 0 (false) or a -1 (true)")]
    FormOverflowCheckUnparsable,

    /// `FlPointCheck` can only be 0 or -1.
    #[error("FlPointCheck can only be a 0 (false) or a -1 (true)")]
    FormFlPointCheckUnparsable,

    /// `FDIVCheck` value is invalid.
    #[error("FDIVCheck can only be a 0 (PentiumFDivBugCheck::CheckPentiumFDivBug) or a -1 (PentiumFDivBugCheck::NoPentiumFDivBugCheck)")]
    FormFDIVCheckUnparsable,

    /// `UnroundedFP` value is invalid.
    #[error("UnroundedFP can only be a 0 (UnroundedFloatingPoint::DoNotAllow) or a -1 (UnroundedFloatingPoint::Allow)")]
    FormUnroundedFPUnparsable,

    /// `StartMode` value is invalid.
    #[error("StartMode can only be a 0 (StartMode::StandAlone) or a 1 (StartMode::Automation)")]
    FormStartModeUnparsable,

    /// `Unattended` value is invalid.
    #[error("Unattended can only be a 0 (Unattended::False) or a -1 (Unattended::True)")]
    FormUnattendedUnparsable,

    /// `Retained` value is invalid.
    #[error(
        "Retained can only be a 0 (Retained::UnloadOnExit) or a 1 (Retained::RetainedInMemory)"
    )]
    FormRetainedUnparsable,

    /// Unable to parse the `ShortCut` property.
    #[error("Unable to parse the ShortCut property.")]
    FormShortCutUnparsable,

    /// `DebugStartup` can only be 0 or -1.
    #[error("DebugStartup can only be a 0 (false) or a -1 (true)")]
    FormDebugStartupOptionUnparsable,

    /// `UseExistingBrowser` value is invalid.
    #[error("UseExistingBrowser can only be a 0 (UseExistingBrowser::DoNotUse) or a -1 (UseExistingBrowser::Use)")]
    FormUseExistingBrowserUnparsable,

    /// `AutoRefresh` can only be 0 or -1.
    #[error("AutoRefresh can only be a 0 (false) or a -1 (true)")]
    FormAutoRefreshUnparsable,

    /// `Thread Per Object` is not a number.
    #[error("Thread Per Object is not a number.")]
    FormThreadPerObjectUnparsable,

    /// Unknown attribute in class header file.
    #[error("Unknown attribute in class header file. Must be one of: VB_Name, VB_GlobalNameSpace, VB_Creatable, VB_PredeclaredId, VB_Exposed, VB_Description, VB_Ext_KEY")]
    FormUnknownAttribute,

    /// Error parsing header.
    #[error("Error parsing header")]
    FormHeader,

    /// No name in the attribute section of the VB6 file.
    #[error("No name in the attribute section of the VB6 file")]
    FormMissingNameAttribute,

    /// Keyword not found.
    #[error("Keyword not found")]
    FormKeywordNotFound,

    /// Error parsing true/false from header.
    #[error("Error parsing true/false from header. Must be a 0 (false), -1 (true), or 1 (true)")]
    FormTrueFalseOneZeroNegOneUnparsable,

    /// Error parsing the VB6 file contents.
    #[error("Error parsing the VB6 file contents")]
    FormFileContent,

    /// Max Threads is not a number.
    #[error("Max Threads is not a number.")]
    FormMaxThreadsUnparsable,

    /// No `EndProperty` found after `BeginProperty`.
    #[error("No EndProperty found after BeginProperty")]
    FormNoEndProperty,

    /// No line ending after `EndProperty`.
    #[error("No line ending after EndProperty")]
    FormNoLineEndingAfterEndProperty,

    /// Expected namespace after `Begin` keyword.
    #[error("Expected namespace after Begin keyword")]
    FormNoNamespaceAfterBegin,

    /// No dot found after namespace.
    #[error("No dot found after namespace")]
    FormNoDotAfterNamespace,

    /// No User Control name found after namespace and '.'.
    #[error("No User Control name found after namespace and '.'")]
    FormNoUserControlNameAfterDot,

    /// No space after Control kind.
    #[error("No space after Control kind")]
    FormNoSpaceAfterControlKind,

    /// No control name found after Control kind.
    #[error("No control name found after Control kind")]
    FormNoControlNameAfterControlKind,

    /// No line ending after Control name.
    #[error("No line ending after Control name")]
    FormNoLineEndingAfterControlName,

    /// Unknown token in form parsing.
    #[error("Unknown token")]
    FormUnknownToken,

    /// Title text was unparsable.
    #[error("Title text was unparsable")]
    FormTitleUnparsable,

    /// Unable to parse hex color value.
    #[error("Unable to parse hex color value")]
    FormHexColorParseError,

    /// Unknown control in control list.
    #[error("Unknown control in control list")]
    FormUnknownControlKind,

    /// Property name is not a valid ASCII string.
    #[error("Property name is not a valid ASCII string")]
    FormPropertyNameAsciiConversionError,

    /// String is unterminated.
    #[error("String is unterminated")]
    FormUnterminatedString,

    /// Unable to parse VB6 string.
    #[error("Unable to parse VB6 string.")]
    FormStringParseError,

    /// Property value is not a valid ASCII string.
    #[error("Property value is not a valid ASCII string")]
    FormPropertyValueAsciiConversionError,

    /// Key value pair format is incorrect.
    #[error("Key value pair format is incorrect")]
    FormKeyValueParseError,

    /// Namespace is not a valid ASCII string.
    #[error("Namespace is not a valid ASCII string")]
    FormNamespaceAsciiConversionError,

    /// Control kind is not a valid ASCII string.
    #[error("Control kind is not a valid ASCII string")]
    FormControlKindAsciiConversionError,

    /// Qualified control name is not a valid ASCII string.
    #[error("Qualified control name is not a valid ASCII string")]
    FormQualifiedControlNameAsciiConversionError,

    /// Variable names must be less than 255 characters in VB6.
    #[error("Variable names must be less than 255 characters in VB6.")]
    FormVariableNameTooLong,

    /// Invalid top-level control type.
    #[error("Invalid top-level control type: '{control_type}'. Form files must have either 'VB.Form' or 'VB.MDIForm' as the top-level element.")]
    FormInvalidTopLevelControl {
        /// The invalid control type that was found.
        control_type: String,
    },

    /// Internal Parser Error.
    #[error("Internal Parser Error - please report this issue to the developers.")]
    FormInternalParseError,

    // ========================================================================
    // PROJECT FILE ERRORS
    // ========================================================================
    /// A section header was not terminated properly.
    #[error("A section header was expected but was not terminated with a ']' character.")]
    ProjectUnterminatedSectionHeader,

    /// Property name was not found in a property line.
    #[error("Project property line invalid. Expected a Property Name followed by an equal sign '=' and a Property Value.")]
    ProjectPropertyNameNotFound,

    /// Project type is unknown.
    #[error("'Type' property line invalid. Only the values 'Exe', 'OleDll', 'Control', or 'OleExe' are valid.")]
    ProjectTypeUnknown,

    /// Designer file not found.
    #[error("'Designer' line is invalid. Expected a designer path after the equal sign '='. Found a newline or the end of the file instead.")]
    ProjectDesignerFileNotFound,

    /// Reference compiled UUID missing matching brace.
    #[error("'Reference' line is invalid. The line started with '*\\G' indicating a compiled reference. Expected a closing '}}' after the UUID, but found a newline or the end of the file instead.")]
    ProjectReferenceCompiledUuidMissingMatchingBrace,

    /// Reference compiled UUID is invalid.
    #[error("'Reference' line is invalid. The line started with '*\\G' indicating a compiled reference but the contents of the '{{' and '}}' was not a valid UUID.")]
    ProjectReferenceCompiledUuidInvalid,

    /// Reference project path not found.
    #[error("'Reference' line is invalid. Expected a reference path but found a newline or the end of the file instead.")]
    ProjectReferenceProjectPathNotFound,

    /// Reference project path is invalid.
    #[error("'Reference' line is invalid. Expected a reference path to begin with '*\\A' followed by the path to the reference project file ending with a quote '\"' character. Found '{value}' instead.")]
    ProjectReferenceProjectPathInvalid {
        /// The invalid path value that was found.
        value: String,
    },

    /// Reference compiled unknown1 missing.
    #[error("'Reference' line is invalid. Expected a compiled reference 'unknown1' value after the UUID, between '#' characters but found a newline or the end of the file instead.")]
    ProjectReferenceCompiledUnknown1Missing,

    /// Reference compiled unknown2 missing.
    #[error("'Reference' line is invalid. Expected a compiled reference 'unknown2' value after the UUID and 'unknown1', between '#' characters but found a newline or the end of the file instead.")]
    ProjectReferenceCompiledUnknown2Missing,

    /// Reference compiled path not found.
    #[error("'Reference' line is invalid. Expected a compiled reference 'path' value after the UUID, 'unknown1', and 'unknown2', between '#' characters but found a newline or the end of the file instead.")]
    ProjectReferenceCompiledPathNotFound,

    /// Reference compiled description not found.
    #[error("'Reference' line is invalid. Expected a compiled reference 'description' value after the UUID, 'unknown1', 'unknown2', and 'path', but found a newline or the end of the file instead.")]
    ProjectReferenceCompiledDescriptionNotFound,

    /// Reference compiled description is invalid.
    #[error("'Reference' line is invalid. Compiled reference description contains a '#' character, which is not allowed. The description must be a valid ASCII string without any '#' characters.")]
    ProjectReferenceCompiledDescriptionInvalid,

    /// Object project path not found.
    #[error("'Object' line is invalid. Project based objects lines must be quoted strings and begin with '*\\A' followed by the path to the object project file ending with a quote '\"' character. Found a newline or the end of the file instead.")]
    ProjectObjectProjectPathNotFound,

    /// Object compiled missing opening brace.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{'. Found a newline or the end of the file instead.")]
    ProjectObjectCompiledMissingOpeningBrace,

    /// Object compiled UUID missing matching brace.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID. Expected a closing '}}' after the UUID, but found a newline or the end of the file instead.")]
    ProjectObjectCompiledUuidMissingMatchingBrace,

    /// Object compiled UUID is invalid.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID, and end with '}}'. The UUID was not valid. Expected a valid UUID in the format 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' containing only ASCII characters.")]
    ProjectObjectCompiledUuidInvalid,

    /// Object compiled version missing.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by a '#' character, and then a version number. Expected a '#' character, but found a newline or the end of the file instead.")]
    ProjectObjectCompiledVersionMissing,

    /// Object compiled version is invalid.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by a '#' character, and then a version number. The version number was not valid. Expected a valid version number in the format 'x.x'. The version number must contain only '.' or the characters \"0\"..\"9\". Invalid character found instead.")]
    ProjectObjectCompiledVersionInvalid,

    /// Object compiled unknown1 missing.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by '#', a version number, followed by another '#', then an 'unknown1' value followed by \"; \". Expected \"; \", but found a newline or the end of the file instead.")]
    ProjectObjectCompiledUnknown1Missing,

    /// Object compiled file name not found.
    #[error("'Object' line is invalid. Compiled object lines must begin with '{{', enclose a valid UUID with '}}', followed by '#', a version number, followed by another '#', then an 'unknown1' value followed by \"; \", and ending with the object's file name. Expected the object's file name, but found a newline or the end of the file instead.")]
    ProjectObjectCompiledFileNameNotFound,

    /// Module name not found.
    #[error("'Module' line is invalid. Expected a module name followed by a \"; \". Found a newline or the end of the file instead.")]
    ProjectModuleNameNotFound,

    /// Module file name not found.
    #[error("'Module' line is invalid. Expected a module name followed by a \"; \", followed by the module file name. Found a newline or the end of the file instead.")]
    ProjectModuleFileNameNotFound,

    /// Class name not found.
    #[error("'Class' line is invalid. Expected a class name followed by a \"; \". Found a newline or the end of the file instead.")]
    ProjectClassNameNotFound,

    /// Class file name not found.
    #[error("'Class' line is invalid. Expected a class name followed by a \"; \", followed by the class file name. Found a newline or the end of the file instead.")]
    ProjectClassFileNameNotFound,

    /// Path value not found.
    #[error("'{parameter_line_name}' line is invalid. Expected a '{parameter_line_name}' path after the equal sign '='. Found a newline or the end of the file instead.")]
    ProjectPathValueNotFound {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value not found.
    #[error("'{parameter_line_name}' line is invalid. Expected a quoted '{parameter_line_name}' value after the equal sign '='. Found a newline or the end of the file instead.")]
    ProjectParameterValueNotFound {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value missing opening quote.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing an opening quote. Expected a closing quote for the '{parameter_line_name}' value.")]
    ProjectParameterValueMissingOpeningQuote {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value missing matching quote.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing a matching quote. Expected a closing quote for the '{parameter_line_name}' value.")]
    ProjectParameterValueMissingMatchingQuote {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value missing quotes.
    #[error("'{parameter_line_name}' line is invalid. '{parameter_line_name}' is missing both opening and closing quotes. Expected a quoted '{parameter_line_name}' value after the equal sign '='.")]
    ProjectParameterValueMissingQuotes {
        /// The parameter line name.
        parameter_line_name: String,
    },

    /// Parameter value is invalid.
    #[error("'{parameter_line_name}' line is invalid. '{invalid_value}' is not a valid value for '{parameter_line_name}'. Only {valid_value_message} are valid values for '{parameter_line_name}'.")]
    ProjectParameterValueInvalid {
        /// The parameter line name.
        parameter_line_name: String,
        /// The invalid value that was found.
        invalid_value: String,
        /// Valid value message.
        valid_value_message: String,
    },

    /// `DllBaseAddress` not found.
    #[error("'DllBaseAddress' line is invalid. Expected a hex address after the equal sign '='. Found a newline or the end of the file instead.")]
    ProjectDllBaseAddressNotFound,

    /// `DllBaseAddress` missing hex prefix.
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address beginning with '&h' after the equal sign '='.")]
    ProjectDllBaseAddressMissingHexPrefix,

    /// `DllBaseAddress` unparsable.
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address after the equal sign '=' beginning with '&h'. Unable to parse hex value '{hex_value}'.")]
    ProjectDllBaseAddressUnparsable {
        /// The hex value that couldn't be parsed.
        hex_value: String,
    },

    /// `DllBaseAddress` unparsable (empty).
    #[error("'DllBaseAddress' line is invalid. Expected a valid hex address after the equal sign '=' beginning with '&h'. Unable to parse empty hex value.")]
    ProjectDllBaseAddressUnparsableEmpty,

    /// Parameter line is unknown.
    #[error("'{parameter_line_name}' line is unknown.")]
    ProjectParameterLineUnknown {
        /// The parameter line name.
        parameter_line_name: String,
    },

    // ========================================================================
    // RESOURCE FILE ERRORS
    // ========================================================================
    /// I/O error while reading the resource file.
    #[error("Failed to read resource file: {message}")]
    ResourceIoError {
        /// Error message from I/O error.
        message: String,
    },

    /// Requested offset is beyond the end of the file.
    #[error("Offset {offset} is out of bounds for file of length {file_length}")]
    ResourceOffsetOutOfBounds {
        /// The offset that is out of bounds.
        offset: usize,
        /// The length of the file.
        file_length: usize,
    },

    /// Invalid or corrupted data at the specified offset.
    #[error("Invalid data at offset {offset}: {details}")]
    ResourceInvalidData {
        /// The offset where the invalid data was found.
        offset: usize,
        /// Details about the invalid data.
        details: String,
    },

    /// Failed to read header bytes at the specified offset.
    #[error("Failed to read header at offset {offset}: {reason}")]
    ResourceHeaderReadError {
        /// The offset where the read error occurred.
        offset: usize,
        /// The reason for the read error.
        reason: String,
    },

    /// Record size fields don't match expected values.
    #[error("Record size mismatch at offset {offset}: expected {expected}, got {actual}")]
    ResourceSizeMismatch {
        /// The offset where the size mismatch occurred.
        offset: usize,
        /// The expected size.
        expected: usize,
        /// The actual size found.
        actual: usize,
    },

    /// Buffer slice conversion failed.
    #[error("Failed to convert buffer slice at offset {offset} to fixed-size array")]
    ResourceBufferConversionError {
        /// The offset where the conversion error occurred.
        offset: usize,
    },

    /// Detected corruption in list items structure.
    #[error("Corrupted list items at offset {offset}: {details}")]
    ResourceCorruptedListItems {
        /// The offset where the corruption was detected.
        offset: usize,
        /// Details about the corruption.
        details: String,
    },

    // ========================================================================
    // SOURCE FILE DECODING ERRORS
    // ========================================================================
    /// The source file is malformed.
    #[error("Unable to parse source file: {message}")]
    SourceFileMalformed {
        /// The error message describing the issue.
        message: String,
    },
}

/// Represents the severity level of a parsing diagnostic.
///
/// This enum is used to distinguish between different types of issues
/// encountered during parsing, from informational notes to fatal errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub enum Severity {
    /// Informational message, not a problem.
    Note,
    /// Potential issue that should be addressed but doesn't prevent usage.
    Warning,
    /// Fatal error that prevents successful parsing or usage.
    #[default]
    Error,
}

impl Display for Severity {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Severity::Note => write!(f, "note"),
            Severity::Warning => write!(f, "warning"),
            Severity::Error => write!(f, "error"),
        }
    }
}

/// Represents a span of source code, typically associated with an error or diagnostic.
///
/// A span identifies a region in the source code by offset, line numbers, and length.
/// This is used to highlight the exact location of errors in diagnostic messages.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Span {
    /// The byte offset into the source content where this span starts.
    pub offset: u32,
    /// The starting line number (1-based).
    pub line_start: u32,
    /// The ending line number (1-based).
    pub line_end: u32,
    /// The length of this span in bytes.
    pub length: u32,
}

impl Span {
    /// Creates a new span.
    #[must_use]
    pub fn new(offset: u32, line_start: u32, line_end: u32, length: u32) -> Self {
        Self {
            offset,
            line_start,
            line_end,
            length,
        }
    }

    /// Creates a zero-length span at offset 0.
    #[must_use]
    pub fn zero() -> Self {
        Self {
            offset: 0,
            line_start: 0,
            line_end: 0,
            length: 0,
        }
    }

    /// Creates a span of length 1 at the given offset and line.
    #[must_use]
    pub fn at(offset: u32, line: u32) -> Self {
        Self {
            offset,
            line_start: line,
            line_end: line,
            length: 1,
        }
    }
}

/// Represents a labeled span in a multi-span diagnostic.
///
/// Labels are used to annotate multiple locations in the source code
/// within a single error message, providing context for complex errors.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagnosticLabel {
    /// The span this label refers to.
    pub span: Span,
    /// The message to display for this label.
    pub message: String,
}

impl DiagnosticLabel {
    /// Creates a new label.
    pub fn new(span: Span, message: impl Into<String>) -> Self {
        Self {
            span,
            message: message.into(),
        }
    }
}

/// Contains detailed information about an error that occurred during parsing.
///
/// This struct contains the source name, source content, error offset,
/// line start and end positions, and the kind of error. All errors now use
/// the unified [`ErrorKind`] type.
///
/// Example usage:
/// ```rust
/// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity};
///
/// let error_details = ErrorDetails {
///     source_name: "example.cls".to_string().into_boxed_str(),
///     source_content: "Some VB6 code here...",
///     error_offset: 10,
///     line_start: 1,
///     line_end: 1,
///     kind: ErrorKind::UnknownToken { token: "???".to_string() },
///     severity: Severity::Error,
///     labels: vec![],
///     notes: vec![],
/// };
/// error_details.print();
/// ```
#[derive(Debug, Clone)]
pub struct ErrorDetails<'a> {
    /// The name of the source file where the error occurred.
    pub source_name: Box<str>,
    /// The content of the source file where the error occurred.
    pub source_content: &'a str,
    /// The offset in the source content where the error occurred.
    ///
    /// Note: This is a u32 to reflect VB6's 32-bit addressing limitations.
    pub error_offset: u32,
    /// The starting line number of the error.
    ///
    /// Note: This is a u32 to reflect VB6's 32-bit addressing limitations.
    pub line_start: u32,
    /// The ending line number of the error.
    ///
    /// Note: This is a u32 to reflect VB6's 32-bit addressing limitations.
    pub line_end: u32,
    /// The kind of error that occurred.
    pub kind: ErrorKind,
    /// The severity of this diagnostic (Error, Warning, or Note).
    pub severity: Severity,
    /// Additional labeled spans for multi-span diagnostics.
    /// This allows annotating multiple locations in the source code
    /// within a single error message.
    pub labels: Vec<DiagnosticLabel>,
    /// Additional notes to provide context for this diagnostic.
    /// These are displayed after the main error message.
    pub notes: Vec<String>,
}

impl<'a> ErrorDetails<'a> {
    /// Creates a basic `ErrorDetails` with no labels or notes.
    ///
    /// This is a convenience constructor for the common case where
    /// only the basic error information is needed.
    #[must_use]
    pub fn basic(
        source_name: Box<str>,
        source_content: &'a str,
        error_offset: u32,
        line_start: u32,
        line_end: u32,
        kind: ErrorKind,
        severity: Severity,
    ) -> ErrorDetails<'a> {
        ErrorDetails {
            source_name,
            source_content,
            error_offset,
            line_start,
            line_end,
            kind,
            severity,
            labels: Vec::new(),
            notes: Vec::new(),
        }
    }

    /// Adds a labeled span to this error.
    #[must_use]
    pub fn with_label(mut self, label: DiagnosticLabel) -> Self {
        self.labels.push(label);
        self
    }

    /// Adds multiple labeled spans to this error.
    #[must_use]
    pub fn with_labels(mut self, labels: Vec<DiagnosticLabel>) -> Self {
        self.labels.extend(labels);
        self
    }

    /// Adds a note to this error.
    #[must_use]
    pub fn with_note(mut self, note: impl Into<String>) -> Self {
        self.notes.push(note.into());
        self
    }

    /// Adds multiple notes to this error.
    #[must_use]
    pub fn with_notes(mut self, notes: Vec<String>) -> Self {
        self.notes.extend(notes);
        self
    }
}

impl Display for ErrorDetails<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ErrorDetails {{ source_name: {}, error_offset: {}, line_start: {}, line_end: {}, kind: {:?} }}",
            self.source_name,
            self.error_offset,
            self.line_start,
            self.line_end,
            self.kind,
        )
    }
}

impl ErrorDetails<'_> {
    /// Print the `ErrorDetails` using ariadne for formatting.
    ///
    /// Example usage:
    /// ```rust
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity};
    ///
    /// let error_details = ErrorDetails {
    ///     source_name: "example.cls".to_string().into_boxed_str(),
    ///     source_content: "Some VB6 code here...",
    ///     error_offset: 10,
    ///     line_start: 1,
    ///     line_end: 1,
    ///     kind: ErrorKind::UnknownToken { token: "???".to_string() },
    ///     severity: Severity::Error,
    ///     labels: vec![],
    ///     notes: vec![],
    /// };
    /// error_details.print();
    /// ```
    pub fn print(&self) {
        let cache = (
            self.source_name.to_string(),
            Source::from(self.source_content),
        );

        let mut report = Report::build(
            ReportKind::Error,
            (
                self.source_name.to_string(),
                self.line_start as usize..=self.line_end as usize,
            ),
        )
        .with_message(self.kind.to_string())
        .with_label(
            Label::new((
                self.source_name.to_string(),
                self.error_offset as usize..=self.error_offset as usize,
            ))
            .with_message("error here"),
        );

        // Add additional labeled spans
        for label in &self.labels {
            report = report.with_label(
                Label::new((
                    self.source_name.to_string(),
                    label.span.offset as usize
                        ..=(label.span.offset + label.span.length.max(1) - 1) as usize,
                ))
                .with_message(&label.message),
            );
        }

        // Add notes
        for note in &self.notes {
            report = report.with_note(note);
        }

        let result = report.finish().print(cache);

        if let Some(e) = result.err() {
            eprint!("Error attempting to build ErrorDetails print message {e:?}");
        }
    }

    /// Eprint the `ErrorDetails` using ariadne for formatting.
    ///
    /// Example usage:
    /// ```rust
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity};
    ///
    /// let error_details = ErrorDetails {
    ///     source_name: "example.cls".to_string().into_boxed_str(),
    ///     source_content: "Some VB6 code here...",
    ///     error_offset: 10,
    ///     line_start: 1,
    ///     line_end: 1,
    ///     kind: ErrorKind::UnknownToken {
    ///         token: "???".to_string(),
    ///     },
    ///     severity: Severity::Error,
    ///     labels: vec![],
    ///     notes: vec![],
    /// };
    /// error_details.eprint();
    /// ```
    pub fn eprint(&self) {
        let cache = (
            self.source_name.to_string(),
            Source::from(self.source_content),
        );

        let mut report = Report::build(
            ReportKind::Error,
            (
                self.source_name.to_string(),
                self.line_start as usize..=self.line_end as usize,
            ),
        )
        .with_message(format!("{:?}", self.kind))
        .with_label(
            Label::new((
                self.source_name.to_string(),
                self.error_offset as usize..=self.error_offset as usize,
            ))
            .with_message("error here"),
        );

        // Add additional labeled spans
        for label in &self.labels {
            report = report.with_label(
                Label::new((
                    self.source_name.to_string(),
                    label.span.offset as usize
                        ..=(label.span.offset + label.span.length.max(1) - 1) as usize,
                ))
                .with_message(&label.message),
            );
        }

        // Add notes
        for note in &self.notes {
            report = report.with_note(note);
        }

        let result = report.finish().eprint(cache);

        if let Some(e) = result.err() {
            eprint!("Error attempting to build ErrorDetails eprint message {e:?}");
        }
    }

    /// Convert the `ErrorDetails` into a string using ariadne for formatting
    ///
    /// # Errors
    /// This function will return an error if there is an issue converting the
    /// formatted report into a UTF-8 string.
    pub fn print_to_string(&self) -> Result<String, Box<dyn Error>> {
        let cache = (
            self.source_name.to_string(),
            Source::from(self.source_content),
        );

        let mut buf = Vec::new();

        let _ = Report::build(
            ReportKind::Error,
            (
                self.source_name.to_string(),
                self.line_start as usize..=self.line_end as usize,
            ),
        )
        .with_message(self.kind.to_string())
        .with_label(
            Label::new((
                self.source_name.to_string(),
                self.error_offset as usize..=self.error_offset as usize,
            ))
            .with_message("error here"),
        )
        .finish()
        .write(cache, &mut buf);

        let text = String::from_utf8(buf.clone())?;

        Ok(text)
    }
}
