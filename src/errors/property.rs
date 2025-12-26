//! Errors related to VB6 property parsing.
//!
//! This module contains error types for issues that occur during:
//! - Control and form property value validation
//! - Property name parsing
//! - Resource file property references

/// Errors related to property parsing and validation.
#[derive(thiserror::Error, PartialEq, Eq, Debug, Clone)]
pub enum PropertyError {
    /// Indicates that the Appearance property has an invalid value.
    #[error("Appearance can only be a 0 (Flat) or a 1 (ThreeD)")]
    AppearanceInvalid,

    /// Indicates that the `BorderStyle` property has an invalid value.
    #[error("BorderStyle can only be a 0 (None) or 1 (FixedSingle)")]
    BorderStyleInvalid,

    /// Indicates that the `ClipControls` property has an invalid value.
    #[error("ClipControls can only be a 0 (false) or a 1 (true)")]
    ClipControlsInvalid,

    /// Indicates that the `DragMode` property has an invalid value.
    #[error("DragMode can only be 0 (Manual) or 1 (Automatic)")]
    DragModeInvalid,

    /// Indicates that the Enabled property has an invalid value.
    #[error("Enabled can only be 0 (false) or a 1 (true)")]
    EnabledInvalid,

    /// Indicates that the `MousePointer` property has an invalid value.
    #[error("MousePointer can only be 0 (Default), 1 (Arrow), 2 (Cross), 3 (IBeam), 6 (SizeNESW), 7 (SizeNS), 8 (SizeNWSE), 9 (SizeWE), 10 (UpArrow), 11 (Hourglass), 12 (NoDrop), 13 (ArrowHourglass), 14 (ArrowQuestion), 15 (SizeAll), or 99 (Custom)")]
    MousePointerInvalid,

    /// Indicates that the `OLEDropMode` property has an invalid value.
    #[error("OLEDropMode can only be 0 (None), or 1 (Manual)")]
    OLEDropModeInvalid,

    /// Indicates that the `RightToLeft` property has an invalid value.
    #[error("RightToLeft can only be 0 (false) or a 1 (true)")]
    RightToLeftInvalid,

    /// Indicates that the Visible property has an invalid value.
    #[error("Visible can only be 0 (false) or a 1 (true)")]
    VisibleInvalid,

    /// Indicates that the property is unknown in the header file.
    #[error("Unknown property in header file")]
    UnknownProperty,

    /// Indicates that the property value is invalid for any value except 0 or -1.
    #[error("Invalid property value. Only 0 or -1 are valid for this property")]
    InvalidPropertyValueZeroNegOne,

    /// Indicates that the property name could not be parsed.
    #[error("Unable to parse the property name")]
    NameUnparsable,

    /// Indicates that the resource file name could not be parsed.
    #[error("Unable to parse the resource file name")]
    ResourceFileNameUnparsable,

    /// Indicates that the offset into the resource file could not be parsed.
    #[error("Unable to parse the offset into the resource file for property")]
    OffsetUnparsable,

    /// Indicates that the property value is invalid for any value except True or False.
    #[error("Invalid property value. Only True or False are valid for this property")]
    InvalidPropertyValueTrueFalse,
}
