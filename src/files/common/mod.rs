//! Common utilities and structures for VB6 file format parsing
//!
//! This module contains shared functionality used across all VB6 file parsers:
//! - Header parsing (VERSION lines, BEGIN blocks)
//! - Attribute statement parsing
//! - Property parsing (generic key-value properties)
//! - Object reference parsing
//!
//! These utilities are used by the class, module, form, and project file parsers.

pub mod header;
pub mod properties;
pub mod property_group_conversions;
pub mod references;

pub use header::*;
pub use properties::*;
// Property group conversions are now implemented using standard TryFrom/From traits
// No exports needed from property_group_conversions module
pub use references::*;
