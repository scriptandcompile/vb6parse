//! Defines compilation settings enums and structs for VB6 projects.
//!
//! Includes settings for native code compilation and P-Code.
//! Each setting is represented as an enum with variants corresponding to possible values.
//! Provides methods to update individual settings while maintaining immutability.
//!
use std::str::FromStr;

use num_enum::TryFromPrimitive;
use serde::Serialize;
use strum_macros::{EnumIter, EnumMessage};

/// Represents whether unrounded floating point is allowed.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum UnroundedFloatingPoint {
    /// Do not use unrounded floating point.
    #[default]
    #[strum(message = "Do not use unrounded floating point")]
    DoNotAllow = 0,
    /// Use unrounded floating point.
    #[strum(message = "Use unrounded floating point")]
    Allow = -1,
}

impl TryFrom<&str> for UnroundedFloatingPoint {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(UnroundedFloatingPoint::DoNotAllow),
            "-1" => Ok(UnroundedFloatingPoint::Allow),
            _ => Err(format!("Unknown UnroundedFloatingPoint value: '{value}'")),
        }
    }
}

/// Represents whether to check for the Pentium FDIV bug.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum PentiumFDivBugCheck {
    /// Check for the Pentium FDIV bug.
    #[strum(message = "Check for Pentium FDIV bug")]
    CheckPentiumFDivBug = 0,
    /// Do not check for the Pentium FDIV bug.
    #[default]
    #[strum(message = "Ignore Pentium FDIV bug")]
    NoPentiumFDivBugCheck = -1,
}

impl TryFrom<&str> for PentiumFDivBugCheck {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(PentiumFDivBugCheck::CheckPentiumFDivBug),
            "-1" => Ok(PentiumFDivBugCheck::NoPentiumFDivBugCheck),
            _ => Err(format!("Unknown PentiumFDivBugCheck value: '{value}'")),
        }
    }
}

/// Represents whether to perform bounds checking.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum BoundsCheck {
    /// Perform bounds checking.
    #[default]
    #[strum(message = "Perform bounds checking")]
    CheckBounds = 0,
    /// Do not perform bounds checking.
    #[strum(message = "Do not perform bounds checking")]
    NoBoundsCheck = -1,
}

impl TryFrom<&str> for BoundsCheck {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(BoundsCheck::CheckBounds),
            "-1" => Ok(BoundsCheck::NoBoundsCheck),
            _ => Err(format!("Unknown BoundsCheck value: '{value}'")),
        }
    }
}

/// Represents whether to perform overflow checking.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum OverflowCheck {
    /// Perform overflow checking.
    #[default]
    #[strum(message = "Check for overflow")]
    CheckOverflow = 0,
    /// Do not perform overflow checking.
    #[strum(message = "Do not check for overflow")]
    NoOverflowCheck = -1,
}

impl TryFrom<&str> for OverflowCheck {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(OverflowCheck::CheckOverflow),
            "-1" => Ok(OverflowCheck::NoOverflowCheck),
            _ => Err(format!("Unknown OverflowCheck value: '{value}'")),
        }
    }
}

/// Represents whether to check for floating point errors.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum FloatingPointErrorCheck {
    /// Perform floating point error checking.
    #[default]
    #[strum(message = "Check for floating point errors")]
    CheckFloatingPointError = 0,
    /// Do not perform floating point error checking.
    #[strum(message = "Do not check for floating point errors")]
    NoFloatingPointErrorCheck = -1,
}

impl TryFrom<&str> for FloatingPointErrorCheck {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(FloatingPointErrorCheck::CheckFloatingPointError),
            "-1" => Ok(FloatingPointErrorCheck::NoFloatingPointErrorCheck),
            _ => Err(format!("Unknown FloatingPointErrorCheck value: '{value}'")),
        }
    }
}

/// Represents whether to create CodeView debug information.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum CodeViewDebugInfo {
    /// Do not create CodeView debug information.
    #[default]
    #[strum(message = "Do not create CodeView debug info")]
    NotCreated = 0,
    /// Create CodeView debug information.
    #[strum(message = "Create CodeView debug info")]
    Created = -1,
}

impl TryFrom<&str> for CodeViewDebugInfo {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(CodeViewDebugInfo::NotCreated),
            "-1" => Ok(CodeViewDebugInfo::Created),
            _ => Err(format!("Unknown CodeViewDebugInfo value: '{value}'")),
        }
    }
}

/// Represents whether to favor Pentium Pro optimizations.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum FavorPentiumPro {
    /// Do not favor Pentium Pro optimizations.
    #[default]
    #[strum(message = "Do not favor Pentium Pro optimizations")]
    False = 0,
    /// Favor Pentium Pro optimizations.
    #[strum(message = "Favor Pentium Pro optimizations")]
    True = -1,
}

impl TryFrom<&str> for FavorPentiumPro {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(FavorPentiumPro::False),
            "-1" => Ok(FavorPentiumPro::True),
            _ => Err(format!("Unknown FavorPentiumPro value: '{value}'")),
        }
    }
}

/// Represents whether to assume aliasing.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum Aliasing {
    /// Assume aliasing.
    #[default]
    #[strum(message = "Assume aliasing")]
    AssumeAliasing = 0,
    /// Do not assume aliasing.
    #[strum(message = "Do not assume aliasing")]
    AssumeNoAliasing = -1,
}

impl TryFrom<&str> for Aliasing {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(Aliasing::AssumeAliasing),
            "-1" => Ok(Aliasing::AssumeNoAliasing),
            _ => Err(format!("Unknown Aliasing value: '{value}'")),
        }
    }
}

/// Represents the optimization type for native code compilation.
#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum OptimizationType {
    /// Favor fast code optimizations.
    #[default]
    #[strum(message = "Favor fast code")]
    FavorFastCode = 0,
    /// Favor small code optimizations.
    #[strum(message = "Favor small code")]
    FavorSmallCode = 1,
    /// Do not optimize.
    #[strum(message = "Do not optimize")]
    NoOptimization = 2,
}

impl TryFrom<&str> for OptimizationType {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(OptimizationType::FavorFastCode),
            "1" => Ok(OptimizationType::FavorSmallCode),
            "2" => Ok(OptimizationType::NoOptimization),
            _ => Err(format!("Unknown OptimizationType value: '{value}'")),
        }
    }
}

/// Settings specific to native code compilation.
#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, Default)]
pub struct NativeCodeSettings {
    /// Optimization type setting.
    pub optimization_type: OptimizationType,
    /// Whether to favor Pentium Pro optimizations.
    pub favor_pentium_pro: FavorPentiumPro,
    /// Whether to create CodeView debug information.
    pub code_view_debug_info: CodeViewDebugInfo,
    /// Whether to assume aliasing.
    pub aliasing: Aliasing,
    /// Whether to perform bounds checking.
    pub bounds_check: BoundsCheck,
    /// Whether to perform overflow checking.
    pub overflow_check: OverflowCheck,
    /// Whether to perform floating point error checking.
    pub floating_point_check: FloatingPointErrorCheck,
    /// Whether to check for the Pentium FDIV bug.
    pub pentium_fdiv_bug_check: PentiumFDivBugCheck,
    /// Whether to use unrounded floating point.
    pub unrounded_floating_point: UnroundedFloatingPoint,
}

/// Represents the compilation type and its associated settings.
#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, Default)]
pub enum CompilationType {
    /// Native code compilation with specific settings.
    /// Contains various optimization and checking settings.
    ///
    /// Saved as "0" (False) in project files.
    NativeCode(NativeCodeSettings),
    /// P-Code compilation.
    /// Saved as "-1" (True) in project files.
    #[default]
    PCode,
}

impl CompilationType {
    /// Updates the optimization type setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new optimization type to set.
    ///
    /// # Returns
    ///
    /// * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_optimization_type(&mut self, setting: OptimizationType) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                optimization_type: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.optimization_type = setting;
                CompilationType::NativeCode(value)
            }
        }
    }

    /// Updates the favor Pentium Pro setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new favor Pentium Pro setting to set.
    ///
    /// # Returns
    ///
    /// * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_favor_pentium_pro(&mut self, setting: FavorPentiumPro) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                favor_pentium_pro: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.favor_pentium_pro = setting;
                CompilationType::NativeCode(value)
            }
        }
    }

    /// Updates the CodeView debug info setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new CodeView debug info setting to set.
    ///
    /// # Returns
    ///
    /// * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_code_view_debug_info(&mut self, setting: CodeViewDebugInfo) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                code_view_debug_info: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.code_view_debug_info = setting;
                CompilationType::NativeCode(value)
            }
        }
    }

    /// Updates the aliasing setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new aliasing setting to set.
    ///
    /// # Returns
    ///
    ///  * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_aliasing(&mut self, setting: Aliasing) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                aliasing: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.aliasing = setting;
                CompilationType::NativeCode(value)
            }
        }
    }

    /// Updates the bounds check setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new bounds check setting to set.
    ///
    /// # Returns
    ///
    /// * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_bounds_check(&mut self, setting: BoundsCheck) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                bounds_check: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.bounds_check = setting;
                CompilationType::NativeCode(value)
            }
        }
    }

    /// Updates the overflow check setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new overflow check setting to set.
    ///
    /// # Returns
    ///
    /// * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_overflow_check(&mut self, setting: OverflowCheck) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                overflow_check: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.overflow_check = setting;
                CompilationType::NativeCode(value)
            }
        }
    }

    /// Updates the floating point error check setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new floating point error check setting to set.
    ///
    /// # Returns
    ///
    /// * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_floating_point_check(
        &mut self,
        setting: FloatingPointErrorCheck,
    ) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                floating_point_check: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.floating_point_check = setting;
                CompilationType::NativeCode(value)
            }
        }
    }

    /// Updates the Pentium FDIV bug check setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new Pentium FDIV bug check setting to set.
    ///
    /// # Returns
    ///
    /// * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_pentium_fdiv_bug_check(self, setting: PentiumFDivBugCheck) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                pentium_fdiv_bug_check: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.pentium_fdiv_bug_check = setting;
                CompilationType::NativeCode(value)
            }
        }
    }

    /// Updates the unrounded floating point setting.
    ///
    /// # Arguments
    ///
    /// * `setting` - The new unrounded floating point setting to set.
    ///
    /// # Returns
    ///
    /// * `CompilationType` - A new CompilationType with the updated setting.
    ///
    #[must_use]
    pub fn update_unrounded_floating_point(
        &mut self,
        setting: UnroundedFloatingPoint,
    ) -> CompilationType {
        match self {
            CompilationType::PCode => CompilationType::NativeCode(NativeCodeSettings {
                unrounded_floating_point: setting,
                ..Default::default()
            }),
            CompilationType::NativeCode(mut value) => {
                value.unrounded_floating_point = setting;
                CompilationType::NativeCode(value)
            }
        }
    }
}

impl FromStr for CompilationType {
    type Err = String;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        match value {
            "0" => Ok(CompilationType::NativeCode(Default::default())),
            "-1" => Ok(CompilationType::PCode),
            _ => Err(format!("Unknown CompilationType value: '{value}'")),
        }
    }
}

impl TryFrom<&str> for CompilationType {
    type Error = String;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(CompilationType::NativeCode(Default::default())),
            "-1" => Ok(CompilationType::PCode),
            _ => Err(format!("Unknown CompilationType value: '{value}'")),
        }
    }
}
