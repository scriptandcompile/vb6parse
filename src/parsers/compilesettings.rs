use std::str::FromStr;

use num_enum::TryFromPrimitive;
use serde::Serialize;
use strum_macros::{EnumIter, EnumMessage};

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum UnroundedFloatingPoint {
    #[default]
    #[strum(message = "Do not use unrounded floating point")]
    DoNotAllow = 0,
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

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum PentiumFDivBugCheck {
    #[strum(message = "Check for Pentium FDIV bug")]
    CheckPentiumFDivBug = 0,
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

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum BoundsCheck {
    #[default]
    #[strum(message = "Perform bounds checking")]
    CheckBounds = 0,
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

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum OverflowCheck {
    #[default]
    #[strum(message = "Check for overflow")]
    CheckOverflow = 0,
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

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum FloatingPointErrorCheck {
    #[default]
    #[strum(message = "Check for floating point errors")]
    CheckFloatingPointError = 0,
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

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum CodeViewDebugInfo {
    #[default]
    #[strum(message = "Do not create CodeView debug info")]
    NotCreated = 0,
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

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum FavorPentiumPro {
    #[default]
    #[strum(message = "Do not favor Pentium Pro optimizations")]
    False = 0,
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

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum Aliasing {
    #[default]
    #[strum(message = "Assume aliasing")]
    AssumeAliasing = 0,
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

#[derive(
    Debug, PartialEq, Eq, Copy, Clone, Serialize, Default, TryFromPrimitive, EnumIter, EnumMessage,
)]
#[repr(i16)]
pub enum OptimizationType {
    #[default]
    #[strum(message = "Favor fast code")]
    FavorFastCode = 0,
    #[strum(message = "Favor small code")]
    FavorSmallCode = 1,
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

#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, Default)]
pub struct NativeCodeSettings {
    pub optimization_type: OptimizationType,
    pub favor_pentium_pro: FavorPentiumPro,
    pub code_view_debug_info: CodeViewDebugInfo,
    pub aliasing: Aliasing,
    pub bounds_check: BoundsCheck,
    pub overflow_check: OverflowCheck,
    pub floating_point_check: FloatingPointErrorCheck,
    pub pentium_fdiv_bug_check: PentiumFDivBugCheck,
    pub unrounded_floating_point: UnroundedFloatingPoint,
}

#[derive(Debug, PartialEq, Eq, Copy, Clone, Serialize, Default)]
pub enum CompilationType {
    // 0
    NativeCode(NativeCodeSettings),
    // -1
    #[default]
    PCode,
}

impl CompilationType {
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
