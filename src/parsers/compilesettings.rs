use num_enum::TryFromPrimitive;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum UnroundedFloatingPoint {
    #[default]
    DoNotAllow = 0,
    Allow = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum PentiumFDivBugCheck {
    CheckPentiumFDivBug = 0,
    #[default]
    NoPentiumFDivBugCheck = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum BoundsCheck {
    #[default]
    CheckBounds = 0,
    NoBoundsCheck = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum OverflowCheck {
    #[default]
    CheckOverflow = 0,
    NoOverflowCheck = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum FloatingPointErrorCheck {
    #[default]
    CheckFloatingPointError = 0,
    NoFloatingPointErrorCheck = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum CodeViewDebugInfo {
    #[default]
    NotCreated = 0,
    Created = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum FavorPentiumPro {
    #[default]
    False = 0,
    True = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum Aliasing {
    #[default]
    AssumeAliasing = 0,
    AssumeNoAliasing = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i16)]
pub enum OptimizationType {
    #[default]
    FavorFastCode = 0,
    FavorSmallCode = 1,
    NoOptimization = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum CompilationType {
    // 0
    NativeCode {
        optimization_type: OptimizationType,
        favor_pentium_pro: FavorPentiumPro,
        code_view_debug_info: CodeViewDebugInfo,
        aliasing: Aliasing,
        bounds_check: BoundsCheck,
        overflow_check: OverflowCheck,
        floating_point_check: FloatingPointErrorCheck,
        pentium_fdiv_bug_check: PentiumFDivBugCheck,
        unrounded_floating_point: UnroundedFloatingPoint,
    },
    // -1, default
    PCode,
}
