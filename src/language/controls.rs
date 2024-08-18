use crate::language::VB6Color;

use bstr::{BStr, ByteSlice};

/// Represents a VB6 control.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Control<'a> {
    pub common: VB6ControlCommonInformation<'a>,
    pub kind: VB6ControlKind<'a>,
}

/// Represents a VB6 control common information.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ControlCommonInformation<'a> {
    pub name: &'a BStr,
    pub caption: &'a BStr,
    pub back_color: VB6Color,
}

/// Represents a VB6 control kind.
/// A VB6 control kind is an enumeration of the different kinds of
/// standard VB6 controls.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum VB6ControlKind<'a> {
    CommandButton {},
    TextBox {},
    CheckBox {},
    Line {},
    Label {},
    Frame {},
    PictureBox {},
    ComboBox {},
    HScrollBar {},
    VScrollBar {},
    Menu {
        caption: &'a BStr,
        controls: Vec<VB6Control<'a>>,
    },
    Form {
        controls: Vec<VB6Control<'a>>,
    },
}
