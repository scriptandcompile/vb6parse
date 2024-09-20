use bstr::BStr;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// Represents a VB6 menu control.
/// This should only be used as a child of a Form.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6MenuControl<'a> {
    pub name: &'a str,
    pub tag: &'a str,
    pub index: i32,
    pub properties: MenuProperties<'a>,
    pub sub_menus: Vec<VB6MenuControl<'a>>,
}

/// Properties for a Menu control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Menu`](crate::language::controls::VB6ControlKind::Menu).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
/// This is represented within the parsing code independently of
/// [`VB6MenuControl`](crate::language::controls::VB6MenuControl)'s.
///
/// This currently redundant, but is included for the future where the correct
/// behavior of a menu control only being a child of a form is enforced.
///
/// As is, the parser will not enforce this, but the VB6 IDE will.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct MenuProperties<'a> {
    pub caption: &'a BStr,
    pub checked: bool,
    pub enabled: bool,
    pub help_context_id: i32,
    pub negotiate_position: NegotiatePosition,
    pub shortcut: Option<ShortCut>,
    pub visible: bool,
    pub window_list: bool,
}

impl Default for MenuProperties<'_> {
    fn default() -> Self {
        MenuProperties {
            caption: BStr::new(""),
            checked: false,
            enabled: true,
            help_context_id: 0,
            negotiate_position: NegotiatePosition::None,
            shortcut: None,
            visible: true,
            window_list: false,
        }
    }
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum NegotiatePosition {
    #[default]
    None = 0,
    Left = 1,
    Middle = 2,
    Right = 3,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum ShortCut {
    CtrlA,
    CtrlB,
    CtrlC,
    CtrlD,
    CtrlE,
    CtrlF,
    CtrlG,
    CtrlH,
    CtrlI,
    CtrlJ,
    CtrlK,
    CtrlL,
    CtrlM,
    CtrlN,
    CtrlO,
    CtrlP,
    CtrlQ,
    CtrlR,
    CtrlS,
    CtrlT,
    CtrlU,
    CtrlV,
    CtrlW,
    CtrlX,
    CtrlY,
    CtrlZ,
    F1,
    F2,
    F3,
    F4,
    F5,
    F6,
    F7,
    F8,
    F9,
    // F10 is not included.
    F11,
    F12,
    CtrlF1,
    CtrlF2,
    CtrlF3,
    CtrlF4,
    CtrlF5,
    CtrlF6,
    CtrlF7,
    CtrlF8,
    CtrlF9,
    // CtrlF10 is not included.
    CtrlF11,
    CtrlF12,
    ShiftF1,
    ShiftF2,
    ShiftF3,
    ShiftF4,
    ShiftF5,
    ShiftF6,
    ShiftF7,
    ShiftF8,
    ShiftF9,
    // ShiftF10 is not included.
    ShiftF11,
    ShiftF12,
    ShiftCtrlF1,
    ShiftCtrlF2,
    ShiftCtrlF3,
    ShiftCtrlF4,
    ShiftCtrlF5,
    ShiftCtrlF6,
    ShiftCtrlF7,
    ShiftCtrlF8,
    ShiftCtrlF9,
    // ShiftCtrlF10 is not included.
    ShiftCtrlF11,
    ShiftCtrlF12,
    CtrlIns,
    ShiftIns,
    Del,
    ShiftDel,
    AltBKsp,
}

impl ShortCut {
    pub fn from_str(s: &str) -> Option<Self> {
        match s {
            "^A" => Some(ShortCut::CtrlA),
            "^B" => Some(ShortCut::CtrlB),
            "^C" => Some(ShortCut::CtrlC),
            "^D" => Some(ShortCut::CtrlD),
            "^E" => Some(ShortCut::CtrlE),
            "^F" => Some(ShortCut::CtrlF),
            "^G" => Some(ShortCut::CtrlG),
            "^H" => Some(ShortCut::CtrlH),
            "^I" => Some(ShortCut::CtrlI),
            "^J" => Some(ShortCut::CtrlJ),
            "^K" => Some(ShortCut::CtrlK),
            "^L" => Some(ShortCut::CtrlL),
            "^M" => Some(ShortCut::CtrlM),
            "^N" => Some(ShortCut::CtrlN),
            "^O" => Some(ShortCut::CtrlO),
            "^P" => Some(ShortCut::CtrlP),
            "^Q" => Some(ShortCut::CtrlQ),
            "^R" => Some(ShortCut::CtrlR),
            "^S" => Some(ShortCut::CtrlS),
            "^T" => Some(ShortCut::CtrlT),
            "^U" => Some(ShortCut::CtrlU),
            "^V" => Some(ShortCut::CtrlV),
            "^W" => Some(ShortCut::CtrlW),
            "^X" => Some(ShortCut::CtrlX),
            "^Y" => Some(ShortCut::CtrlY),
            "^Z" => Some(ShortCut::CtrlZ),
            "{F1}" => Some(ShortCut::F1),
            "{F2}" => Some(ShortCut::F2),
            "{F3}" => Some(ShortCut::F3),
            "{F4}" => Some(ShortCut::F4),
            "{F5}" => Some(ShortCut::F5),
            "{F6}" => Some(ShortCut::F6),
            "{F7}" => Some(ShortCut::F7),
            "{F8}" => Some(ShortCut::F8),
            "{F9}" => Some(ShortCut::F9),
            "{F11}" => Some(ShortCut::F11),
            "{F12}" => Some(ShortCut::F12),
            "^{F1}" => Some(ShortCut::CtrlF1),
            "^{F2}" => Some(ShortCut::CtrlF2),
            "^{F3}" => Some(ShortCut::CtrlF3),
            "^{F4}" => Some(ShortCut::CtrlF4),
            "^{F5}" => Some(ShortCut::CtrlF5),
            "^{F6}" => Some(ShortCut::CtrlF6),
            "^{F7}" => Some(ShortCut::CtrlF7),
            "^{F8}" => Some(ShortCut::CtrlF8),
            "^{F9}" => Some(ShortCut::CtrlF9),
            "^{F11}" => Some(ShortCut::CtrlF11),
            "^{F12}" => Some(ShortCut::CtrlF12),
            "+{F1}" => Some(ShortCut::ShiftF1),
            "+{F2}" => Some(ShortCut::ShiftF2),
            "+{F3}" => Some(ShortCut::ShiftF3),
            "+{F4}" => Some(ShortCut::ShiftF4),
            "+{F5}" => Some(ShortCut::ShiftF5),
            "+{F6}" => Some(ShortCut::ShiftF6),
            "+{F7}" => Some(ShortCut::ShiftF7),
            "+{F8}" => Some(ShortCut::ShiftF8),
            "+{F9}" => Some(ShortCut::ShiftF9),
            "+{F11}" => Some(ShortCut::ShiftF11),
            "+{F12}" => Some(ShortCut::ShiftF12),
            "+^{F1}" => Some(ShortCut::ShiftCtrlF1),
            "+^{F2}" => Some(ShortCut::ShiftCtrlF2),
            "+^{F3}" => Some(ShortCut::ShiftCtrlF3),
            "+^{F4}" => Some(ShortCut::ShiftCtrlF4),
            "+^{F5}" => Some(ShortCut::ShiftCtrlF5),
            "+^{F6}" => Some(ShortCut::ShiftCtrlF6),
            "+^{F7}" => Some(ShortCut::ShiftCtrlF7),
            "+^{F8}" => Some(ShortCut::ShiftCtrlF8),
            "+^{F9}" => Some(ShortCut::ShiftCtrlF9),
            "+^{F11}" => Some(ShortCut::ShiftCtrlF11),
            "+^{F12}" => Some(ShortCut::ShiftCtrlF12),
            "^{INSERT}" => Some(ShortCut::CtrlIns),
            "+{INSERT}" => Some(ShortCut::ShiftIns),
            "{DEL}" => Some(ShortCut::Del),
            "+{DEL}" => Some(ShortCut::ShiftDel),
            "%{BKSP}" => Some(ShortCut::AltBKsp),
            _ => None,
        }
    }
}
