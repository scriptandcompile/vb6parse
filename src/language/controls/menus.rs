use bstr::BString;
use num_enum::TryFromPrimitive;
use serde::Serialize;

use crate::parsers::Properties;

use crate::errors::VB6ErrorKind;

/// Represents a VB6 menu control.
/// This should only be used as a child of a Form.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6MenuControl {
    pub name: BString,
    pub tag: BString,
    pub index: i32,
    pub properties: MenuProperties,
    pub sub_menus: Vec<VB6MenuControl>,
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
/// This is currently redundant, but is included for the future where the correct
/// behavior of a menu control only being a child of a form is enforced.
///
/// As is, the parser will not enforce this, but the VB6 IDE will.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct MenuProperties {
    pub caption: BString,
    pub checked: bool,
    pub enabled: bool,
    pub help_context_id: i32,
    pub negotiate_position: NegotiatePosition,
    pub shortcut: Option<ShortCut>,
    pub visible: bool,
    pub window_list: bool,
}

impl Default for MenuProperties {
    fn default() -> Self {
        MenuProperties {
            caption: BString::from(""),
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

impl<'a> From<Properties<'a>> for MenuProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut menu_prop = MenuProperties::default();

        menu_prop.caption = match prop.get(b"Caption".into()) {
            Some(caption) => caption.into(),
            None => menu_prop.caption,
        };
        menu_prop.checked = prop.get_bool(b"Checked".into(), menu_prop.checked);
        menu_prop.enabled = prop.get_bool(b"Enabled".into(), menu_prop.enabled);
        menu_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), menu_prop.help_context_id);
        menu_prop.negotiate_position =
            prop.get_property(b"NegotiationPosition".into(), menu_prop.negotiate_position);
        menu_prop.shortcut = prop.get_option(b"Shortcut".into(), menu_prop.shortcut);
        menu_prop.visible = prop.get_bool(b"Visible".into(), menu_prop.visible);
        menu_prop.window_list = prop.get_bool(b"WindowList".into(), menu_prop.window_list);

        menu_prop
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

impl TryFrom<&str> for ShortCut {
    type Error = VB6ErrorKind;

    fn try_from(s: &str) -> Result<Self, VB6ErrorKind> {
        match s {
            "^A" => Ok(ShortCut::CtrlA),
            "^B" => Ok(ShortCut::CtrlB),
            "^C" => Ok(ShortCut::CtrlC),
            "^D" => Ok(ShortCut::CtrlD),
            "^E" => Ok(ShortCut::CtrlE),
            "^F" => Ok(ShortCut::CtrlF),
            "^G" => Ok(ShortCut::CtrlG),
            "^H" => Ok(ShortCut::CtrlH),
            "^I" => Ok(ShortCut::CtrlI),
            "^J" => Ok(ShortCut::CtrlJ),
            "^K" => Ok(ShortCut::CtrlK),
            "^L" => Ok(ShortCut::CtrlL),
            "^M" => Ok(ShortCut::CtrlM),
            "^N" => Ok(ShortCut::CtrlN),
            "^O" => Ok(ShortCut::CtrlO),
            "^P" => Ok(ShortCut::CtrlP),
            "^Q" => Ok(ShortCut::CtrlQ),
            "^R" => Ok(ShortCut::CtrlR),
            "^S" => Ok(ShortCut::CtrlS),
            "^T" => Ok(ShortCut::CtrlT),
            "^U" => Ok(ShortCut::CtrlU),
            "^V" => Ok(ShortCut::CtrlV),
            "^W" => Ok(ShortCut::CtrlW),
            "^X" => Ok(ShortCut::CtrlX),
            "^Y" => Ok(ShortCut::CtrlY),
            "^Z" => Ok(ShortCut::CtrlZ),
            "{F1}" => Ok(ShortCut::F1),
            "{F2}" => Ok(ShortCut::F2),
            "{F3}" => Ok(ShortCut::F3),
            "{F4}" => Ok(ShortCut::F4),
            "{F5}" => Ok(ShortCut::F5),
            "{F6}" => Ok(ShortCut::F6),
            "{F7}" => Ok(ShortCut::F7),
            "{F8}" => Ok(ShortCut::F8),
            "{F9}" => Ok(ShortCut::F9),
            "{F11}" => Ok(ShortCut::F11),
            "{F12}" => Ok(ShortCut::F12),
            "^{F1}" => Ok(ShortCut::CtrlF1),
            "^{F2}" => Ok(ShortCut::CtrlF2),
            "^{F3}" => Ok(ShortCut::CtrlF3),
            "^{F4}" => Ok(ShortCut::CtrlF4),
            "^{F5}" => Ok(ShortCut::CtrlF5),
            "^{F6}" => Ok(ShortCut::CtrlF6),
            "^{F7}" => Ok(ShortCut::CtrlF7),
            "^{F8}" => Ok(ShortCut::CtrlF8),
            "^{F9}" => Ok(ShortCut::CtrlF9),
            "^{F11}" => Ok(ShortCut::CtrlF11),
            "^{F12}" => Ok(ShortCut::CtrlF12),
            "+{F1}" => Ok(ShortCut::ShiftF1),
            "+{F2}" => Ok(ShortCut::ShiftF2),
            "+{F3}" => Ok(ShortCut::ShiftF3),
            "+{F4}" => Ok(ShortCut::ShiftF4),
            "+{F5}" => Ok(ShortCut::ShiftF5),
            "+{F6}" => Ok(ShortCut::ShiftF6),
            "+{F7}" => Ok(ShortCut::ShiftF7),
            "+{F8}" => Ok(ShortCut::ShiftF8),
            "+{F9}" => Ok(ShortCut::ShiftF9),
            "+{F11}" => Ok(ShortCut::ShiftF11),
            "+{F12}" => Ok(ShortCut::ShiftF12),
            "+^{F1}" => Ok(ShortCut::ShiftCtrlF1),
            "+^{F2}" => Ok(ShortCut::ShiftCtrlF2),
            "+^{F3}" => Ok(ShortCut::ShiftCtrlF3),
            "+^{F4}" => Ok(ShortCut::ShiftCtrlF4),
            "+^{F5}" => Ok(ShortCut::ShiftCtrlF5),
            "+^{F6}" => Ok(ShortCut::ShiftCtrlF6),
            "+^{F7}" => Ok(ShortCut::ShiftCtrlF7),
            "+^{F8}" => Ok(ShortCut::ShiftCtrlF8),
            "+^{F9}" => Ok(ShortCut::ShiftCtrlF9),
            "+^{F11}" => Ok(ShortCut::ShiftCtrlF11),
            "+^{F12}" => Ok(ShortCut::ShiftCtrlF12),
            "^{INSERT}" => Ok(ShortCut::CtrlIns),
            "+{INSERT}" => Ok(ShortCut::ShiftIns),
            "{DEL}" => Ok(ShortCut::Del),
            "+{DEL}" => Ok(ShortCut::ShiftDel),
            "%{BKSP}" => Ok(ShortCut::AltBKsp),
            _ => Err(VB6ErrorKind::ShortCutUnparseable),
        }
    }
}
