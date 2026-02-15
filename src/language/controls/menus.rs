//! Properties and structures for a `Menu` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::Menu`](crate::language::controls::ControlKind::Menu).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!
//! This should only be used as a child of a Form / `MDIForm`.

use std::convert::TryFrom;
use std::fmt::Display;
use std::str::FromStr;

use crate::errors::ErrorKind;
use crate::files::common::Properties;
use crate::language::controls::{Activation, Visibility};

use num_enum::TryFromPrimitive;
use serde::Serialize;

/// Represents a VB6 menu control.
/// This should only be used as a child of a Form / `MDIForm`.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct MenuControl {
    /// The name of the menu control.
    name: String,
    /// The tag of the menu control.
    tag: String,
    /// The index of the menu control.
    index: i32,
    /// The properties of the menu control.
    properties: MenuProperties,
    /// The sub-menus of the menu control.
    sub_menus: Vec<MenuControl>,
}

impl MenuControl {
    /// Creates a new `MenuControl` with the specified properties.
    ///
    /// # Arguments
    ///
    /// * `name` - The name of the menu control
    /// * `tag` - The tag of the menu control
    /// * `index` - The index of the menu control
    /// * `properties` - The properties of the menu control
    /// * `sub_menus` - The sub-menus of the menu control
    ///
    /// # Returns
    ///
    /// A new `MenuControl` instance.
    #[must_use]
    pub fn new(
        name: String,
        tag: String,
        index: i32,
        properties: MenuProperties,
        sub_menus: Vec<MenuControl>,
    ) -> Self {
        Self {
            name,
            tag,
            index,
            properties,
            sub_menus,
        }
    }

    /// Returns the name of the menu control.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the tag of the menu control.
    #[must_use]
    pub fn tag(&self) -> &str {
        &self.tag
    }

    /// Returns the index of the menu control.
    #[must_use]
    pub fn index(&self) -> i32 {
        self.index
    }

    /// Returns a reference to the menu properties.
    #[must_use]
    pub fn properties(&self) -> &MenuProperties {
        &self.properties
    }

    /// Returns a slice of the sub-menus.
    #[must_use]
    pub fn sub_menus(&self) -> &[MenuControl] {
        &self.sub_menus
    }

    /// Consumes the menu control and returns its name.
    #[must_use]
    pub fn into_name(self) -> String {
        self.name
    }

    /// Consumes the menu control and returns its tag.
    #[must_use]
    pub fn into_tag(self) -> String {
        self.tag
    }

    /// Consumes the menu control and returns its properties.
    #[must_use]
    pub fn into_properties(self) -> MenuProperties {
        self.properties
    }

    /// Consumes the menu control and returns its sub-menus.
    #[must_use]
    pub fn into_sub_menus(self) -> Vec<MenuControl> {
        self.sub_menus
    }

    /// Consumes the menu control and returns all of its parts as a tuple.
    ///
    /// # Returns
    ///
    /// A tuple containing `(name, tag, index, properties, sub_menus)`.
    #[must_use]
    pub fn into_parts(self) -> (String, String, i32, MenuProperties, Vec<MenuControl>) {
        (
            self.name,
            self.tag,
            self.index,
            self.properties,
            self.sub_menus,
        )
    }
}

impl Display for MenuControl {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "MenuControl: {}", self.name)
    }
}

/// Properties for a `Menu` control.
///
/// This is used as an enum variant of
/// [`ControlKind::Menu`](crate::language::controls::ControlKind::Menu).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
/// This is represented within the parsing code independently of
/// [`MenuControl`](crate::language::controls::MenuControl)'s.
///
/// This is currently redundant, but is included for the future where the correct
/// behavior of a menu control only being a child of a form is enforced.
///
/// As is, the parser will not enforce this, but the VB6 IDE will.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct MenuProperties {
    /// Caption of the menu.
    pub caption: String,
    /// Whether or not the menu is checked.
    pub checked: bool,
    /// Enabled state of the menu.
    pub enabled: Activation,
    /// Help context ID of the menu.
    pub help_context_id: i32,
    /// Negotiation position of the menu.
    pub negotiate_position: NegotiatePosition,
    /// Shortcut key of the menu.
    pub shortcut: Option<ShortCut>,
    /// Visibility of the menu.
    pub visible: Visibility,
    /// Whether the menu is part of the window list.
    pub window_list: bool,
}

impl Default for MenuProperties {
    fn default() -> Self {
        MenuProperties {
            caption: String::new(),
            checked: false,
            enabled: Activation::Enabled,
            help_context_id: 0,
            negotiate_position: NegotiatePosition::None,
            shortcut: None,
            visible: Visibility::Visible,
            window_list: false,
        }
    }
}

impl From<Properties> for MenuProperties {
    fn from(prop: Properties) -> Self {
        let mut menu_prop = MenuProperties::default();

        menu_prop.caption = match prop.get("Caption") {
            Some(caption) => caption.into(),
            None => menu_prop.caption,
        };
        menu_prop.checked = prop.get_bool("Checked", menu_prop.checked);
        menu_prop.enabled = prop.get_property("Enabled", menu_prop.enabled);
        menu_prop.help_context_id = prop.get_i32("HelpContextID", menu_prop.help_context_id);
        menu_prop.negotiate_position =
            prop.get_property("NegotiationPosition", menu_prop.negotiate_position);
        menu_prop.shortcut = prop.get_option("Shortcut", menu_prop.shortcut);
        menu_prop.visible = prop.get_property("Visible", menu_prop.visible);
        menu_prop.window_list = prop.get_bool("WindowList", menu_prop.window_list);

        menu_prop
    }
}

/// Determines whether or not top-level Menu controls are displayed on the menu
/// bar while a linked object or embedded object  on a form is active and
/// displaying its menus.
///
/// Using the `NegotiatePosition` property, you determine the individual menus on
/// the menu bar of your form that share (or negotiate) menu bar space with the
/// menus of an active object on the form. Any menu with `NegotiatePosition` set
/// to a nonzero value is displayed on the menu bar of the form along with menus
/// from the active object.
///
/// If the `NegotiateMenus` property of the corresponding `FormProperties` is
/// false, the setting of the `NegotiatePosition` property has no effect.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa278135(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum NegotiatePosition {
    /// The menu is not displayed on the menu bar.
    ///
    /// This is the default value.
    #[default]
    None = 0,
    /// The menu is displayed at the left end of the menu bar when the object is active.
    Left = 1,
    /// The menu is displayed in the middle of the menu bar when the object is active.
    Middle = 2,
    /// The menu is displayed at the right end of the menu bar when the object is active.
    Right = 3,
}

impl FromStr for NegotiatePosition {
    type Err = ErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "0" => Ok(NegotiatePosition::None),
            "1" => Ok(NegotiatePosition::Left),
            "2" => Ok(NegotiatePosition::Middle),
            "3" => Ok(NegotiatePosition::Right),
            _ => Err(ErrorKind::FormInvalidNegotiatePosition { value: s.to_string() }),
        }
    }
}

impl TryFrom<&str> for NegotiatePosition {
    type Error = ErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        NegotiatePosition::from_str(value)
    }
}

impl Display for NegotiatePosition {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            NegotiatePosition::None => "None",
            NegotiatePosition::Left => "Left",
            NegotiatePosition::Middle => "Middle",
            NegotiatePosition::Right => "Right",
        };
        write!(f, "{text}")
    }
}

/// Represents a keyboard shortcut for a menu item.
///
/// Note:
///
/// In addition to shortcut keys, you can also assign access keys to
/// commands, menus, and controls by using an ampersand (&) in the `Caption`
/// property setting.
///
/// The F10, Ctrl+F10, Shift+F10, and Ctrl+Shift+F10 keys are not valid shortcut keys.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Copy, Hash, PartialOrd, Ord)]
pub enum ShortCut {
    /// Ctrl + A
    ///
    /// This is stored in the Form file as "^A"
    CtrlA,
    /// Ctrl + B
    ///
    /// This is stored in the Form file as "^B"
    CtrlB,
    /// Ctrl + C
    ///
    /// This is stored in the Form file as "^C"
    CtrlC,
    /// Ctrl + D
    ///
    /// This is stored in the Form file as "^D"
    CtrlD,
    /// Ctrl + E
    ///
    /// This is stored in the Form file as "^E"
    CtrlE,
    /// Ctrl + F
    ///
    /// This is stored in the Form file as "^F"
    CtrlF,
    /// Ctrl + G
    ///
    /// This is stored in the Form file as "^G"
    CtrlG,
    /// Ctrl + H
    ///
    /// This is stored in the Form file as "^H"
    CtrlH,
    /// Ctrl + I
    ///
    /// This is stored in the Form file as "^I"
    CtrlI,
    /// Ctrl + J
    ///
    /// This is stored in the Form file as "^J"
    CtrlJ,
    /// Ctrl + K
    ///
    /// This is stored in the Form file as "^K"
    CtrlK,
    /// Ctrl + L
    ///
    /// This is stored in the Form file as "^L"
    CtrlL,
    /// Ctrl + M
    ///
    /// This is stored in the Form file as "^M"
    CtrlM,
    /// Ctrl + N
    ///
    /// This is stored in the Form file as "^N"
    CtrlN,
    /// Ctrl + O
    ///
    /// This is stored in the Form file as "^O"
    CtrlO,
    /// Ctrl + P
    ///
    /// This is stored in the Form file as "^P"
    CtrlP,
    /// Ctrl + Q
    ///
    /// This is stored in the Form file as "^Q"
    CtrlQ,
    /// Ctrl + R
    ///
    /// This is stored in the Form file as "^R"
    CtrlR,
    /// Ctrl + S
    ///
    /// This is stored in the Form file as "^S"
    CtrlS,
    /// Ctrl + T
    ///
    /// This is stored in the Form file as "^T"
    CtrlT,
    /// Ctrl + U
    ///
    /// This is stored in the Form file as "^U"
    CtrlU,
    /// Ctrl + V
    ///
    /// This is stored in the Form file as "^V"
    CtrlV,
    /// Ctrl + W
    ///
    /// This is stored in the Form file as "^W"
    CtrlW,
    /// Ctrl + X
    ///
    /// This is stored in the Form file as "^X"
    CtrlX,
    /// Ctrl + Y
    ///
    /// This is stored in the Form file as "^Y"
    CtrlY,
    /// Ctrl + Z
    ///
    /// This is stored in the Form file as "^Z"
    CtrlZ,
    /// The F1 function key.
    ///
    /// This is stored in the Form file as "{F1}"
    F1,
    /// The F2 function key.
    ///
    /// This is stored in the Form file as "{F2}"
    F2,
    /// The F3 function key.
    ///
    /// This is stored in the Form file as "{F3}"
    F3,
    /// The F4 function key.
    ///
    /// This is stored in the Form file as "{F4}"
    F4,
    /// The F5 function key.
    ///
    /// This is stored in the Form file as "{F5}"
    F5,
    /// The F6 function key.
    ///
    /// This is stored in the Form file as "{F6}"
    F6,
    /// The F7 function key.
    ///
    /// This is stored in the Form file as "{F7}"
    F7,
    /// The F8 function key.
    ///
    /// This is stored in the Form file as "{F8}"
    F8,
    /// The F9 function key.
    ///
    /// This is stored in the Form file as "{F9}"
    F9,
    /// The F11 function key.
    ///
    /// This is stored in the Form file as "{F11}"
    F11,
    /// The F12 function key.
    ///
    /// This is stored in the Form file as "{F12}"
    F12,
    /// Ctrl + F1
    ///
    /// This is stored in the Form file as "^{F1}"
    CtrlF1,
    /// Ctrl + F2
    ///
    /// This is stored in the Form file as "^{F2}"
    CtrlF2,
    /// Ctrl + F3
    ///
    /// This is stored in the Form file as "^{F3}"
    CtrlF3,
    /// Ctrl + F4
    ///
    /// This is stored in the Form file as "^{F4}"
    CtrlF4,
    /// Ctrl + F5
    ///
    /// This is stored in the Form file as "^{F5}"
    CtrlF5,
    /// Ctrl + F6
    ///
    /// This is stored in the Form file as "^{F6}"
    CtrlF6,
    /// Ctrl + F7
    ///
    /// This is stored in the Form file as "^{F7}"
    CtrlF7,
    /// Ctrl + F8
    ///
    /// This is stored in the Form file as "^{F8}"
    CtrlF8,
    /// Ctrl + F9
    ///
    /// This is stored in the Form file as "^{F9}"
    CtrlF9,
    /// Ctrl + F11
    ///
    /// This is stored in the Form file as "^{F11}"
    CtrlF11,
    /// Ctrl + F12
    ///
    /// This is stored in the Form file as "^{F12}"
    CtrlF12,
    /// Shift + F1
    ///
    /// This is stored in the Form file as "+{F1}"
    ShiftF1,
    /// Shift + F2
    ///
    /// This is stored in the Form file as "+{F2}"
    ShiftF2,
    /// Shift + F3
    ///
    /// This is stored in the Form file as "+{F3}"
    ShiftF3,
    /// Shift + F4
    ///
    /// This is stored in the Form file as "+{F4}"
    ShiftF4,
    /// Shift + F5
    ///
    /// This is stored in the Form file as "+{F5}"
    ShiftF5,
    /// Shift + F6
    ///
    /// This is stored in the Form file as "+{F6}"
    ShiftF6,
    /// Shift + F7
    ///
    /// This is stored in the Form file as "+{F7}"
    ShiftF7,
    /// Shift + F8
    ///
    /// This is stored in the Form file as "+{F8}"
    ShiftF8,
    /// Shift + F9
    ///
    /// This is stored in the Form file as "+{F9}"
    ShiftF9,
    /// Shift + F11
    ///
    /// This is stored in the Form file as "+{F11}"
    ShiftF11,
    /// Shift + F12
    ///
    /// This is stored in the Form file as "+{F12}"
    ShiftF12,
    /// Shift + Ctrl + F1
    ///
    /// This is stored in the Form file as "+^{F1}"
    ShiftCtrlF1,
    /// Shift + Ctrl + F2
    ///
    /// This is stored in the Form file as "+^{F2}"
    ShiftCtrlF2,
    /// Shift + Ctrl + F3
    ///
    /// This is stored in the Form file as "+^{F3}"
    ShiftCtrlF3,
    /// Shift + Ctrl + F4
    ///
    /// This is stored in the Form file as "+^{F4}"
    ShiftCtrlF4,
    /// Shift + Ctrl + F5
    ///
    /// This is stored in the Form file as "+^{F5}"
    ShiftCtrlF5,
    /// Shift + Ctrl + F6
    ///
    /// This is stored in the Form file as "+^{F6}"
    ShiftCtrlF6,
    /// Shift + Ctrl + F7
    ///
    /// This is stored in the Form file as "+^{F7}"
    ShiftCtrlF7,
    /// Shift + Ctrl + F8
    ///
    /// This is stored in the Form file as "+^{F8}"
    ShiftCtrlF8,
    /// Shift + Ctrl + F9
    ///
    /// This is stored in the Form file as "+^{F9}"
    ShiftCtrlF9,
    /// Shift + Ctrl + F11
    ///
    /// This is stored in the Form file as "+^{F11}"
    ShiftCtrlF11,
    /// Shift + Ctrl + F12
    ///
    /// This is stored in the Form file as "+^{F12}"
    ShiftCtrlF12,
    /// Ctrl + Insert
    ///
    /// This is stored in the Form file as "^{INSERT}"
    CtrlIns,
    /// Shift + Insert
    ///
    /// This is stored in the Form file as "+{INSERT}"
    ShiftIns,
    /// Delete
    ///
    /// This is stored in the Form file as "{DEL}"
    Del,
    /// Shift + Delete
    ///
    /// This is stored in the Form file as "+{DEL}"
    ShiftDel,
    /// Alt + Backspace
    ///
    /// This is stored in the Form file as "%{BKSP}"
    AltBKsp,
}

impl Display for ShortCut {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            ShortCut::CtrlA => "Ctrl+A",
            ShortCut::CtrlB => "Ctrl+B",
            ShortCut::CtrlC => "Ctrl+C",
            ShortCut::CtrlD => "Ctrl+D",
            ShortCut::CtrlE => "Ctrl+E",
            ShortCut::CtrlF => "Ctrl+F",
            ShortCut::CtrlG => "Ctrl+G",
            ShortCut::CtrlH => "Ctrl+H",
            ShortCut::CtrlI => "Ctrl+I",
            ShortCut::CtrlJ => "Ctrl+J",
            ShortCut::CtrlK => "Ctrl+K",
            ShortCut::CtrlL => "Ctrl+L",
            ShortCut::CtrlM => "Ctrl+M",
            ShortCut::CtrlN => "Ctrl+N",
            ShortCut::CtrlO => "Ctrl+O",
            ShortCut::CtrlP => "Ctrl+P",
            ShortCut::CtrlQ => "Ctrl+Q",
            ShortCut::CtrlR => "Ctrl+R",
            ShortCut::CtrlS => "Ctrl+S",
            ShortCut::CtrlT => "Ctrl+T",
            ShortCut::CtrlU => "Ctrl+U",
            ShortCut::CtrlV => "Ctrl+V",
            ShortCut::CtrlW => "Ctrl+W",
            ShortCut::CtrlX => "Ctrl+X",
            ShortCut::CtrlY => "Ctrl+Y",
            ShortCut::CtrlZ => "Ctrl+Z",
            ShortCut::F1 => "F1",
            ShortCut::F2 => "F2",
            ShortCut::F3 => "F3",
            ShortCut::F4 => "F4",
            ShortCut::F5 => "F5",
            ShortCut::F6 => "F6",
            ShortCut::F7 => "F7",
            ShortCut::F8 => "F8",
            ShortCut::F9 => "F9",
            ShortCut::F11 => "F11",
            ShortCut::F12 => "F12",
            ShortCut::CtrlF1 => "Ctrl+F1",
            ShortCut::CtrlF2 => "Ctrl+F2",
            ShortCut::CtrlF3 => "Ctrl+F3",
            ShortCut::CtrlF4 => "Ctrl+F4",
            ShortCut::CtrlF5 => "Ctrl+F5",
            ShortCut::CtrlF6 => "Ctrl+F6",
            ShortCut::CtrlF7 => "Ctrl+F7",
            ShortCut::CtrlF8 => "Ctrl+F8",
            ShortCut::CtrlF9 => "Ctrl+F9",
            ShortCut::CtrlF11 => "Ctrl+F11",
            ShortCut::CtrlF12 => "Ctrl+F12",
            ShortCut::ShiftF1 => "Shift+F1",
            ShortCut::ShiftF2 => "Shift+F2",
            ShortCut::ShiftF3 => "Shift+F3",
            ShortCut::ShiftF4 => "Shift+F4",
            ShortCut::ShiftF5 => "Shift+F5",
            ShortCut::ShiftF6 => "Shift+F6",
            ShortCut::ShiftF7 => "Shift+F7",
            ShortCut::ShiftF8 => "Shift+F8",
            ShortCut::ShiftF9 => "Shift+F9",
            ShortCut::ShiftF11 => "Shift+F11",
            ShortCut::ShiftF12 => "Shift+F12",
            ShortCut::ShiftCtrlF1 => "Shift+Ctrl+F1",
            ShortCut::ShiftCtrlF2 => "Shift+Ctrl+F2",
            ShortCut::ShiftCtrlF3 => "Shift+Ctrl+F3",
            ShortCut::ShiftCtrlF4 => "Shift+Ctrl+F4",
            ShortCut::ShiftCtrlF5 => "Shift+Ctrl+F5",
            ShortCut::ShiftCtrlF6 => "Shift+Ctrl+F6",
            ShortCut::ShiftCtrlF7 => "Shift+Ctrl+F7",
            ShortCut::ShiftCtrlF8 => "Shift+Ctrl+F8",
            ShortCut::ShiftCtrlF9 => "Shift+Ctrl+F9",
            ShortCut::ShiftCtrlF11 => "Shift+Ctrl+F11",
            ShortCut::ShiftCtrlF12 => "Shift+Ctrl+F12",
            ShortCut::CtrlIns => "Ctrl+Insert",
            ShortCut::ShiftIns => "Shift+Insert",
            ShortCut::Del => "Delete",
            ShortCut::ShiftDel => "Shift+Delete",
            ShortCut::AltBKsp => "Alt+Backspace",
        };
        write!(f, "{text}")
    }
}

impl FromStr for ShortCut {
    type Err = ErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        ShortCut::try_from(s)
    }
}

impl TryFrom<&str> for ShortCut {
    type Error = ErrorKind;

    fn try_from(s: &str) -> Result<Self, ErrorKind> {
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
            _ => Err(ErrorKind::FormShortCutUnparsable),
        }
    }
}
