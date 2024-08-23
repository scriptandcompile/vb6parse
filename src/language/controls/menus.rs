/// Represents a VB6 menu control.
/// This should only be used as a child of a Form.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6MenuControl<'a> {
    pub name: &'a str,
    pub tag: &'a str,
    pub index: i32,
    pub properties: MenuProperties<'a>,
    pub sub_menus: Vec<VB6MenuControl<'a>>,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct MenuProperties<'a> {
    pub caption: &'a str,
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
            caption: "",
            enabled: true,
            help_context_id: 0,
            negotiate_position: NegotiatePosition::None,
            shortcut: None,
            visible: true,
            window_list: false,
        }
    }
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum NegotiatePosition {
    None = 0,
    Left = 1,
    Middle = 2,
    Right = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
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
