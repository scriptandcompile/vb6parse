use crate::language::controls::{NegotiatePosition, ShortCut};

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
