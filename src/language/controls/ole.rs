use crate::language::controls::{
    Appearance, BackStyle, BorderStyle, DragMode, MousePointer, SizeMode,
};
use crate::language::VB6Color;

use image::DynamicImage;

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum OLETypeAllowed {
    Link = 0,
    Embedded = 1,
    Either = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum UpdateOptions {
    Automatic = 0,
    Frozen = 1,
    Manual = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum AutoActivate {
    Manual = 0,
    GetFocus = 1,
    DoubleClick = 2,
    Automatic = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DisplayType {
    Content = 0,
    Icon = 1,
}

#[derive(Debug, PartialEq, Clone)]
pub struct OLEProperties<'a> {
    pub appearance: Appearance,
    pub auto_activate: AutoActivate,
    pub auto_verb_menu: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub class: Option<&'a str>,
    pub data_field: &'a str,
    pub data_source: &'a str,
    pub display_type: DisplayType,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub host_name: &'a str,
    pub left: i32,
    pub misc_flags: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_allowed: bool,
    pub ole_type_allowed: OLETypeAllowed,
    pub size_mode: SizeMode,
    //pub source_doc: &'a str,
    //pub source_item: &'a str,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub top: i32,
    pub update_options: UpdateOptions,
    pub verb: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}
