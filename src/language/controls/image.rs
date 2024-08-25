use crate::language::controls::{
    Appearance, BorderStyle, DragMode, MousePointer, OLEDragMode, OLEDropMode,
};

use image::DynamicImage;

#[derive(Debug, PartialEq, Clone)]
pub struct ImageProperties<'a> {
    pub appearance: Appearance,
    pub border_style: BorderStyle,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub left: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub stretch: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}
