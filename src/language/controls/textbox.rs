use crate::language::controls::{
    Alignment, Appearance, BorderStyle, DragMode, LinkMode, MousePointer, OLEDragMode, OLEDropMode,
};
use crate::VB6Color;

use image::DynamicImage;

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum ScrollBars {
    None = 0,
    Horizontal = 1,
    Vertical = 2,
    Both = 3,
}

#[derive(Debug, PartialEq, Clone)]
pub struct TextBoxProperties<'a> {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub hide_selection: bool,
    pub left: i32,
    pub link_item: &'a str,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: &'a str,
    pub locked: bool,
    pub max_length: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub multi_line: bool,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub password_char: Option<char>,
    pub right_to_left: bool,
    pub scroll_bars: ScrollBars,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub text: &'a str,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for TextBoxProperties<'_> {
    fn default() -> Self {
        TextBoxProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H80000005&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            causes_validation: true,
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000008&").unwrap(),
            height: 30,
            help_context_id: 0,
            hide_selection: true,
            left: 30,
            link_item: "",
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "",
            locked: false,
            max_length: 0,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            multi_line: false,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::None,
            password_char: None,
            right_to_left: false,
            scroll_bars: ScrollBars::None,
            tab_index: 0,
            tab_stop: true,
            text: "",
            tool_tip_text: "",
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}
