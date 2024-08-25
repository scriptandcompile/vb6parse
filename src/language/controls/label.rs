use crate::language::controls::{
    Alignment, Appearance, BackStyle, BorderStyle, DragMode, LinkMode, MousePointer, OLEDropMode,
};
use crate::VB6Color;

use image::DynamicImage;

#[derive(Debug, PartialEq, Clone)]
pub struct LabelProperties<'a> {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub auto_size: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub caption: &'a str,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub left: i32,
    pub link_item: &'a str,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: &'a str,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub tab_index: i32,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub use_mnemonic: bool,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
    pub word_wrap: bool,
}

impl Default for LabelProperties<'_> {
    fn default() -> Self {
        LabelProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            auto_size: false,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            back_style: BackStyle::Opaque,
            border_style: BorderStyle::None,
            caption: "Label1",
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            left: 30,
            link_item: "",
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "",
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::None,
            right_to_left: false,
            tab_index: 0,
            tool_tip_text: "",
            top: 30,
            use_mnemonic: true,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
            word_wrap: false,
        }
    }
}
