use crate::language::controls::{
    Appearance, DragMode, JustifyAlignment, MousePointer, OLEDropMode, Style,
};
use crate::language::VB6Color;

use image::DynamicImage;

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum OptionButtonValue {
    UnSelected = 0,
    Selected = 1,
}

#[derive(Debug, PartialEq, Clone)]
pub struct OptionButtonProperties<'a> {
    pub alignment: JustifyAlignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub caption: &'a str,
    pub causes_validation: bool,
    pub disabled_picture: Option<DynamicImage>,
    pub down_picture: Option<DynamicImage>,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub mask_color: VB6Color,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub right_to_left: bool,
    pub style: Style,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub use_mask_color: bool,
    pub value: OptionButtonValue,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for OptionButtonProperties<'_> {
    fn default() -> Self {
        OptionButtonProperties {
            alignment: JustifyAlignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            caption: "Option1",
            causes_validation: true,
            disabled_picture: None,
            down_picture: None,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            left: 30,
            mask_color: VB6Color::from_hex("&H00C0C0C0&").unwrap(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::None,
            picture: None,
            right_to_left: false,
            style: Style::Standard,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: "",
            top: 30,
            use_mask_color: false,
            value: OptionButtonValue::UnSelected,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}
