use crate::language::controls::{Appearance, DragMode, MousePointer, OLEDragMode, OLEDropMode};
use crate::VB6Color;

use image::DynamicImage;

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum ComboBoxStyle {
    DropDownCombo = 0,
    SimpleCombo = 1,
    DropDownList = 2,
}

#[derive(Debug, PartialEq, Clone)]
pub struct ComboBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
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
    pub integral_height: bool,
    //pub item_data: Vec<&'a str>,
    pub left: i32,
    // pub list: Vec<&'a str>,
    pub locked: bool,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub sorted: bool,
    pub style: ComboBoxStyle,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub text: &'a str,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ComboBoxProperties<'_> {
    fn default() -> Self {
        ComboBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H80000005&").unwrap(),
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
            integral_height: true,
            //item_data:
            left: 30,
            //list:
            locked: false,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::None,
            right_to_left: false,
            sorted: false,
            style: ComboBoxStyle::DropDownCombo,
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
