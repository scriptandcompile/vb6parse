use crate::language::controls::{Appearance, DragMode, MousePointer, OLEDropMode, Style};
use crate::VB6Color;

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct CommandButtonProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub cancel: bool,
    pub caption: &'a str,
    pub causes_validation: bool,
    pub default: bool,
    //pub disabled_picture: Option<ImageBuffer>,
    //pub down_picture: Option<ImageBuffer>,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub mask_color: VB6Color,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    //pub picture: Option<ImageBuffer>,
    pub right_to_left: bool,
    pub style: Style,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub use_mask_color: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for CommandButtonProperties<'_> {
    fn default() -> Self {
        CommandButtonProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            cancel: false,
            caption: "Command1",
            causes_validation: true,
            default: false,
            drag_mode: DragMode::Manual,
            enabled: true,
            height: 30,
            help_context_id: 0,
            left: 30,
            mask_color: VB6Color::from_hex("&H00C0C0C0&").unwrap(),
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::None,
            right_to_left: false,
            style: Style::Standard,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: "",
            top: 30,
            use_mask_color: false,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}
