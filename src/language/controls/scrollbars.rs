use crate::language::controls::{DragMode, MousePointer};

use image::DynamicImage;

#[derive(Debug, PartialEq, Clone)]
pub struct ScrollBarProperties {
    pub causes_validation: bool,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub large_change: i32,
    pub left: i32,
    pub max: i32,
    pub min: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub right_to_left: bool,
    pub small_change: i32,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub top: i32,
    pub value: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ScrollBarProperties {
    fn default() -> Self {
        ScrollBarProperties {
            causes_validation: true,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            height: 30,
            help_context_id: 0,
            large_change: 1,
            left: 30,
            max: 32767,
            min: 0,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            right_to_left: false,
            small_change: 1,
            tab_index: 0,
            tab_stop: true,
            top: 30,
            value: 0,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}
