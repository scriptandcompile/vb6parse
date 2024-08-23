use crate::language::controls::{Appearance, DragMode, MousePointer, OLEDragMode, OLEDropMode};
use crate::VB6Color;

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct DirListBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}
