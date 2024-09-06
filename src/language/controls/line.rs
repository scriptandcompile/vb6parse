use crate::language::color::VB6Color;
use crate::language::controls::{DrawMode, DrawStyle};

/// Properties for a Line control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Line`](crate::language::controls::VB6ControlKind::Line).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub struct LineProperties {
    pub border_color: VB6Color,
    pub border_style: DrawStyle,
    pub border_width: i32,
    pub draw_mode: DrawMode,
    pub visible: bool,
    pub x1: i32,
    pub y1: i32,
    pub x2: i32,
    pub y2: i32,
}

impl Default for LineProperties {
    fn default() -> Self {
        LineProperties {
            border_color: VB6Color::from_hex("&H80000008&").unwrap(),
            border_style: DrawStyle::Solid,
            border_width: 1,
            draw_mode: DrawMode::CopyPen,
            visible: true,
            x1: 0,
            y1: 0,
            x2: 100,
            y2: 100,
        }
    }
}
