use crate::language::color::{Color, VB_WINDOW_TEXT};
use crate::language::controls::{DrawMode, DrawStyle, Visibility};
use crate::parsers::Properties;

use serde::Serialize;

/// Properties for a `Line` control.
///
/// This is used as an enum variant of
/// [`ControlKind::Line`](crate::language::controls::ControlKind::Line).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct LineProperties {
    pub border_color: Color,
    pub border_style: DrawStyle,
    pub border_width: i32,
    pub draw_mode: DrawMode,
    pub visible: Visibility,
    pub x1: i32,
    pub y1: i32,
    pub x2: i32,
    pub y2: i32,
}

impl Default for LineProperties {
    fn default() -> Self {
        LineProperties {
            border_color: VB_WINDOW_TEXT,
            border_style: DrawStyle::Solid,
            border_width: 1,
            draw_mode: DrawMode::CopyPen,
            visible: Visibility::Visible,
            x1: 0,
            y1: 0,
            x2: 100,
            y2: 100,
        }
    }
}

impl From<Properties> for LineProperties {
    fn from(prop: Properties) -> Self {
        let mut line_prop = LineProperties::default();

        line_prop.border_color = prop.get_color("BorderColor", line_prop.border_color);
        line_prop.border_style = prop.get_property("BorderStyle", line_prop.border_style);
        line_prop.border_width = prop.get_i32("BorderWidth", line_prop.border_width);
        line_prop.draw_mode = prop.get_property("DrawMode", line_prop.draw_mode);
        line_prop.visible = prop.get_property("Visible", line_prop.visible);
        line_prop.x1 = prop.get_i32("X1", line_prop.x1);
        line_prop.y1 = prop.get_i32("Y1", line_prop.y1);
        line_prop.x2 = prop.get_i32("X2", line_prop.x2);
        line_prop.y2 = prop.get_i32("Y2", line_prop.y2);

        line_prop
    }
}
