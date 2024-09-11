use crate::language::color::VB6Color;
use crate::language::controls::{DrawMode, DrawStyle};

use super::BackStyle;

/// Properties for a Shape control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Shape`](crate::language::controls::VB6ControlKind::Shape).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub struct ShapeProperties {
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_color: VB6Color,
    pub border_style: DrawStyle,
    pub border_width: i32,
    pub draw_mode: DrawMode,
    pub fill_color: VB6Color,
    pub fill_style: DrawStyle,
    pub height: i32,
    pub left: i32,
    pub shape: Shape,
    pub top: i32,
    pub visible: bool,
    pub width: i32,
}

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]

pub enum Shape {
    Rectangle = 0,
    Square = 1,
    Oval = 2,
    Circle = 3,
    RoundedRectangle = 4,
    RoundSquare = 5,
}

impl Default for ShapeProperties {
    fn default() -> Self {
        ShapeProperties {
            back_color: VB6Color::System { index: 5 },
            back_style: BackStyle::Transparent,
            border_color: VB6Color::System { index: 8 },
            border_style: DrawStyle::Solid,
            border_width: 1,
            draw_mode: DrawMode::CopyPen,
            fill_color: VB6Color::RGB {
                red: 0,
                green: 0,
                blue: 0,
            },
            fill_style: DrawStyle::Transparent,
            height: 355,
            left: 30,
            shape: Shape::Rectangle,
            top: 200,
            visible: true,
            width: 355,
        }
    }
}
