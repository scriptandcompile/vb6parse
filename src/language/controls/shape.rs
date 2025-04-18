use crate::language::color::VB6Color;
use crate::language::controls::{BackStyle, DrawMode, DrawStyle};
use crate::parsers::Properties;

use num_enum::TryFromPrimitive;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum Shape {
    #[default]
    Rectangle = 0,
    Square = 1,
    Oval = 2,
    Circle = 3,
    RoundedRectangle = 4,
    RoundSquare = 5,
}

/// Properties for a Shape control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Shape`](crate::language::controls::VB6ControlKind::Shape).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
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

impl<'a> From<Properties<'a>> for ShapeProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut shape_prop = ShapeProperties::default();

        shape_prop.back_color = prop.get_color(b"BackColor".into(), shape_prop.back_color);
        shape_prop.back_style = prop.get_property(b"BackStyle".into(), shape_prop.back_style);
        shape_prop.border_color = prop.get_color(b"BorderColor".into(), shape_prop.border_color);
        shape_prop.border_style = prop.get_property(b"BorderStyle".into(), shape_prop.border_style);
        shape_prop.border_width = prop.get_i32(b"BorderWidth".into(), shape_prop.border_width);
        shape_prop.draw_mode = prop.get_property(b"DrawMode".into(), shape_prop.draw_mode);
        shape_prop.fill_color = prop.get_color(b"FillColor".into(), shape_prop.fill_color);
        shape_prop.fill_style = prop.get_property(b"FillStyle".into(), shape_prop.fill_style);
        shape_prop.height = prop.get_i32(b"Height".into(), shape_prop.height);
        shape_prop.left = prop.get_i32(b"Left".into(), shape_prop.left);
        shape_prop.shape = prop.get_property(b"Shape".into(), shape_prop.shape);
        shape_prop.top = prop.get_i32(b"Top".into(), shape_prop.top);
        shape_prop.visible = prop.get_bool(b"Visible".into(), shape_prop.visible);
        shape_prop.width = prop.get_i32(b"Width".into(), shape_prop.width);

        shape_prop
    }
}
