//! Properties for Shape controls.
//!
//! This is used as an enum variant of
//! [`ControlKind::Shape`](crate::language::controls::ControlKind::Shape).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use std::convert::TryFrom;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

use crate::language::color::Color;
use crate::language::controls::{BackStyle, DrawMode, DrawStyle, Visibility};
use crate::files::common::Properties;

use num_enum::TryFromPrimitive;
use serde::Serialize;

/// The specific kind of shape to draw for a `Shape` control.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445683(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum Shape {
    /// A rectangle.
    ///
    /// This is the default shape.
    #[default]
    Rectangle = 0,
    /// A square.
    Square = 1,
    /// An oval.
    Oval = 2,
    /// A circle.
    Circle = 3,
    /// A rounded rectangle.
    RoundedRectangle = 4,
    /// A rounded square.
    RoundSquare = 5,
}

impl TryFrom<&str> for Shape {
    type Error = crate::errors::FormErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(Shape::Rectangle),
            "1" => Ok(Shape::Square),
            "2" => Ok(Shape::Oval),
            "3" => Ok(Shape::Circle),
            "4" => Ok(Shape::RoundedRectangle),
            "5" => Ok(Shape::RoundSquare),
            _ => Err(crate::errors::FormErrorKind::InvalidShape(
                value.to_string(),
            )),
        }
    }
}

impl FromStr for Shape {
    type Err = crate::errors::FormErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Shape::try_from(s)
    }
}

impl Display for Shape {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            Shape::Rectangle => "Rectangle",
            Shape::Square => "Square",
            Shape::Oval => "Oval",
            Shape::Circle => "Circle",
            Shape::RoundedRectangle => "RoundedRectangle",
            Shape::RoundSquare => "RoundSquare",
        };
        write!(f, "{text}")
    }
}

/// Properties for a `Shape` control.
///
/// This is used as an enum variant of
/// [`ControlKind::Shape`](crate::language::controls::ControlKind::Shape).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct ShapeProperties {
    /// Background color of the shape.
    pub back_color: Color,
    /// Background style of the shape.
    pub back_style: BackStyle,
    /// Border color of the shape.
    pub border_color: Color,
    /// Border style of the shape.
    pub border_style: DrawStyle,
    /// Border width of the shape.
    pub border_width: i32,
    /// Draw mode of the shape.
    pub draw_mode: DrawMode,
    /// Fill color of the shape.
    pub fill_color: Color,
    /// Fill style of the shape.
    pub fill_style: DrawStyle,
    /// Height of the shape control.
    pub height: i32,
    /// Left position of the shape control.
    pub left: i32,
    /// Specific shape to draw.
    pub shape: Shape,
    /// Top position of the shape control.
    pub top: i32,
    /// Visibility of the shape control.
    pub visible: Visibility,
    /// Width of the shape control.
    pub width: i32,
}

impl Default for ShapeProperties {
    fn default() -> Self {
        ShapeProperties {
            back_color: Color::System { index: 5 },
            back_style: BackStyle::Transparent,
            border_color: Color::System { index: 8 },
            border_style: DrawStyle::Solid,
            border_width: 1,
            draw_mode: DrawMode::CopyPen,
            fill_color: Color::RGB {
                red: 0,
                green: 0,
                blue: 0,
            },
            fill_style: DrawStyle::Transparent,
            height: 355,
            left: 30,
            shape: Shape::Rectangle,
            top: 200,
            visible: Visibility::Visible,
            width: 355,
        }
    }
}

impl From<Properties> for ShapeProperties {
    fn from(prop: Properties) -> Self {
        let mut shape_prop = ShapeProperties::default();

        shape_prop.back_color = prop.get_color("BackColor", shape_prop.back_color);
        shape_prop.back_style = prop.get_property("BackStyle", shape_prop.back_style);
        shape_prop.border_color = prop.get_color("BorderColor", shape_prop.border_color);
        shape_prop.border_style = prop.get_property("BorderStyle", shape_prop.border_style);
        shape_prop.border_width = prop.get_i32("BorderWidth", shape_prop.border_width);
        shape_prop.draw_mode = prop.get_property("DrawMode", shape_prop.draw_mode);
        shape_prop.fill_color = prop.get_color("FillColor", shape_prop.fill_color);
        shape_prop.fill_style = prop.get_property("FillStyle", shape_prop.fill_style);
        shape_prop.height = prop.get_i32("Height", shape_prop.height);
        shape_prop.left = prop.get_i32("Left", shape_prop.left);
        shape_prop.shape = prop.get_property("Shape", shape_prop.shape);
        shape_prop.top = prop.get_i32("Top", shape_prop.top);
        shape_prop.visible = prop.get_property("Visible", shape_prop.visible);
        shape_prop.width = prop.get_i32("Width", shape_prop.width);

        shape_prop
    }
}
