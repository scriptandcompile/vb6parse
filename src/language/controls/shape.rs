use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::color::VB6Color;
use crate::language::controls::{BackStyle, DrawMode, DrawStyle};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};

use bstr::BStr;
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

impl ShapeProperties {
    pub fn construct_control(properties: &HashMap<&BStr, &BStr>) -> Result<Self, VB6ErrorKind> {
        let mut shape_properties = ShapeProperties::default();

        shape_properties.back_color =
            build_color_property(properties, b"BackColor", shape_properties.back_color);
        shape_properties.back_style = build_property(properties, b"BackStyle");
        shape_properties.border_color =
            build_color_property(properties, b"BorderColor", shape_properties.border_color);
        shape_properties.border_style = build_property(properties, b"BorderStyle");
        shape_properties.border_width =
            build_i32_property(properties, b"BorderWidth", shape_properties.border_width);
        shape_properties.draw_mode = build_property(properties, b"DrawMode");
        shape_properties.fill_color =
            build_color_property(properties, b"FillColor", shape_properties.fill_color);
        shape_properties.fill_style = build_property(properties, b"FillStyle");
        shape_properties.height =
            build_i32_property(properties, b"Height", shape_properties.height);
        shape_properties.left = build_i32_property(properties, b"Left", shape_properties.left);
        shape_properties.shape = build_property(properties, b"Shape");
        shape_properties.top = build_i32_property(properties, b"Top", shape_properties.top);
        shape_properties.visible =
            build_bool_property(properties, b"Visible", shape_properties.visible);
        shape_properties.width = build_i32_property(properties, b"Width", shape_properties.width);

        Ok(shape_properties)
    }
}
