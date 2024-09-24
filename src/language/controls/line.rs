use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::color::VB6Color;
use crate::language::controls::{DrawMode, DrawStyle};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};

use bstr::BStr;
use serde::Serialize;

/// Properties for a Line control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Line`](crate::language::controls::VB6ControlKind::Line).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
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

impl LineProperties {
    pub fn construct_control(properties: &HashMap<&BStr, &BStr>) -> Result<Self, VB6ErrorKind> {
        let mut line_properties = LineProperties::default();

        line_properties.border_color = build_color_property(
            properties,
            BStr::new("BorderColor"),
            line_properties.border_color,
        );
        line_properties.border_style =
            build_property::<DrawStyle>(properties, BStr::new("BorderStyle"));
        line_properties.border_width = build_i32_property(
            properties,
            BStr::new("BorderWidth"),
            line_properties.border_width,
        );
        line_properties.draw_mode = build_property::<DrawMode>(properties, BStr::new("DrawMode"));
        line_properties.visible =
            build_bool_property(properties, BStr::new("Visible"), line_properties.visible);
        line_properties.x1 = build_i32_property(properties, BStr::new("X1"), line_properties.x1);
        line_properties.y1 = build_i32_property(properties, BStr::new("Y1"), line_properties.y1);
        line_properties.x2 = build_i32_property(properties, BStr::new("X2"), line_properties.x2);
        line_properties.y2 = build_i32_property(properties, BStr::new("Y2"), line_properties.y2);

        Ok(line_properties)
    }
}
