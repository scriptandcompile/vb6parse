use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{
    Appearance, BorderStyle, DragMode, MousePointer, OLEDragMode, OLEDropMode,
};

use crate::parsers::form::{build_bool_property, build_i32_property, build_property};

use bstr::BStr;

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `Image` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Image`](crate::language::controls::VB6ControlKind::Image).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ImageProperties<'a> {
    pub appearance: Appearance,
    pub border_style: BorderStyle,
    pub data_field: &'a BStr,
    pub data_format: &'a BStr,
    pub data_member: &'a BStr,
    pub data_source: &'a BStr,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub left: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub stretch: bool,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ImageProperties<'_> {
    fn default() -> Self {
        ImageProperties {
            appearance: Appearance::ThreeD,
            border_style: BorderStyle::None,
            data_field: BStr::new(""),
            data_format: BStr::new(""),
            data_member: BStr::new(""),
            data_source: BStr::new(""),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            height: 975,
            left: 1080,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            stretch: false,
            tool_tip_text: BStr::new(""),
            top: 960,
            visible: true,
            whats_this_help_id: 0,
            width: 615,
        }
    }
}

impl Serialize for ImageProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("ImageProperties", 21)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("border_style", &self.border_style)?;
        s.serialize_field("data_field", &self.data_field)?;
        s.serialize_field("data_format", &self.data_format)?;
        s.serialize_field("data_member", &self.data_member)?;
        s.serialize_field("data_source", &self.data_source)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("left", &self.left)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;

        let option_text = self.picture.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("picture", &option_text)?;
        s.serialize_field("stretch", &self.stretch)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl<'a> ImageProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut image_properties = ImageProperties::default();

        image_properties.appearance = build_property(properties, b"Appearance");
        image_properties.border_style = build_property(properties, b"BorderStyle");
        image_properties.data_field = properties
            .get(BStr::new("DataField"))
            .unwrap_or(&image_properties.data_field);
        image_properties.data_format = properties
            .get(BStr::new("DataFormat"))
            .unwrap_or(&image_properties.data_format);
        image_properties.data_member = properties
            .get(BStr::new("DataMember"))
            .unwrap_or(&image_properties.data_member);
        image_properties.data_source = properties
            .get(BStr::new("DataSource"))
            .unwrap_or(&image_properties.data_source);

        // DragIcon

        image_properties.drag_mode = build_property(properties, b"DragMode");
        image_properties.enabled =
            build_bool_property(properties, b"Enabled", image_properties.enabled);
        image_properties.height =
            build_i32_property(properties, b"Height", image_properties.height);
        image_properties.left = build_i32_property(properties, b"Left", image_properties.left);

        // MouseIcon

        image_properties.mouse_pointer = build_property(properties, b"MousePointer");
        image_properties.ole_drag_mode = build_property(properties, b"OLEDragMode");
        image_properties.ole_drop_mode = build_property(properties, b"OLEDropMode");

        // Picture

        image_properties.stretch =
            build_bool_property(properties, b"Stretch", image_properties.stretch);
        image_properties.tool_tip_text = properties
            .get(BStr::new("ToolTipText"))
            .unwrap_or(&image_properties.tool_tip_text);
        image_properties.top = build_i32_property(properties, b"Top", image_properties.top);
        image_properties.visible =
            build_bool_property(properties, b"Visible", image_properties.visible);
        image_properties.whats_this_help_id = build_i32_property(
            properties,
            b"WhatsThisHelpID",
            image_properties.whats_this_help_id,
        );
        image_properties.width = build_i32_property(properties, b"Width", image_properties.width);

        Ok(image_properties)
    }
}
