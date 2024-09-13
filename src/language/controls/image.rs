use crate::language::controls::{
    Appearance, BorderStyle, DragMode, MousePointer, OLEDragMode, OLEDropMode,
};

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
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
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
    pub tool_tip_text: &'a str,
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
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
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
            tool_tip_text: "",
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
