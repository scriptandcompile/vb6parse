use crate::language::controls::{
    Activation, Appearance, BorderStyle, DragMode, MousePointer, OLEDragMode, OLEDropMode,
    Visibility,
};

use crate::parsers::Properties;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `Image` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Image`](crate::language::controls::VB6ControlKind::Image).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ImageProperties {
    pub appearance: Appearance,
    pub border_style: BorderStyle,
    pub data_field: BString,
    pub data_format: BString,
    pub data_member: BString,
    pub data_source: BString,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: Activation,
    pub height: i32,
    pub left: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub stretch: bool,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ImageProperties {
    fn default() -> Self {
        ImageProperties {
            appearance: Appearance::ThreeD,
            border_style: BorderStyle::None,
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            height: 975,
            left: 1080,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            stretch: false,
            tool_tip_text: "".into(),
            top: 960,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 615,
        }
    }
}

impl Serialize for ImageProperties {
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

impl<'a> From<Properties<'a>> for ImageProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut image_prop = ImageProperties::default();

        image_prop.appearance = prop.get_property(b"Appearance".into(), image_prop.appearance);
        image_prop.border_style = prop.get_property(b"BorderStyle".into(), image_prop.border_style);
        image_prop.data_field = match prop.get(b"DataField".into()) {
            Some(data_field) => data_field.into(),
            None => image_prop.data_field,
        };
        image_prop.data_format = match prop.get(b"DataFormat".into()) {
            Some(data_format) => data_format.into(),
            None => image_prop.data_format,
        };
        image_prop.data_member = match prop.get(b"DataMember".into()) {
            Some(data_member) => data_member.into(),
            None => image_prop.data_member,
        };
        image_prop.data_source = match prop.get(b"DataSource".into()) {
            Some(data_source) => data_source.into(),
            None => image_prop.data_source,
        };

        // DragIcon

        image_prop.drag_mode = prop.get_property(b"DragMode".into(), image_prop.drag_mode);
        image_prop.enabled = prop.get_property(b"Enabled".into(), image_prop.enabled);
        image_prop.height = prop.get_i32(b"Height".into(), image_prop.height);
        image_prop.left = prop.get_i32(b"Left".into(), image_prop.left);

        // MouseIcon

        image_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), image_prop.mouse_pointer);
        image_prop.ole_drag_mode =
            prop.get_property(b"OLEDragMode".into(), image_prop.ole_drag_mode);
        image_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), image_prop.ole_drop_mode);

        // Picture

        image_prop.stretch = prop.get_bool(b"Stretch".into(), image_prop.stretch);
        image_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        image_prop.top = prop.get_i32(b"Top".into(), image_prop.top);
        image_prop.visible = prop.get_property(b"Visible".into(), image_prop.visible);
        image_prop.whats_this_help_id =
            prop.get_i32(b"WhatsThisHelpID".into(), image_prop.whats_this_help_id);
        image_prop.width = prop.get_i32(b"Width".into(), image_prop.width);

        image_prop
    }
}
