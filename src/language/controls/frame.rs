use crate::language::color::VB6Color;
use crate::language::controls::{
    Appearance, BorderStyle, ClipControls, DragMode, MousePointer, OLEDropMode, Visibility,
};
use crate::parsers::Properties;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `Frame` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Frame`](crate::language::controls::VB6ControlKind::Frame).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FrameProperties {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub caption: BString,
    pub clip_controls: ClipControls,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub tab_index: i32,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for FrameProperties {
    fn default() -> Self {
        FrameProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            caption: BString::from("Frame1"),
            clip_controls: ClipControls::True,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            left: 30,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: false,
            tab_index: 0,
            tool_tip_text: BString::from(""),
            top: 30,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for FrameProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("FrameProperties", 20)?;

        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("border_style", &self.border_style)?;
        s.serialize_field("caption", &self.caption)?;
        s.serialize_field("clip_controls", &self.clip_controls)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("left", &self.left)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl<'a> From<Properties<'a>> for FrameProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut frame_prop = FrameProperties::default();

        frame_prop.appearance = prop.get_property(b"Appearance".into(), frame_prop.appearance);
        frame_prop.back_color = prop.get_color(b"BackColor".into(), frame_prop.back_color);
        frame_prop.border_style = prop.get_property(b"BorderStyle".into(), frame_prop.border_style);
        frame_prop.caption = match prop.get(b"Caption".into()) {
            Some(caption) => caption.into(),
            None => frame_prop.caption,
        };
        frame_prop.clip_controls =
            prop.get_property(b"ClipControls".into(), frame_prop.clip_controls);

        // drag_icon

        frame_prop.drag_mode = prop.get_property(b"DragMode".into(), frame_prop.drag_mode);
        frame_prop.enabled = prop.get_bool(b"Enabled".into(), frame_prop.enabled);
        frame_prop.fore_color = prop.get_color(b"ForeColor".into(), frame_prop.fore_color);
        frame_prop.height = prop.get_i32(b"Height".into(), frame_prop.height);
        frame_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), frame_prop.help_context_id);
        frame_prop.left = prop.get_i32(b"Left".into(), frame_prop.left);

        // Implement mouse_icon

        frame_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), frame_prop.mouse_pointer);
        frame_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), frame_prop.ole_drop_mode);
        frame_prop.right_to_left = prop.get_bool(b"RightToLeft".into(), frame_prop.right_to_left);
        frame_prop.tab_index = prop.get_i32(b"TabIndex".into(), frame_prop.tab_index);
        frame_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => frame_prop.tool_tip_text,
        };
        frame_prop.top = prop.get_i32(b"Top".into(), frame_prop.top);
        frame_prop.visible = prop.get_property(b"Visible".into(), frame_prop.visible);
        frame_prop.whats_this_help_id =
            prop.get_i32(b"WhatsThisHelp".into(), frame_prop.whats_this_help_id);
        frame_prop.width = prop.get_i32(b"Width".into(), frame_prop.width);

        frame_prop
    }
}
