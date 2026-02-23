//! Properties for a `Frame` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::Frame`](crate::language::controls::ControlKind::Frame).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use crate::{
    files::common::Properties,
    language::{
        color::{Color, VB_BUTTON_FACE, VB_BUTTON_TEXT},
        controls::{
            Activation, Appearance, BorderStyle, ClipControls, DragMode, Font, MousePointer,
            OLEDropMode, ReferenceOrValue, TextDirection, Visibility,
        },
    },
};

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `Frame` control.
///
/// This is used as an enum variant of
/// [`ControlKind::Frame`](crate::language::controls::ControlKind::Frame).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FrameProperties {
    /// Appearance of the frame.
    pub appearance: Appearance,
    /// Background color of the frame.
    pub back_color: Color,
    /// Border style of the frame.
    pub border_style: BorderStyle,
    /// Caption of the frame.
    pub caption: String,
    /// Clip controls setting of the frame.
    pub clip_controls: ClipControls,
    /// Drag icon of the frame.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Drag mode of the frame.
    pub drag_mode: DragMode,
    /// Enabled state of the frame.
    pub enabled: Activation,
    /// The font style for the form.
    pub font: Option<Font>,
    /// Foreground color of the frame.
    pub fore_color: Color,
    /// Height of the frame.
    pub height: i32,
    /// Help context ID of the frame.
    pub help_context_id: i32,
    /// Left position of the frame.
    pub left: i32,
    /// Mouse icon of the frame.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Mouse pointer of the frame.
    pub mouse_pointer: MousePointer,
    /// OLE drop mode of the frame.
    pub ole_drop_mode: OLEDropMode,
    /// Text direction of the frame.
    pub right_to_left: TextDirection,
    /// Tab index of the frame.
    pub tab_index: i32,
    /// Tool tip text of the frame.
    pub tool_tip_text: String,
    /// Top position of the frame.
    pub top: i32,
    /// Visibility of the frame.
    pub visible: Visibility,
    /// "What's This?" help ID of the frame.
    pub whats_this_help_id: i32,
    /// Width of the frame.
    pub width: i32,
}

impl Default for FrameProperties {
    fn default() -> Self {
        FrameProperties {
            appearance: Appearance::ThreeD,
            back_color: VB_BUTTON_FACE,
            border_style: BorderStyle::FixedSingle,
            caption: String::from("Frame1"),
            clip_controls: ClipControls::Clipped,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            font: Some(Font::default()),
            fore_color: VB_BUTTON_TEXT,
            height: 30,
            help_context_id: 0,
            left: 30,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: TextDirection::LeftToRight,
            tab_index: 0,
            tool_tip_text: String::new(),
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

impl From<Properties> for FrameProperties {
    fn from(prop: Properties) -> Self {
        let mut frame_prop = FrameProperties::default();

        frame_prop.appearance = prop.get_property("Appearance", frame_prop.appearance);
        frame_prop.back_color = prop.get_color("BackColor", frame_prop.back_color);
        frame_prop.border_style = prop.get_property("BorderStyle", frame_prop.border_style);
        frame_prop.caption = match prop.get("Caption") {
            Some(caption) => caption.into(),
            None => frame_prop.caption,
        };
        frame_prop.clip_controls = prop.get_property("ClipControls", frame_prop.clip_controls);

        // TODO: process drag_icon
        // drag_icon

        frame_prop.drag_mode = prop.get_property("DragMode", frame_prop.drag_mode);
        frame_prop.enabled = prop.get_property("Enabled", frame_prop.enabled);
        frame_prop.fore_color = prop.get_color("ForeColor", frame_prop.fore_color);
        frame_prop.height = prop.get_i32("Height", frame_prop.height);
        frame_prop.help_context_id = prop.get_i32("HelpContextID", frame_prop.help_context_id);
        frame_prop.left = prop.get_i32("Left", frame_prop.left);

        // TODO: process mouse_icon
        // Implement mouse_icon

        frame_prop.mouse_pointer = prop.get_property("MousePointer", frame_prop.mouse_pointer);
        frame_prop.ole_drop_mode = prop.get_property("OLEDropMode", frame_prop.ole_drop_mode);
        frame_prop.right_to_left = prop.get_property("RightToLeft", frame_prop.right_to_left);
        frame_prop.tab_index = prop.get_i32("TabIndex", frame_prop.tab_index);
        frame_prop.tool_tip_text = match prop.get("ToolTipText") {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => frame_prop.tool_tip_text,
        };
        frame_prop.top = prop.get_i32("Top", frame_prop.top);
        frame_prop.visible = prop.get_property("Visible", frame_prop.visible);
        frame_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelp", frame_prop.whats_this_help_id);
        frame_prop.width = prop.get_i32("Width", frame_prop.width);

        frame_prop
    }
}
