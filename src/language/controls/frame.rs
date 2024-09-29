use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::color::VB6Color;
use crate::language::controls::{
    Appearance, BorderStyle, ClipControls, DragMode, MousePointer, OLEDropMode,
};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property, VB6PropertyGroup,
};

use bstr::BStr;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `Frame` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Frame`](crate::language::controls::VB6ControlKind::Frame).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FrameProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub caption: &'a BStr,
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
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for FrameProperties<'_> {
    fn default() -> Self {
        FrameProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            caption: BStr::new("Frame1"),
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
            tool_tip_text: BStr::new(""),
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for FrameProperties<'_> {
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

impl<'a> FrameProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
        _property_groups: &[VB6PropertyGroup<'a>],
    ) -> Result<Self, VB6ErrorKind> {
        // TODO: We are not correctly handling property assignment for each control.

        let mut frame_properties = FrameProperties::default();

        frame_properties.appearance = build_property(properties, b"Appearance");
        frame_properties.back_color =
            build_color_property(properties, b"BackColor", frame_properties.back_color);
        frame_properties.border_style = build_property(properties, b"BorderStyle");
        frame_properties.caption = properties
            .get(BStr::new("Caption"))
            .unwrap_or(&frame_properties.caption);
        frame_properties.clip_controls = build_property(properties, b"ClipControls");

        // drag_icon

        frame_properties.drag_mode = build_property(properties, b"DragMode");
        frame_properties.enabled =
            build_bool_property(properties, b"Enabled", frame_properties.enabled);
        frame_properties.fore_color =
            build_color_property(properties, b"ForeColor", frame_properties.fore_color);
        frame_properties.height =
            build_i32_property(properties, b"Height", frame_properties.height);
        frame_properties.help_context_id = build_i32_property(
            properties,
            b"HelpContextID",
            frame_properties.help_context_id,
        );
        frame_properties.left = build_i32_property(properties, b"Left", frame_properties.left);

        // Implement mouse_icon

        frame_properties.mouse_pointer = build_property(properties, b"MousePointer");
        frame_properties.ole_drop_mode = build_property(properties, b"OLEDropMode");
        frame_properties.right_to_left =
            build_bool_property(properties, b"RightToLeft", frame_properties.right_to_left);
        frame_properties.tab_index =
            build_i32_property(properties, b"TabIndex", frame_properties.tab_index);
        frame_properties.tool_tip_text = properties
            .get(BStr::new("ToolTipText"))
            .unwrap_or(&frame_properties.tool_tip_text);
        frame_properties.top = build_i32_property(properties, b"Top", frame_properties.top);
        frame_properties.visible =
            build_bool_property(properties, b"Visible", frame_properties.visible);
        frame_properties.whats_this_help_id = build_i32_property(
            properties,
            b"WhatsThisHelp",
            frame_properties.whats_this_help_id,
        );
        frame_properties.width = build_i32_property(properties, b"Width", frame_properties.width);

        Ok(frame_properties)
    }
}
