use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{Appearance, DragMode, MousePointer, OLEDropMode, Style};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};
use crate::VB6Color;

use bstr::BStr;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `CommandButton` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::CommandButton`](crate::language::controls::VB6ControlKind::CommandButton).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct CommandButtonProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub cancel: bool,
    pub caption: &'a BStr,
    pub causes_validation: bool,
    pub default: bool,
    pub disabled_picture: Option<DynamicImage>,
    pub down_picture: Option<DynamicImage>,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub mask_color: VB6Color,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub right_to_left: bool,
    pub style: Style,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub use_mask_color: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for CommandButtonProperties<'_> {
    fn default() -> Self {
        CommandButtonProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            cancel: false,
            caption: BStr::new("Command1"),
            causes_validation: true,
            default: false,
            disabled_picture: None,
            down_picture: None,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            height: 30,
            help_context_id: 0,
            left: 30,
            mask_color: VB6Color::from_hex("&H00C0C0C0&").unwrap(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            right_to_left: false,
            style: Style::Standard,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: BStr::new(""),
            top: 30,
            use_mask_color: false,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for CommandButtonProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::ser::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("CommandButtonProperties", 24)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("cancel", &self.cancel)?;
        s.serialize_field("caption", &self.caption)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;
        s.serialize_field("default", &self.default)?;

        let option_text = self.disabled_picture.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("disabled_picture", &option_text)?;

        let option_text = self.down_picture.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("down_picture", &option_text)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("left", &self.left)?;
        s.serialize_field("mask_color", &self.mask_color)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;

        let option_text = self.picture.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("picture", &option_text)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("style", &self.style)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("use_mask_color", &self.use_mask_color)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl<'a> CommandButtonProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut command_button_properties = CommandButtonProperties::default();

        command_button_properties.appearance = build_property(properties, b"Appearance");
        command_button_properties.back_color = build_color_property(
            properties,
            b"BackColor",
            command_button_properties.back_color,
        );
        command_button_properties.cancel =
            build_bool_property(properties, b"Cancel", command_button_properties.cancel);

        command_button_properties.caption = properties
            .get(BStr::new("Caption"))
            .unwrap_or(&command_button_properties.caption);
        command_button_properties.causes_validation = build_bool_property(
            properties,
            b"CausesValidation",
            command_button_properties.causes_validation,
        );
        command_button_properties.default =
            build_bool_property(properties, b"Default", command_button_properties.default);

        // disabled_picture
        // down_picture
        // drag_icon

        command_button_properties.drag_mode = build_property(properties, b"DragMode");
        command_button_properties.enabled =
            build_bool_property(properties, b"Enabled", command_button_properties.enabled);
        command_button_properties.height =
            build_i32_property(properties, b"Height", command_button_properties.height);
        command_button_properties.help_context_id = build_i32_property(
            properties,
            b"HelpContextID",
            command_button_properties.help_context_id,
        );
        command_button_properties.left =
            build_i32_property(properties, b"Left", command_button_properties.left);
        command_button_properties.mask_color = build_color_property(
            properties,
            b"MaskColor",
            command_button_properties.mask_color,
        );

        // mouse_icon

        command_button_properties.mouse_pointer = build_property(properties, b"MousePointer");
        command_button_properties.ole_drop_mode = build_property(properties, b"OLEDropMode");

        // picture

        command_button_properties.right_to_left = build_bool_property(
            properties,
            b"RightToLeft",
            command_button_properties.right_to_left,
        );
        command_button_properties.style = build_property(properties, b"Style");
        command_button_properties.tab_index =
            build_i32_property(properties, b"TabIndex", command_button_properties.tab_index);
        command_button_properties.tab_stop =
            build_bool_property(properties, b"TabStop", command_button_properties.tab_stop);
        command_button_properties.tool_tip_text = properties
            .get(BStr::new("ToolTipText"))
            .unwrap_or(&command_button_properties.tool_tip_text);
        command_button_properties.top =
            build_i32_property(properties, b"Top", command_button_properties.top);
        command_button_properties.use_mask_color = build_bool_property(
            properties,
            b"UseMaskColor",
            command_button_properties.use_mask_color,
        );
        command_button_properties.whats_this_help_id = build_i32_property(
            properties,
            b"WhatsThisHelp",
            command_button_properties.whats_this_help_id,
        );
        command_button_properties.width =
            build_i32_property(properties, b"Width", command_button_properties.width);

        Ok(command_button_properties)
    }
}
