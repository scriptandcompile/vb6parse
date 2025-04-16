use crate::language::controls::{Appearance, DragMode, MousePointer, OLEDropMode, Style};
use crate::parsers::Properties;
use crate::VB6Color;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `CommandButton` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::CommandButton`](crate::language::controls::VB6ControlKind::CommandButton).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct CommandButtonProperties {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub cancel: bool,
    pub caption: BString,
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
    pub tool_tip_text: BString,
    pub top: i32,
    pub use_mask_color: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for CommandButtonProperties {
    fn default() -> Self {
        CommandButtonProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            cancel: false,
            caption: "".into(),
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
            tool_tip_text: "".into(),
            top: 30,
            use_mask_color: false,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for CommandButtonProperties {
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

impl<'a> From<Properties<'a>> for CommandButtonProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut command_button_prop = CommandButtonProperties::default();

        command_button_prop.appearance =
            prop.get_property(b"Appearance".into(), command_button_prop.appearance);
        command_button_prop.back_color =
            prop.get_color(b"BackColor".into(), command_button_prop.back_color);
        command_button_prop.cancel = prop.get_bool(b"Cancel".into(), command_button_prop.cancel);

        command_button_prop.caption = match prop.get("Caption".into()) {
            Some(caption) => caption.into(),
            None => command_button_prop.caption,
        };
        command_button_prop.causes_validation = prop.get_bool(
            b"CausesValidation".into(),
            command_button_prop.causes_validation,
        );
        command_button_prop.default = prop.get_bool(b"Default".into(), command_button_prop.default);

        // disabled_picture
        // down_picture
        // drag_icon

        command_button_prop.drag_mode =
            prop.get_property(b"DragMode".into(), command_button_prop.drag_mode);
        command_button_prop.enabled = prop.get_bool(b"Enabled".into(), command_button_prop.enabled);
        command_button_prop.height = prop.get_i32(b"Height".into(), command_button_prop.height);
        command_button_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), command_button_prop.help_context_id);
        command_button_prop.left = prop.get_i32(b"Left".into(), command_button_prop.left);
        command_button_prop.mask_color =
            prop.get_color(b"MaskColor".into(), command_button_prop.mask_color);

        // mouse_icon

        command_button_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), command_button_prop.mouse_pointer);
        command_button_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), command_button_prop.ole_drop_mode);

        // picture

        command_button_prop.right_to_left =
            prop.get_bool(b"RightToLeft".into(), command_button_prop.right_to_left);
        command_button_prop.style = prop.get_property(b"Style".into(), command_button_prop.style);
        command_button_prop.tab_index =
            prop.get_i32(b"TabIndex".into(), command_button_prop.tab_index);
        command_button_prop.tab_stop =
            prop.get_bool(b"TabStop".into(), command_button_prop.tab_stop);
        command_button_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        command_button_prop.top = prop.get_i32(b"Top".into(), command_button_prop.top);
        command_button_prop.use_mask_color =
            prop.get_bool(b"UseMaskColor".into(), command_button_prop.use_mask_color);
        command_button_prop.whats_this_help_id = prop.get_i32(
            b"WhatsThisHelp".into(),
            command_button_prop.whats_this_help_id,
        );
        command_button_prop.width = prop.get_i32(b"Width".into(), command_button_prop.width);

        command_button_prop
    }
}
