//! Properties for a `CommandButton` control.
//!
//! This module defines the `CommandButtonProperties` struct, which encapsulates
//! the various properties associated with a `CommandButton` control in a GUI
//! application. It includes default values, serialization logic, and conversion
//! from a generic `Properties` struct.
//!
//! The properties covered include appearance, colors, captions, validation behavior,
//! images, dimensions, and other control-specific settings.
//!
//! This struct is intended to be used as part of a larger control framework,
//! specifically as a variant of the `ControlKind::CommandButton` enum.
//!
//! See [`ControlKind::CommandButton`](crate::language::controls::ControlKind::CommandButton)
//! for usage.
//!
//! # References
//! - [Microsoft Docs: CommandButton Control](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa240834(v=vs.60))
//!

use crate::{
    language::controls::{
        Activation, Appearance, CausesValidation, DragMode, MousePointer, OLEDropMode,
        ReferenceOrValue, Style, TabStop, TextDirection, UseMaskColor,
    },
    parsers::Properties,
    Color, VB_BUTTON_FACE,
};

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `CommandButton` control.
///
/// This is used as an enum variant of
/// [`ControlKind::CommandButton`](crate::language::controls::ControlKind::CommandButton).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct CommandButtonProperties {
    /// The appearance of the command button.
    pub appearance: Appearance,
    /// The background color of the command button.
    pub back_color: Color,
    /// Indicates if the button is a cancel button.
    pub cancel: bool,
    /// The caption text of the command button.
    pub caption: String,
    /// Indicates if the button causes validation.
    pub causes_validation: CausesValidation,
    /// Indicates if the button is the default button.
    pub default: bool,
    /// The picture displayed when the button is disabled.
    pub disabled_picture: Option<ReferenceOrValue<DynamicImage>>,
    /// The picture displayed when the button is pressed down.
    pub down_picture: Option<ReferenceOrValue<DynamicImage>>,
    /// The icon used during drag operations.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// The drag mode of the command button.
    pub drag_mode: DragMode,
    /// Indicates if the button is enabled.
    pub enabled: Activation,
    /// The height of the command button.
    pub height: i32,
    /// The help context ID of the command button.
    pub help_context_id: i32,
    /// The left position of the command button.
    pub left: i32,
    /// The mask color of the command button.
    pub mask_color: Color,
    /// The mouse icon of the command button.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// The mouse pointer type of the command button.
    pub mouse_pointer: MousePointer,
    /// The OLE drop mode of the command button.
    pub ole_drop_mode: OLEDropMode,
    /// The picture displayed on the command button.
    pub picture: Option<ReferenceOrValue<DynamicImage>>,
    /// The text direction of the command button.
    pub right_to_left: TextDirection,
    /// The style of the command button.
    pub style: Style,
    /// The tab index of the command button.
    pub tab_index: i32,
    /// The tab stop behavior of the command button.
    pub tab_stop: TabStop,
    /// The tool tip text of the command button.
    pub tool_tip_text: String,
    /// The top position of the command button.
    pub top: i32,
    /// Indicates if the mask color is used.
    pub use_mask_color: UseMaskColor,
    /// The "What's This?" help ID of the command button.
    pub whats_this_help_id: i32,
    /// The width of the command button.
    pub width: i32,
}

impl Default for CommandButtonProperties {
    fn default() -> Self {
        CommandButtonProperties {
            appearance: Appearance::ThreeD,
            back_color: VB_BUTTON_FACE,
            cancel: false,
            caption: String::new(),
            causes_validation: CausesValidation::Yes,
            default: false,
            disabled_picture: None,
            down_picture: None,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            height: 30,
            help_context_id: 0,
            left: 30,
            mask_color: Color::new(0xC0, 0xC0, 0xC0),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            right_to_left: TextDirection::LeftToRight,
            style: Style::Standard,
            tab_index: 0,
            tab_stop: TabStop::Included,
            tool_tip_text: String::new(),
            top: 30,
            use_mask_color: UseMaskColor::DoNotUseMaskColor,
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

impl From<Properties> for CommandButtonProperties {
    fn from(prop: Properties) -> Self {
        let mut command_button_prop = CommandButtonProperties::default();

        command_button_prop.appearance =
            prop.get_property("Appearance", command_button_prop.appearance);
        command_button_prop.back_color =
            prop.get_color("BackColor", command_button_prop.back_color);
        command_button_prop.cancel = prop.get_bool("Cancel", command_button_prop.cancel);

        command_button_prop.caption = match prop.get("Caption") {
            Some(caption) => caption.into(),
            None => command_button_prop.caption,
        };
        command_button_prop.causes_validation =
            prop.get_property("CausesValidation", command_button_prop.causes_validation);
        command_button_prop.default = prop.get_bool("Default", command_button_prop.default);

        // disabled_picture
        // down_picture
        // drag_icon

        command_button_prop.drag_mode =
            prop.get_property("DragMode", command_button_prop.drag_mode);
        command_button_prop.enabled = prop.get_property("Enabled", command_button_prop.enabled);
        command_button_prop.height = prop.get_i32("Height", command_button_prop.height);
        command_button_prop.help_context_id =
            prop.get_i32("HelpContextID", command_button_prop.help_context_id);
        command_button_prop.left = prop.get_i32("Left", command_button_prop.left);
        command_button_prop.mask_color =
            prop.get_color("MaskColor", command_button_prop.mask_color);

        // mouse_icon

        command_button_prop.mouse_pointer =
            prop.get_property("MousePointer", command_button_prop.mouse_pointer);
        command_button_prop.ole_drop_mode =
            prop.get_property("OLEDropMode", command_button_prop.ole_drop_mode);

        // picture

        command_button_prop.right_to_left =
            prop.get_property("RightToLeft", command_button_prop.right_to_left);
        command_button_prop.style = prop.get_property("Style", command_button_prop.style);
        command_button_prop.tab_index = prop.get_i32("TabIndex", command_button_prop.tab_index);
        command_button_prop.tab_stop = prop.get_property("TabStop", command_button_prop.tab_stop);
        command_button_prop.tool_tip_text = match prop.get("ToolTipText") {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => String::new(),
        };
        command_button_prop.top = prop.get_i32("Top", command_button_prop.top);
        command_button_prop.use_mask_color =
            prop.get_property("UseMaskColor", command_button_prop.use_mask_color);
        command_button_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelp", command_button_prop.whats_this_help_id);
        command_button_prop.width = prop.get_i32("Width", command_button_prop.width);

        command_button_prop
    }
}
