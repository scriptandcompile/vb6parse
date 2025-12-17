//! Properties for a `DriveListBox` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::DriveListBox`](crate::language::controls::ControlKind::Drive
//! ListBox).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use crate::language::controls::{
    Activation, Appearance, CausesValidation, DragMode, MousePointer, OLEDropMode,
    ReferenceOrValue, TabStop, Visibility,
};
use crate::parsers::Properties;
use crate::Color;

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `DriveListBox` control.
///
/// This is used as an enum variant of
/// [`ControlKind::DriveListBox`](crate::language::controls::ControlKind::DriveListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct DriveListBoxProperties {
    /// The appearance of the DriveListBox.
    pub appearance: Appearance,
    /// The background color of the DriveListBox.
    pub back_color: Color,
    /// Whether the DriveListBox causes validation.
    pub causes_validation: CausesValidation,
    /// The drag icon of the DriveListBox.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// The drag mode of the DriveListBox.
    pub drag_mode: DragMode,
    /// Whether the DriveListBox is enabled.
    pub enabled: Activation,
    /// The foreground color of the DriveListBox.
    pub fore_color: Color,
    /// The height of the DriveListBox.
    pub height: i32,
    /// The help context ID of the DriveListBox.
    pub help_context_id: i32,
    /// The left position of the DriveListBox.
    pub left: i32,
    /// The mouse icon of the DriveListBox.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// The mouse pointer of the DriveListBox.
    pub mouse_pointer: MousePointer,
    /// The OLE drop mode of the DriveListBox.
    pub ole_drop_mode: OLEDropMode,
    /// The tab index of the DriveListBox.
    pub tab_index: i32,
    /// The tab stop of the DriveListBox.
    pub tab_stop: TabStop,
    /// The tool tip text of the DriveListBox.
    pub tool_tip_text: String,
    /// The top position of the DriveListBox.
    pub top: i32,
    /// The visibility of the DriveListBox.
    pub visible: Visibility,
    /// The What's This Help ID of the DriveListBox.
    pub whats_this_help_id: i32,
    /// The width of the DriveListBox.
    pub width: i32,
}

impl Default for DriveListBoxProperties {
    fn default() -> Self {
        DriveListBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: Color::System { index: 5 },
            causes_validation: CausesValidation::Yes,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            fore_color: Color::System { index: 8 },
            height: 319,
            help_context_id: 0,
            left: 480,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            tab_index: 0,
            tab_stop: TabStop::Included,
            tool_tip_text: "".into(),
            top: 960,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 1455,
        }
    }
}

impl Serialize for DriveListBoxProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("DriveListBoxProperties", 20)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;

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
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl From<Properties> for DriveListBoxProperties {
    fn from(prop: Properties) -> Self {
        let mut drive_list_box_prop = DriveListBoxProperties::default();

        drive_list_box_prop.appearance =
            prop.get_property("Appearance", drive_list_box_prop.appearance);
        drive_list_box_prop.back_color =
            prop.get_color("BackColor", drive_list_box_prop.back_color);
        drive_list_box_prop.causes_validation =
            prop.get_property("CausesValidation", drive_list_box_prop.causes_validation);

        // TODO: Implement DragIcon parsing
        // DragIcon

        drive_list_box_prop.drag_mode =
            prop.get_property("DragMode", drive_list_box_prop.drag_mode);
        drive_list_box_prop.enabled = prop.get_property("Enabled", drive_list_box_prop.enabled);
        drive_list_box_prop.fore_color =
            prop.get_color("ForeColor", drive_list_box_prop.fore_color);
        drive_list_box_prop.height = prop.get_i32("Height", drive_list_box_prop.height);
        drive_list_box_prop.help_context_id =
            prop.get_i32("HelpContextID", drive_list_box_prop.help_context_id);
        drive_list_box_prop.left = prop.get_i32("Left", drive_list_box_prop.left);
        drive_list_box_prop.mouse_pointer =
            prop.get_property("MousePointer", drive_list_box_prop.mouse_pointer);
        drive_list_box_prop.ole_drop_mode =
            prop.get_property("OLEDropMode", drive_list_box_prop.ole_drop_mode);
        drive_list_box_prop.tab_index = prop.get_i32("TabIndex", drive_list_box_prop.tab_index);
        drive_list_box_prop.tab_stop = prop.get_property("TabStop", drive_list_box_prop.tab_stop);
        drive_list_box_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => drive_list_box_prop.tool_tip_text,
        };
        drive_list_box_prop.top = prop.get_i32("Top", drive_list_box_prop.top);
        drive_list_box_prop.visible = prop.get_property("Visible", drive_list_box_prop.visible);
        drive_list_box_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", drive_list_box_prop.whats_this_help_id);
        drive_list_box_prop.width = prop.get_i32("Width", drive_list_box_prop.width);

        drive_list_box_prop
    }
}
