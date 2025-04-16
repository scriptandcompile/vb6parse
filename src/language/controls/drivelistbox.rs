use crate::language::controls::{Appearance, DragMode, MousePointer, OLEDropMode};
use crate::parsers::Properties;
use crate::VB6Color;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `DriveListBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::DriveListBox`](crate::language::controls::VB6ControlKind::DriveListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct DriveListBoxProperties {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
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
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for DriveListBoxProperties {
    fn default() -> Self {
        DriveListBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::System { index: 5 },
            causes_validation: true,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::System { index: 8 },
            height: 319,
            help_context_id: 0,
            left: 480,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: "".into(),
            top: 960,
            visible: true,
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

impl<'a> From<Properties<'a>> for DriveListBoxProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut drive_list_box_prop = DriveListBoxProperties::default();

        drive_list_box_prop.appearance =
            prop.get_property(b"Appearance".into(), drive_list_box_prop.appearance);
        drive_list_box_prop.back_color =
            prop.get_color(b"BackColor".into(), drive_list_box_prop.back_color);
        drive_list_box_prop.causes_validation = prop.get_bool(
            b"CausesValidation".into(),
            drive_list_box_prop.causes_validation,
        );

        // DragIcon

        drive_list_box_prop.drag_mode =
            prop.get_property(b"DragMode".into(), drive_list_box_prop.drag_mode);
        drive_list_box_prop.enabled = prop.get_bool(b"Enabled".into(), drive_list_box_prop.enabled);
        drive_list_box_prop.fore_color =
            prop.get_color(b"ForeColor".into(), drive_list_box_prop.fore_color);
        drive_list_box_prop.height = prop.get_i32(b"Height".into(), drive_list_box_prop.height);
        drive_list_box_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), drive_list_box_prop.help_context_id);
        drive_list_box_prop.left = prop.get_i32(b"Left".into(), drive_list_box_prop.left);
        drive_list_box_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), drive_list_box_prop.mouse_pointer);
        drive_list_box_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), drive_list_box_prop.ole_drop_mode);
        drive_list_box_prop.tab_index =
            prop.get_i32(b"TabIndex".into(), drive_list_box_prop.tab_index);
        drive_list_box_prop.tab_stop =
            prop.get_bool(b"TabStop".into(), drive_list_box_prop.tab_stop);
        drive_list_box_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => drive_list_box_prop.tool_tip_text,
        };
        drive_list_box_prop.top = prop.get_i32(b"Top".into(), drive_list_box_prop.top);
        drive_list_box_prop.visible = prop.get_bool(b"Visible".into(), drive_list_box_prop.visible);
        drive_list_box_prop.whats_this_help_id = prop.get_i32(
            b"WhatsThisHelpID".into(),
            drive_list_box_prop.whats_this_help_id,
        );
        drive_list_box_prop.width = prop.get_i32(b"Width".into(), drive_list_box_prop.width);

        drive_list_box_prop
    }
}
