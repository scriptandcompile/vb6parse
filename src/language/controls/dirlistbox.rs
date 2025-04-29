use crate::language::controls::{
    Appearance, DragMode, MousePointer, OLEDragMode, OLEDropMode, Visibility,
};
use crate::parsers::Properties;
use crate::VB6Color;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `DirListBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::DirListBox`](crate::language::controls::VB6ControlKind::DirListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct DirListBoxProperties {
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
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for DirListBoxProperties {
    fn default() -> Self {
        DirListBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::System { index: 5 },
            causes_validation: true,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::System { index: 8 },
            height: 3195,
            help_context_id: 0,
            left: 720,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: "".into(),
            top: 720,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 975,
        }
    }
}

impl Serialize for DirListBoxProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("DirListBoxProperties", 20)?;
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
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
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

impl<'a> From<Properties<'a>> for DirListBoxProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut dir_list_box_prop = DirListBoxProperties::default();

        dir_list_box_prop.appearance =
            prop.get_property(b"Appearance".into(), dir_list_box_prop.appearance);
        dir_list_box_prop.back_color =
            prop.get_color(b"BackColor".into(), dir_list_box_prop.back_color);
        dir_list_box_prop.causes_validation = prop.get_bool(
            b"CausesValidation".into(),
            dir_list_box_prop.causes_validation,
        );

        // DragIcon

        dir_list_box_prop.drag_mode =
            prop.get_property(b"DragMode".into(), dir_list_box_prop.drag_mode);
        dir_list_box_prop.enabled = prop.get_bool(b"Enabled".into(), dir_list_box_prop.enabled);
        dir_list_box_prop.fore_color =
            prop.get_color(b"ForeColor".into(), dir_list_box_prop.fore_color);
        dir_list_box_prop.height = prop.get_i32(b"Height".into(), dir_list_box_prop.height);
        dir_list_box_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), dir_list_box_prop.help_context_id);
        dir_list_box_prop.left = prop.get_i32(b"Left".into(), dir_list_box_prop.left);
        dir_list_box_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), dir_list_box_prop.mouse_pointer);
        dir_list_box_prop.ole_drag_mode =
            prop.get_property(b"OLEDragMode".into(), dir_list_box_prop.ole_drag_mode);
        dir_list_box_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), dir_list_box_prop.ole_drop_mode);
        dir_list_box_prop.tab_index = prop.get_i32(b"TabIndex".into(), dir_list_box_prop.tab_index);
        dir_list_box_prop.tab_stop = prop.get_bool(b"TabStop".into(), dir_list_box_prop.tab_stop);
        dir_list_box_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        dir_list_box_prop.top = prop.get_i32(b"Top".into(), dir_list_box_prop.top);
        dir_list_box_prop.visible = prop.get_property(b"Visible".into(), dir_list_box_prop.visible);
        dir_list_box_prop.whats_this_help_id = prop.get_i32(
            b"WhatsThisHelpID".into(),
            dir_list_box_prop.whats_this_help_id,
        );
        dir_list_box_prop.width = prop.get_i32(b"Width".into(), dir_list_box_prop.width);

        dir_list_box_prop
    }
}
