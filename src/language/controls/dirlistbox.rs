//! Properties for a `DirListBox` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::DirListBox`](crate::language::controls::ControlKind::DirListBox).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!
use crate::language::controls::{
    Activation, Appearance, CausesValidation, DragMode, MousePointer, OLEDragMode, OLEDropMode,
    ReferenceOrValue, TabStop, Visibility,
};
use crate::parsers::Properties;
use crate::Color;

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `DirListBox` control.
///
/// This is used as an enum variant of
/// [`ControlKind::DirListBox`](crate::language::controls::ControlKind::DirListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct DirListBoxProperties {
    /// The appearance of the DirListBox.
    pub appearance: Appearance,
    /// The background color of the DirListBox.
    pub back_color: Color,
    /// Whether the DirListBox causes validation.
    pub causes_validation: CausesValidation,
    /// The drag icon of the DirListBox.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// The drag mode of the DirListBox.
    pub drag_mode: DragMode,
    /// Whether the DirListBox is enabled.
    pub enabled: Activation,
    /// The foreground color of the DirListBox.
    pub fore_color: Color,
    /// The height of the DirListBox.
    pub height: i32,
    /// The help context ID of the DirListBox.
    pub help_context_id: i32,
    /// The left position of the DirListBox.
    pub left: i32,
    /// The mouse icon of the DirListBox.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// The mouse pointer of the DirListBox.
    pub mouse_pointer: MousePointer,
    /// The OLE drag mode of the DirListBox.
    pub ole_drag_mode: OLEDragMode,
    /// The OLE drop mode of the DirListBox.
    pub ole_drop_mode: OLEDropMode,
    /// The tab index of the DirListBox.
    pub tab_index: i32,
    /// The tab stop of the DirListBox.
    pub tab_stop: TabStop,
    /// The tool tip text of the DirListBox.
    pub tool_tip_text: String,
    /// The top position of the DirListBox.
    pub top: i32,
    /// Whether the DirListBox is visible.
    pub visible: Visibility,
    /// The "What's This" help ID of the DirListBox.
    pub whats_this_help_id: i32,
    /// The width of the DirListBox.
    pub width: i32,
}

impl Default for DirListBoxProperties {
    fn default() -> Self {
        DirListBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: Color::System { index: 5 },
            causes_validation: CausesValidation::Yes,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            fore_color: Color::System { index: 8 },
            height: 3195,
            help_context_id: 0,
            left: 720,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            tab_index: 0,
            tab_stop: TabStop::Included,
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

impl From<Properties> for DirListBoxProperties {
    fn from(prop: Properties) -> Self {
        let mut dir_list_box_prop = DirListBoxProperties::default();

        dir_list_box_prop.appearance =
            prop.get_property("Appearance", dir_list_box_prop.appearance);
        dir_list_box_prop.back_color = prop.get_color("BackColor", dir_list_box_prop.back_color);
        dir_list_box_prop.causes_validation =
            prop.get_property("CausesValidation", dir_list_box_prop.causes_validation);

        // TODO: Implement DragIcon parsing
        // DragIcon

        dir_list_box_prop.drag_mode = prop.get_property("DragMode", dir_list_box_prop.drag_mode);
        dir_list_box_prop.enabled = prop.get_property("Enabled", dir_list_box_prop.enabled);
        dir_list_box_prop.fore_color = prop.get_color("ForeColor", dir_list_box_prop.fore_color);
        dir_list_box_prop.height = prop.get_i32("Height", dir_list_box_prop.height);
        dir_list_box_prop.help_context_id =
            prop.get_i32("HelpContextID", dir_list_box_prop.help_context_id);
        dir_list_box_prop.left = prop.get_i32("Left", dir_list_box_prop.left);
        dir_list_box_prop.mouse_pointer =
            prop.get_property("MousePointer", dir_list_box_prop.mouse_pointer);
        dir_list_box_prop.ole_drag_mode =
            prop.get_property("OLEDragMode", dir_list_box_prop.ole_drag_mode);
        dir_list_box_prop.ole_drop_mode =
            prop.get_property("OLEDropMode", dir_list_box_prop.ole_drop_mode);
        dir_list_box_prop.tab_index = prop.get_i32("TabIndex", dir_list_box_prop.tab_index);
        dir_list_box_prop.tab_stop = prop.get_property("TabStop", dir_list_box_prop.tab_stop);
        dir_list_box_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        dir_list_box_prop.top = prop.get_i32("Top", dir_list_box_prop.top);
        dir_list_box_prop.visible = prop.get_property("Visible", dir_list_box_prop.visible);
        dir_list_box_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", dir_list_box_prop.whats_this_help_id);
        dir_list_box_prop.width = prop.get_i32("Width", dir_list_box_prop.width);

        dir_list_box_prop
    }
}
