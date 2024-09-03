use crate::language::controls::{
    Appearance, DragMode, MousePointer, MultiSelect, OLEDragMode, OLEDropMode,
};
use crate::VB6Color;

use image::DynamicImage;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub enum ListBoxStyle {
    Standard = 0,
    Checkbox = 1,
}

/// Properties for a ListBox control. This is used as an enum variant of
/// [VB6ControlKind::ListBox](crate::language::controls::VB6ControlKind::ListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [VB6Control](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ListBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    pub columns: u32,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub integral_height: bool,
    // pub item_data: Vec<&'a str>,
    pub left: i32,
    // pub list: Vec<&'a str>,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub multi_select: MultiSelect,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub sorted: bool,
    pub style: ListBoxStyle,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ListBoxProperties<'_> {
    fn default() -> Self {
        ListBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            causes_validation: true,
            columns: 0,
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            integral_height: true,
            left: 30,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            multi_select: MultiSelect::None,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::None,
            right_to_left: false,
            sorted: false,
            style: ListBoxStyle::Standard,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: "",
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for ListBoxProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("ListBoxProperties", 26)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;
        s.serialize_field("columns", &self.columns)?;
        s.serialize_field("data_field", &self.data_field)?;
        s.serialize_field("data_format", &self.data_format)?;
        s.serialize_field("data_member", &self.data_member)?;
        s.serialize_field("data_source", &self.data_source)?;

        let option_text = match &self.drag_icon {
            Some(_) => "Some(DynamicImage)",
            None => "None",
        };

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("integral_height", &self.integral_height)?;
        s.serialize_field("left", &self.left)?;

        let option_text = match &self.mouse_icon {
            Some(_) => "Some(DynamicImage)",
            None => "None",
        };

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("multi_select", &self.multi_select)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("sorted", &self.sorted)?;
        s.serialize_field("style", &self.style)?;
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
