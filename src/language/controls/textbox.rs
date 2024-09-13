use crate::language::controls::{
    Alignment, Appearance, BorderStyle, DragMode, LinkMode, MousePointer, OLEDragMode, OLEDropMode,
};
use crate::VB6Color;

use image::DynamicImage;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub enum ScrollBars {
    None = 0,
    Horizontal = 1,
    Vertical = 2,
    Both = 3,
}

/// Properties for a `TextBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::TextBox`](crate::language::controls::VB6ControlKind::TextBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct TextBoxProperties<'a> {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
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
    pub hide_selection: bool,
    pub left: i32,
    pub link_item: &'a str,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: &'a str,
    pub locked: bool,
    pub max_length: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub multi_line: bool,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub password_char: Option<char>,
    pub right_to_left: bool,
    pub scroll_bars: ScrollBars,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub text: &'a str,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for TextBoxProperties<'_> {
    fn default() -> Self {
        TextBoxProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H80000005&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            causes_validation: true,
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000008&").unwrap(),
            height: 30,
            help_context_id: 0,
            hide_selection: true,
            left: 30,
            link_item: "",
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "",
            locked: false,
            max_length: 0,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            multi_line: false,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            password_char: None,
            right_to_left: false,
            scroll_bars: ScrollBars::None,
            tab_index: 0,
            tab_stop: true,
            text: "",
            tool_tip_text: "",
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for TextBoxProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("TextBoxProperties", 33)?;
        s.serialize_field("alignment", &self.alignment)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("border_style", &self.border_style)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;
        s.serialize_field("data_field", &self.data_field)?;
        s.serialize_field("data_format", &self.data_format)?;
        s.serialize_field("data_member", &self.data_member)?;
        s.serialize_field("data_source", &self.data_source)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("hide_selection", &self.hide_selection)?;
        s.serialize_field("left", &self.left)?;
        s.serialize_field("link_item", &self.link_item)?;
        s.serialize_field("link_mode", &self.link_mode)?;
        s.serialize_field("link_timeout", &self.link_timeout)?;
        s.serialize_field("link_topic", &self.link_topic)?;
        s.serialize_field("locked", &self.locked)?;
        s.serialize_field("max_length", &self.max_length)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("multi_line", &self.multi_line)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("password_char", &self.password_char)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("scroll_bars", &self.scroll_bars)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("text", &self.text)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}
