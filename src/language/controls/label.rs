use crate::language::controls::{
    Alignment, Appearance, BackStyle, BorderStyle, DragMode, LinkMode, MousePointer, OLEDropMode,
};
use crate::VB6Color;

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `Label` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Label`](crate::language::controls::VB6ControlKind::Label).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct LabelProperties<'a> {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub auto_size: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub caption: &'a str,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub left: i32,
    pub link_item: &'a str,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: &'a str,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub tab_index: i32,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub use_mnemonic: bool,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
    pub word_wrap: bool,
}

impl Default for LabelProperties<'_> {
    fn default() -> Self {
        LabelProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            auto_size: false,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            back_style: BackStyle::Opaque,
            border_style: BorderStyle::None,
            caption: "Label1",
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            left: 30,
            link_item: "",
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "",
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::None,
            right_to_left: false,
            tab_index: 0,
            tool_tip_text: "",
            top: 30,
            use_mnemonic: true,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
            word_wrap: false,
        }
    }
}

impl Serialize for LabelProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("LabelProperties", 29)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("border_style", &self.border_style)?;
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
        s.serialize_field("left", &self.left)?;
        s.serialize_field("link_item", &self.link_item)?;
        s.serialize_field("link_mode", &self.link_mode)?;
        s.serialize_field("link_timeout", &self.link_timeout)?;
        s.serialize_field("link_topic", &self.link_topic)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("use_mnemonic", &self.use_mnemonic)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;
        s.serialize_field("word_wrap", &self.word_wrap)?;

        s.end()
    }
}
