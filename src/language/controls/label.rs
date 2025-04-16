use crate::language::controls::{
    Alignment, Appearance, BackStyle, BorderStyle, DragMode, LinkMode, MousePointer, OLEDropMode,
};
use crate::parsers::Properties;
use crate::VB6Color;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `Label` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Label`](crate::language::controls::VB6ControlKind::Label).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct LabelProperties {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub auto_size: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub caption: BString,
    pub data_field: BString,
    pub data_format: BString,
    pub data_member: BString,
    pub data_source: BString,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub left: i32,
    pub link_item: BString,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: BString,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub tab_index: i32,
    pub tool_tip_text: BString,
    pub top: i32,
    pub use_mnemonic: bool,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
    pub word_wrap: bool,
}

impl Default for LabelProperties {
    fn default() -> Self {
        LabelProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            auto_size: false,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            back_style: BackStyle::Opaque,
            border_style: BorderStyle::None,
            caption: "".into(),
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            left: 30,
            link_item: "".into(),
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "".into(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: false,
            tab_index: 0,
            tool_tip_text: "".into(),
            top: 30,
            use_mnemonic: true,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
            word_wrap: false,
        }
    }
}

impl Serialize for LabelProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("LabelProperties", 33)?;
        s.serialize_field("alignment", &self.alignment)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("auto_size", &self.auto_size)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("back_style", &self.back_style)?;
        s.serialize_field("border_style", &self.border_style)?;
        s.serialize_field("caption", &self.caption)?;
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

impl<'a> From<Properties<'a>> for LabelProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut label_prop = LabelProperties::default();

        label_prop.alignment = prop.get_property(b"Alignment".into(), label_prop.alignment);
        label_prop.appearance = prop.get_property(b"Appearance".into(), label_prop.appearance);
        label_prop.auto_size = prop.get_bool(b"AutoSize".into(), label_prop.auto_size);
        label_prop.back_color = prop.get_color(b"BackColor".into(), label_prop.back_color);
        label_prop.back_style = prop.get_property(b"BackStyle".into(), label_prop.back_style);
        label_prop.border_style = prop.get_property(b"BorderStyle".into(), label_prop.border_style);
        label_prop.caption = match prop.get(b"Caption".into()) {
            Some(caption) => caption.into(),
            None => "".into(),
        };
        label_prop.data_field = match prop.get(b"DataField".into()) {
            Some(data_field) => data_field.into(),
            None => "".into(),
        };
        label_prop.data_format = match prop.get(b"DataFormat".into()) {
            Some(data_format) => data_format.into(),
            None => "".into(),
        };
        label_prop.data_member = match prop.get(b"DataMember".into()) {
            Some(data_member) => data_member.into(),
            None => "".into(),
        };
        label_prop.data_source = match prop.get(b"DataSource".into()) {
            Some(data_source) => data_source.into(),
            None => "".into(),
        };

        // DragIcon

        label_prop.drag_mode = prop.get_property(b"DragMode".into(), label_prop.drag_mode);
        label_prop.enabled = prop.get_bool(b"Enabled".into(), label_prop.enabled);
        label_prop.fore_color = prop.get_color(b"ForeColor".into(), label_prop.fore_color);
        label_prop.height = prop.get_i32(b"Height".into(), label_prop.height);
        label_prop.left = prop.get_i32(b"Left".into(), label_prop.left);
        label_prop.link_item = match prop.get(b"LinkItem".into()) {
            Some(link_item) => link_item.into(),
            None => "".into(),
        };
        label_prop.link_mode = prop.get_property(b"LinkMode".into(), label_prop.link_mode);
        label_prop.link_timeout = prop.get_i32(b"LinkTimeout".into(), label_prop.link_timeout);
        label_prop.link_topic = match prop.get(b"LinkTopic".into()) {
            Some(link_topic) => link_topic.into(),
            None => "".into(),
        };

        // MouseIcon

        label_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), label_prop.mouse_pointer);
        label_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), label_prop.ole_drop_mode);
        label_prop.right_to_left = prop.get_bool(b"RightToLeft".into(), label_prop.right_to_left);
        label_prop.tab_index = prop.get_i32(b"TabIndex".into(), label_prop.tab_index);
        label_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        label_prop.top = prop.get_i32(b"Top".into(), label_prop.top);
        label_prop.use_mnemonic = prop.get_bool(b"UseMnemonic".into(), label_prop.use_mnemonic);
        label_prop.visible = prop.get_bool(b"Visible".into(), label_prop.visible);
        label_prop.whats_this_help_id =
            prop.get_i32(b"WhatsThisHelpID".into(), label_prop.whats_this_help_id);
        label_prop.width = prop.get_i32(b"Width".into(), label_prop.width);
        label_prop.word_wrap = prop.get_bool(b"WordWrap".into(), label_prop.word_wrap);

        label_prop
    }
}
