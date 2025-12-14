use crate::{
    language::controls::{
        Activation, Alignment, Appearance, AutoSize, BackStyle, BorderStyle, DragMode, LinkMode,
        MousePointer, OLEDropMode, ReferenceOrValue, TextDirection, Visibility,
    },
    parsers::Properties,
    Color, VB_BUTTON_FACE, VB_BUTTON_TEXT,
};

use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// Determines if a `Label` control will wrap text.
#[derive(Debug, PartialEq, Clone, Copy, TryFromPrimitive, Default, serde::Serialize)]
#[repr(i32)]
pub enum WordWrap {
    /// The `Label` control will not wrap text.
    ///
    /// This is the default value.
    #[default]
    NonWrapping = 0,
    /// The `Label` control will wrap text.
    Wrapping = -1,
}

/// Properties for a `Label` control.
///
/// This is used as an enum variant of
/// [`ControlKind::Label`](crate::language::controls::ControlKind::Label).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct LabelProperties {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub auto_size: AutoSize,
    pub back_color: Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub caption: String,
    pub data_field: String,
    pub data_format: String,
    pub data_member: String,
    pub data_source: String,
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    pub drag_mode: DragMode,
    pub enabled: Activation,
    pub fore_color: Color,
    pub height: i32,
    pub left: i32,
    pub link_item: String,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: String,
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: TextDirection,
    pub tab_index: i32,
    pub tool_tip_text: String,
    pub top: i32,
    pub use_mnemonic: bool,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
    pub word_wrap: WordWrap,
}

impl Default for LabelProperties {
    fn default() -> Self {
        LabelProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            auto_size: AutoSize::Fixed,
            back_color: VB_BUTTON_FACE,
            back_style: BackStyle::Opaque,
            border_style: BorderStyle::None,
            caption: "".into(),
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            fore_color: VB_BUTTON_TEXT,
            height: 30,
            left: 30,
            link_item: "".into(),
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "".into(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: TextDirection::LeftToRight,
            tab_index: 0,
            tool_tip_text: "".into(),
            top: 30,
            use_mnemonic: true,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
            word_wrap: WordWrap::NonWrapping,
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

impl From<Properties> for LabelProperties {
    fn from(prop: Properties) -> Self {
        let mut label_prop = LabelProperties::default();

        label_prop.alignment = prop.get_property("Alignment", label_prop.alignment);
        label_prop.appearance = prop.get_property("Appearance", label_prop.appearance);
        label_prop.auto_size = prop.get_property("AutoSize", label_prop.auto_size);
        label_prop.back_color = prop.get_color("BackColor", label_prop.back_color);
        label_prop.back_style = prop.get_property("BackStyle", label_prop.back_style);
        label_prop.border_style = prop.get_property("BorderStyle", label_prop.border_style);
        label_prop.caption = match prop.get("Caption") {
            Some(caption) => caption.into(),
            None => "".into(),
        };
        label_prop.data_field = match prop.get("DataField") {
            Some(data_field) => data_field.into(),
            None => "".into(),
        };
        label_prop.data_format = match prop.get("DataFormat") {
            Some(data_format) => data_format.into(),
            None => "".into(),
        };
        label_prop.data_member = match prop.get("DataMember") {
            Some(data_member) => data_member.into(),
            None => "".into(),
        };
        label_prop.data_source = match prop.get("DataSource") {
            Some(data_source) => data_source.into(),
            None => "".into(),
        };

        // DragIcon

        label_prop.drag_mode = prop.get_property("DragMode", label_prop.drag_mode);
        label_prop.enabled = prop.get_property("Enabled", label_prop.enabled);
        label_prop.fore_color = prop.get_color("ForeColor", label_prop.fore_color);
        label_prop.height = prop.get_i32("Height", label_prop.height);
        label_prop.left = prop.get_i32("Left", label_prop.left);
        label_prop.link_item = match prop.get("LinkItem") {
            Some(link_item) => link_item.into(),
            None => "".into(),
        };
        label_prop.link_mode = prop.get_property("LinkMode", label_prop.link_mode);
        label_prop.link_timeout = prop.get_i32("LinkTimeout", label_prop.link_timeout);
        label_prop.link_topic = match prop.get("LinkTopic") {
            Some(link_topic) => link_topic.into(),
            None => "".into(),
        };

        // MouseIcon

        label_prop.mouse_pointer = prop.get_property("MousePointer", label_prop.mouse_pointer);
        label_prop.ole_drop_mode = prop.get_property("OLEDropMode", label_prop.ole_drop_mode);
        label_prop.right_to_left = prop.get_property("RightToLeft", label_prop.right_to_left);
        label_prop.tab_index = prop.get_i32("TabIndex", label_prop.tab_index);
        label_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        label_prop.top = prop.get_i32("Top", label_prop.top);
        label_prop.use_mnemonic = prop.get_bool("UseMnemonic", label_prop.use_mnemonic);
        label_prop.visible = prop.get_property("Visible", label_prop.visible);
        label_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", label_prop.whats_this_help_id);
        label_prop.width = prop.get_i32("Width", label_prop.width);
        label_prop.word_wrap = prop.get_property("WordWrap", label_prop.word_wrap);

        label_prop
    }
}
