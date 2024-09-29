use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{
    Alignment, Appearance, BackStyle, BorderStyle, DragMode, LinkMode, MousePointer, OLEDropMode,
};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};
use crate::VB6Color;

use bstr::BStr;
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
    pub caption: &'a BStr,
    pub data_field: &'a BStr,
    pub data_format: &'a BStr,
    pub data_member: &'a BStr,
    pub data_source: &'a BStr,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub left: i32,
    pub link_item: &'a BStr,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: &'a BStr,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub tab_index: i32,
    pub tool_tip_text: &'a BStr,
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
            caption: BStr::new("Label1"),
            data_field: BStr::new(""),
            data_format: BStr::new(""),
            data_member: BStr::new(""),
            data_source: BStr::new(""),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            left: 30,
            link_item: BStr::new(""),
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: BStr::new(""),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: false,
            tab_index: 0,
            tool_tip_text: BStr::new(""),
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

impl<'a> LabelProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut label_properties = LabelProperties::default();

        label_properties.alignment = build_property(properties, b"Alignment");
        label_properties.appearance = build_property(properties, b"Appearance");
        label_properties.auto_size =
            build_bool_property(properties, b"AutoSize", label_properties.auto_size);
        label_properties.back_color =
            build_color_property(properties, b"BackColor", label_properties.back_color);
        label_properties.back_style = build_property(properties, b"BackStyle");
        label_properties.border_style = build_property(properties, b"BorderStyle");
        label_properties.caption = properties
            .get(&BStr::new("Caption"))
            .unwrap_or(&label_properties.caption);
        label_properties.data_field = properties
            .get(&BStr::new("DataField"))
            .unwrap_or(&label_properties.data_field);
        label_properties.data_format = properties
            .get(&BStr::new("DataFormat"))
            .unwrap_or(&label_properties.data_format);
        label_properties.data_member = properties
            .get(&BStr::new("DataMember"))
            .unwrap_or(&label_properties.data_member);
        label_properties.data_source = properties
            .get(&BStr::new("DataSource"))
            .unwrap_or(&label_properties.data_source);

        // DragIcon

        label_properties.drag_mode = build_property(properties, b"DragMode");
        label_properties.enabled =
            build_bool_property(properties, b"Enabled", label_properties.enabled);
        label_properties.fore_color =
            build_color_property(properties, b"ForeColor", label_properties.fore_color);
        label_properties.height =
            build_i32_property(properties, b"Height", label_properties.height);
        label_properties.left = build_i32_property(properties, b"Left", label_properties.left);
        label_properties.link_item = properties
            .get(&BStr::new("LinkItem"))
            .unwrap_or(&label_properties.link_item);
        label_properties.link_mode = build_property(properties, b"LinkMode");
        label_properties.link_timeout =
            build_i32_property(properties, b"LinkTimeout", label_properties.link_timeout);
        label_properties.link_topic = properties
            .get(&BStr::new("LinkTopic"))
            .unwrap_or(&label_properties.link_topic);

        // MouseIcon

        label_properties.mouse_pointer = build_property(properties, b"MousePointer");
        label_properties.ole_drop_mode = build_property(properties, b"OLEDropMode");
        label_properties.right_to_left =
            build_bool_property(properties, b"RightToLeft", label_properties.right_to_left);
        label_properties.tab_index =
            build_i32_property(properties, b"TabIndex", label_properties.tab_index);
        label_properties.tool_tip_text = properties
            .get(&BStr::new("ToolTipText"))
            .unwrap_or(&BStr::new(""));
        label_properties.top = build_i32_property(properties, b"Top", label_properties.top);
        label_properties.use_mnemonic =
            build_bool_property(properties, b"UseMnemonic", label_properties.use_mnemonic);
        label_properties.visible =
            build_bool_property(properties, b"Visible", label_properties.visible);
        label_properties.whats_this_help_id = build_i32_property(
            properties,
            b"WhatsThisHelpID",
            label_properties.whats_this_help_id,
        );
        label_properties.width = build_i32_property(properties, b"Width", label_properties.width);
        label_properties.word_wrap =
            build_bool_property(properties, b"WordWrap", label_properties.word_wrap);

        Ok(label_properties)
    }
}
