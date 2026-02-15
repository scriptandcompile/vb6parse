//! Properties for a `Label` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::Label`](crate::language::controls::ControlKind::Label).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use std::convert::TryFrom;
use std::fmt::Display;
use std::str::FromStr;

use crate::{
    errors::{FormError, ErrorKind},
    files::common::Properties,
    language::{
        color::{Color, VB_BUTTON_FACE, VB_BUTTON_TEXT},
        controls::{
            Activation, Alignment, Appearance, AutoSize, BackStyle, BorderStyle, DragMode,
            LinkMode, MousePointer, OLEDropMode, ReferenceOrValue, TextDirection, Visibility,
        },
    },
};

use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// Determines if a `Label` control will wrap text.
#[derive(
    Debug,
    PartialEq,
    Clone,
    Copy,
    TryFromPrimitive,
    Default,
    serde::Serialize,
    Eq,
    Hash,
    PartialOrd,
    Ord,
)]
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

impl FromStr for WordWrap {
    type Err = crate::errors::ErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        WordWrap::try_from(s)
    }
}

impl TryFrom<&str> for WordWrap {
    type Error = crate::errors::ErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" | "NonWrapping" => Ok(WordWrap::NonWrapping),
            "-1" | "Wrapping" => Ok(WordWrap::Wrapping),
            _ => Err(ErrorKind::Form(FormError::InvalidWordWrap {
                value: value.to_string(),
            })),
        }
    }
}

impl TryFrom<bool> for WordWrap {
    type Error = crate::errors::ErrorKind;

    fn try_from(value: bool) -> Result<Self, Self::Error> {
        if value {
            Ok(WordWrap::Wrapping)
        } else {
            Ok(WordWrap::NonWrapping)
        }
    }
}

impl Display for WordWrap {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            WordWrap::NonWrapping => "NonWrapping",
            WordWrap::Wrapping => "Wrapping",
        };
        write!(f, "{text}")
    }
}

/// Properties for a `Label` control.
///
/// This is used as an enum variant of
/// [`ControlKind::Label`](crate::language::controls::ControlKind::Label).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct LabelProperties {
    /// Alignment of the label.
    pub alignment: Alignment,
    /// Appearance of the label.
    pub appearance: Appearance,
    /// Auto size setting of the label.
    pub auto_size: AutoSize,
    /// Background color of the label.
    pub back_color: Color,
    /// Back style of the label.
    pub back_style: BackStyle,
    /// Border style of the label.
    pub border_style: BorderStyle,
    /// Caption of the label.
    pub caption: String,
    /// Data field of the label.
    pub data_field: String,
    /// Data format of the label.
    pub data_format: String,
    /// Data member of the label.
    pub data_member: String,
    /// Data source of the label.
    pub data_source: String,
    /// Drag icon of the label.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Drag mode of the label.
    pub drag_mode: DragMode,
    /// Enabled state of the label.
    pub enabled: Activation,
    /// Foreground color of the label.
    pub fore_color: Color,
    /// Height of the label.
    pub height: i32,
    /// Left position of the label.
    pub left: i32,
    /// Link item of the label.
    pub link_item: String,
    /// Link mode of the label.
    pub link_mode: LinkMode,
    /// Link timeout of the label.
    pub link_timeout: i32,
    /// Link topic of the label.
    pub link_topic: String,
    /// Mouse icon of the label.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Mouse pointer of the label.
    pub mouse_pointer: MousePointer,
    /// OLE drop mode of the label.
    pub ole_drop_mode: OLEDropMode,
    /// Right to left setting of the label.
    pub right_to_left: TextDirection,
    /// Tab index of the label.
    pub tab_index: i32,
    /// Tool tip text of the label.
    pub tool_tip_text: String,
    /// Top position of the label.
    pub top: i32,
    /// Use mnemonic of the label.
    pub use_mnemonic: bool,
    /// Visibility of the label.
    pub visible: Visibility,
    /// What's this help ID of the label.
    pub whats_this_help_id: i32,
    /// Width of the label. ]
    pub width: i32,
    /// Word wrap setting of the label.
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
            caption: String::new(),
            data_field: String::new(),
            data_format: String::new(),
            data_member: String::new(),
            data_source: String::new(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            fore_color: VB_BUTTON_TEXT,
            height: 30,
            left: 30,
            link_item: String::new(),
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: String::new(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: TextDirection::LeftToRight,
            tab_index: 0,
            tool_tip_text: String::new(),
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
            None => String::new(),
        };
        label_prop.data_field = match prop.get("DataField") {
            Some(data_field) => data_field.into(),
            None => String::new(),
        };
        label_prop.data_format = match prop.get("DataFormat") {
            Some(data_format) => data_format.into(),
            None => String::new(),
        };
        label_prop.data_member = match prop.get("DataMember") {
            Some(data_member) => data_member.into(),
            None => String::new(),
        };
        label_prop.data_source = match prop.get("DataSource") {
            Some(data_source) => data_source.into(),
            None => String::new(),
        };

        // TODO: process drag_icon
        // DragIcon

        label_prop.drag_mode = prop.get_property("DragMode", label_prop.drag_mode);
        label_prop.enabled = prop.get_property("Enabled", label_prop.enabled);
        label_prop.fore_color = prop.get_color("ForeColor", label_prop.fore_color);
        label_prop.height = prop.get_i32("Height", label_prop.height);
        label_prop.left = prop.get_i32("Left", label_prop.left);
        label_prop.link_item = match prop.get("LinkItem") {
            Some(link_item) => link_item.into(),
            None => String::new(),
        };
        label_prop.link_mode = prop.get_property("LinkMode", label_prop.link_mode);
        label_prop.link_timeout = prop.get_i32("LinkTimeout", label_prop.link_timeout);
        label_prop.link_topic = match prop.get("LinkTopic") {
            Some(link_topic) => link_topic.into(),
            None => String::new(),
        };

        // TODO: process mouse_icon
        // MouseIcon

        label_prop.mouse_pointer = prop.get_property("MousePointer", label_prop.mouse_pointer);
        label_prop.ole_drop_mode = prop.get_property("OLEDropMode", label_prop.ole_drop_mode);
        label_prop.right_to_left = prop.get_property("RightToLeft", label_prop.right_to_left);
        label_prop.tab_index = prop.get_i32("TabIndex", label_prop.tab_index);
        label_prop.tool_tip_text = match prop.get("ToolTipText") {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => String::new(),
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
