//! Properties for `TextBox` controls.
//!
//! This is used as an enum variant of
//! [`ControlKind::TextBox`](crate::language::controls::ControlKind::TextBox).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use std::convert::TryFrom;
use std::fmt::{Display, Formatter};
use std::str::FromStr;

use crate::{
    files::common::Properties,
    language::{
        color::{Color, VB_WINDOW_BACKGROUND, VB_WINDOW_TEXT},
        controls::{
            Activation, Alignment, Appearance, BorderStyle, CausesValidation, DragMode, LinkMode,
            MousePointer, OLEDragMode, OLEDropMode, ReferenceOrValue, TabStop, TextDirection,
            Visibility,
        },
    },
};

use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// Determines whether an object has horizontal or vertical scroll bars.
///
/// Note:
/// For a `TextBox` control, the multiline property must be set to `true`
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445672(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum ScrollBars {
    /// No scroll bars are displayed.
    ///
    /// This is the default value.
    #[default]
    None = 0,
    /// A horizontal scroll bar is displayed.
    Horizontal = 1,
    /// A vertical scroll bar is displayed.
    Vertical = 2,
    /// Both horizontal and vertical scroll bars are displayed.
    Both = 3,
}

impl TryFrom<&str> for ScrollBars {
    type Error = crate::errors::FormErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(ScrollBars::None),
            "1" => Ok(ScrollBars::Horizontal),
            "2" => Ok(ScrollBars::Vertical),
            "3" => Ok(ScrollBars::Both),
            _ => Err(crate::errors::FormErrorKind::InvalidScrollBars(
                value.to_string(),
            )),
        }
    }
}

impl FromStr for ScrollBars {
    type Err = crate::errors::FormErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "0" => Ok(ScrollBars::None),
            "1" => Ok(ScrollBars::Horizontal),
            "2" => Ok(ScrollBars::Vertical),
            "3" => Ok(ScrollBars::Both),
            _ => Err(crate::errors::FormErrorKind::InvalidScrollBars(
                s.to_string(),
            )),
        }
    }
}

impl Display for ScrollBars {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            ScrollBars::None => "None",
            ScrollBars::Horizontal => "Horizontal",
            ScrollBars::Vertical => "Vertical",
            ScrollBars::Both => "Both",
        };
        write!(f, "{text}")
    }
}

/// `TextBox` controls can be either multi-line or single-line.
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum MultiLine {
    /// The `TextBox` control is a single-line text box.
    #[default]
    SingleLine = 0,
    /// The `TextBox` control is a multi-line text box.
    MultiLine = -1,
}

impl TryFrom<bool> for MultiLine {
    type Error = crate::errors::FormErrorKind;

    fn try_from(value: bool) -> Result<Self, Self::Error> {
        if value {
            Ok(MultiLine::MultiLine)
        } else {
            Ok(MultiLine::SingleLine)
        }
    }
}

impl TryFrom<&str> for MultiLine {
    type Error = crate::errors::FormErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(MultiLine::SingleLine),
            "-1" => Ok(MultiLine::MultiLine),
            _ => Err(crate::errors::FormErrorKind::InvalidMultiLine(
                value.to_string(),
            )),
        }
    }
}

impl FromStr for MultiLine {
    type Err = crate::errors::FormErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        MultiLine::try_from(s)
    }
}

impl Display for MultiLine {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            MultiLine::SingleLine => "SingleLine",
            MultiLine::MultiLine => "MultiLine",
        };
        write!(f, "{text}")
    }
}

/// Properties for a `TextBox` control.
///
/// This is used as an enum variant of
/// [`ControlKind::TextBox`](crate::language::controls::ControlKind::TextBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct TextBoxProperties {
    /// Text alignment within the text box.
    pub alignment: Alignment,
    /// Appearance of the text box.
    pub appearance: Appearance,
    /// Background color of the text box.
    pub back_color: Color,
    /// Border style of the text box.
    pub border_style: BorderStyle,
    /// Indicates whether the control causes validation when it receives focus.
    pub causes_validation: CausesValidation,
    /// Data field associated with the text box.
    pub data_field: String,
    /// Data format for displaying the text.
    pub data_format: String,
    /// Data member associated with the text box.
    pub data_member: String,
    /// Data source associated with the text box.
    pub data_source: String,
    /// Icon displayed when dragging the text box.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Indicates how the control can be dragged.
    pub drag_mode: DragMode,
    /// Indicates whether the control is enabled.
    pub enabled: Activation,
    /// Foreground color of the text box.
    pub fore_color: Color,
    /// Height of the text box control.
    pub height: i32,
    /// Help context ID associated with the control.
    pub help_context_id: i32,
    /// Indicates whether the selected text remains highlighted when the text box loses focus.
    pub hide_selection: bool,
    /// Left position of the text box control.
    pub left: i32,
    /// Link item associated with the text box.
    pub link_item: String,
    /// Link mode of the text box.
    pub link_mode: LinkMode,
    /// Link timeout for the text box.
    pub link_timeout: i32,
    /// Link topic associated with the text box.
    pub link_topic: String,
    /// Indicates whether the text box is locked for editing.
    pub locked: bool,
    /// Maximum length of text that can be entered in the text box.
    pub max_length: i32,
    /// Icon displayed when the mouse is over the text box.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Mouse pointer type when hovering over the text box.
    pub mouse_pointer: MousePointer,
    /// Indicates whether the text box is multi-line or single-line.
    pub multi_line: MultiLine,
    /// OLE drag mode of the text box.
    pub ole_drag_mode: OLEDragMode,
    /// OLE drop mode of the text box.
    pub ole_drop_mode: OLEDropMode,
    /// Character used to mask input in password mode.
    pub password_char: Option<char>,
    /// Text direction for the text box.
    pub right_to_left: TextDirection,
    /// Scroll bars displayed in the text box.
    pub scroll_bars: ScrollBars,
    /// Tab index of the text box control.
    pub tab_index: i32,
    /// Indicates whether the control is included in the tab order.
    pub tab_stop: TabStop,
    /// Text content of the text box.
    pub text: String,
    /// Tool tip text for the text box.
    pub tool_tip_text: String,
    /// Top position of the text box control.
    pub top: i32,
    /// Visibility of the text box control.
    pub visible: Visibility,
    /// "What's This?" help context ID associated with the control.
    pub whats_this_help_id: i32,
    /// Width of the text box control.
    pub width: i32,
}

impl Default for TextBoxProperties {
    fn default() -> Self {
        TextBoxProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB_WINDOW_BACKGROUND,
            border_style: BorderStyle::FixedSingle,
            causes_validation: CausesValidation::Yes,
            data_field: String::new(),
            data_format: String::new(),
            data_member: String::new(),
            data_source: String::new(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            fore_color: VB_WINDOW_TEXT,
            height: 30,
            help_context_id: 0,
            hide_selection: true,
            left: 30,
            link_item: String::new(),
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: String::new(),
            locked: false,
            max_length: 0,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            multi_line: MultiLine::SingleLine,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            password_char: None,
            right_to_left: TextDirection::LeftToRight,
            scroll_bars: ScrollBars::None,
            tab_index: 0,
            tab_stop: TabStop::Included,
            text: String::new(),
            tool_tip_text: String::new(),
            top: 30,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for TextBoxProperties {
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

impl From<Properties> for TextBoxProperties {
    fn from(prop: Properties) -> Self {
        let mut text_box_prop = TextBoxProperties::default();

        text_box_prop.alignment = prop.get_property("Alignment", text_box_prop.alignment);
        text_box_prop.appearance = prop.get_property("Appearance", text_box_prop.appearance);
        text_box_prop.back_color = prop.get_color("BackColor", text_box_prop.back_color);
        text_box_prop.border_style = prop.get_property("BorderStyle", text_box_prop.border_style);
        text_box_prop.causes_validation =
            prop.get_property("CausesValidation", text_box_prop.causes_validation);
        text_box_prop.data_field = match prop.get("DataField") {
            Some(data_field) => data_field.into(),
            None => text_box_prop.data_field,
        };
        text_box_prop.data_format = match prop.get("DataFormat") {
            Some(data_format) => data_format.into(),
            None => text_box_prop.data_format,
        };
        text_box_prop.data_member = match prop.get("DataMember") {
            Some(data_member) => data_member.into(),
            None => text_box_prop.data_member,
        };
        text_box_prop.data_source = match prop.get("DataSource") {
            Some(data_source) => data_source.into(),
            None => text_box_prop.data_source,
        };

        // TODO: process DragIcon
        // drag_icon: Option<DynamicImage>,

        text_box_prop.drag_mode = prop.get_property("DragMode", text_box_prop.drag_mode);
        text_box_prop.enabled = prop.get_property("Enabled", text_box_prop.enabled);
        text_box_prop.fore_color = prop.get_color("ForeColor", text_box_prop.fore_color);
        text_box_prop.height = prop.get_i32("Height", text_box_prop.height);
        text_box_prop.help_context_id =
            prop.get_i32("HelpContextID", text_box_prop.help_context_id);
        text_box_prop.hide_selection = prop.get_bool("HideSelection", text_box_prop.hide_selection);
        text_box_prop.left = prop.get_i32("Left", text_box_prop.left);
        text_box_prop.link_item = match prop.get("LinkItem") {
            Some(link_item) => link_item.into(),
            None => text_box_prop.link_item,
        };
        text_box_prop.link_mode = prop.get_property("LinkMode", text_box_prop.link_mode);
        text_box_prop.link_timeout = prop.get_i32("LinkTimeout", text_box_prop.link_timeout);
        text_box_prop.link_topic = match prop.get("LinkTopic") {
            Some(link_topic) => link_topic.into(),
            None => text_box_prop.link_topic,
        };
        text_box_prop.locked = prop.get_bool("Locked", text_box_prop.locked);
        text_box_prop.max_length = prop.get_i32("MaxLength", text_box_prop.max_length);

        // TODO: process MouseIcon
        // mouse_icon: Option<DynamicImage>,

        text_box_prop.mouse_pointer =
            prop.get_property("MousePointer", text_box_prop.mouse_pointer);

        text_box_prop.multi_line = prop.get_property("MultiLine", text_box_prop.multi_line);

        text_box_prop.ole_drag_mode = prop.get_property("OLEDragMode", text_box_prop.ole_drag_mode);

        text_box_prop.ole_drop_mode = prop.get_property("OLEDropMode", text_box_prop.ole_drop_mode);

        text_box_prop.password_char = match prop.get("PasswordChar") {
            Some(password_char) => {
                if password_char.is_empty() {
                    None
                } else {
                    password_char.chars().next()
                }
            }
            None => text_box_prop.password_char,
        };

        text_box_prop.right_to_left = prop.get_property("RightToLeft", text_box_prop.right_to_left);

        text_box_prop.scroll_bars = prop.get_property("ScrollBars", text_box_prop.scroll_bars);

        text_box_prop.tab_index = prop.get_i32("TabIndex", text_box_prop.tab_index);

        text_box_prop.tab_stop = prop.get_property("TabStop", text_box_prop.tab_stop);

        text_box_prop.text = match prop.get("Text") {
            Some(text) => text.into(),
            None => String::new(),
        };
        text_box_prop.tool_tip_text = match prop.get("ToolTipText") {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => String::new(),
        };
        text_box_prop.top = prop.get_i32("Top", text_box_prop.top);

        text_box_prop.visible = prop.get_property("Visible", text_box_prop.visible);

        text_box_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", text_box_prop.whats_this_help_id);

        text_box_prop.width = prop.get_i32("Width", text_box_prop.width);

        text_box_prop
    }
}
