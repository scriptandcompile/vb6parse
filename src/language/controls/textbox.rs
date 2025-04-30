use crate::{
    language::controls::{
        Alignment, Appearance, BorderStyle, DragMode, LinkMode, MousePointer, OLEDragMode,
        OLEDropMode, TextDirection, Visibility,
    },
    parsers::Properties,
    VB6Color,
};

use bstr::BString;
use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum ScrollBars {
    #[default]
    None = 0,
    Horizontal = 1,
    Vertical = 2,
    Both = 3,
}

/// 'TextBox' controls can be either multi-line or single-line.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum MultiLine {
    // The `TextBox` control is a single-line text box.
    #[default]
    SingleLine = 0,
    // The `TextBox` control is a multi-line text box.
    // Yes, the they used -1 to indicate true here.
    MultiLine = -1,
}

/// Properties for a `TextBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::TextBox`](crate::language::controls::VB6ControlKind::TextBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct TextBoxProperties {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub data_field: BString,
    pub data_format: BString,
    pub data_member: BString,
    pub data_source: BString,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub hide_selection: bool,
    pub left: i32,
    pub link_item: BString,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: BString,
    pub locked: bool,
    pub max_length: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub multi_line: MultiLine,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub password_char: Option<char>,
    pub right_to_left: TextDirection,
    pub scroll_bars: ScrollBars,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub text: BString,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for TextBoxProperties {
    fn default() -> Self {
        TextBoxProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H80000005&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            causes_validation: true,
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000008&").unwrap(),
            height: 30,
            help_context_id: 0,
            hide_selection: true,
            left: 30,
            link_item: "".into(),
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "".into(),
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
            tab_stop: true,
            text: "".into(),
            tool_tip_text: "".into(),
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

impl<'a> From<Properties<'a>> for TextBoxProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut text_box_prop = TextBoxProperties::default();

        text_box_prop.alignment = prop.get_property(b"Alignment".into(), text_box_prop.alignment);
        text_box_prop.appearance =
            prop.get_property(b"Appearance".into(), text_box_prop.appearance);
        text_box_prop.back_color = prop.get_color(b"BackColor".into(), text_box_prop.back_color);
        text_box_prop.border_style =
            prop.get_property(b"BorderStyle".into(), text_box_prop.border_style);
        text_box_prop.causes_validation =
            prop.get_bool(b"CausesValidation".into(), text_box_prop.causes_validation);
        text_box_prop.data_field = match prop.get(b"DataField".into()) {
            Some(data_field) => data_field.into(),
            None => text_box_prop.data_field,
        };
        text_box_prop.data_format = match prop.get(b"DataFormat".into()) {
            Some(data_format) => data_format.into(),
            None => text_box_prop.data_format,
        };
        text_box_prop.data_member = match prop.get(b"DataMember".into()) {
            Some(data_member) => data_member.into(),
            None => text_box_prop.data_member,
        };
        text_box_prop.data_source = match prop.get(b"DataSource".into()) {
            Some(data_source) => data_source.into(),
            None => text_box_prop.data_source,
        };

        // drag_icon: Option<DynamicImage>,

        text_box_prop.drag_mode = prop.get_property(b"DragMode".into(), text_box_prop.drag_mode);
        text_box_prop.enabled = prop.get_bool(b"Enabled".into(), text_box_prop.enabled);
        text_box_prop.fore_color = prop.get_color(b"ForeColor".into(), text_box_prop.fore_color);
        text_box_prop.height = prop.get_i32(b"Height".into(), text_box_prop.height);
        text_box_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), text_box_prop.help_context_id);
        text_box_prop.hide_selection =
            prop.get_bool(b"HideSelection".into(), text_box_prop.hide_selection);
        text_box_prop.left = prop.get_i32(b"Left".into(), text_box_prop.left);
        text_box_prop.link_item = match prop.get(b"LinkItem".into()) {
            Some(link_item) => link_item.into(),
            None => text_box_prop.link_item,
        };
        text_box_prop.link_mode = prop.get_property(b"LinkMode".into(), text_box_prop.link_mode);
        text_box_prop.link_timeout =
            prop.get_i32(b"LinkTimeout".into(), text_box_prop.link_timeout);
        text_box_prop.link_topic = match prop.get(b"LinkTopic".into()) {
            Some(link_topic) => link_topic.into(),
            None => text_box_prop.link_topic,
        };
        text_box_prop.locked = prop.get_bool(b"Locked".into(), text_box_prop.locked);
        text_box_prop.max_length = prop.get_i32(b"MaxLength".into(), text_box_prop.max_length);

        // mouse_icon: Option<DynamicImage>,

        text_box_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), text_box_prop.mouse_pointer);

        text_box_prop.multi_line = prop.get_property(b"MultiLine".into(), text_box_prop.multi_line);

        text_box_prop.ole_drag_mode =
            prop.get_property(b"OLEDragMode".into(), text_box_prop.ole_drag_mode);

        text_box_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), text_box_prop.ole_drop_mode);

        text_box_prop.password_char = match prop.get(b"PasswordChar".into()) {
            Some(password_char) => {
                if password_char.is_empty() {
                    None
                } else {
                    Some(password_char[0] as char)
                }
            }
            None => text_box_prop.password_char,
        };

        text_box_prop.right_to_left =
            prop.get_property(b"RightToLeft".into(), text_box_prop.right_to_left);

        text_box_prop.scroll_bars =
            prop.get_property(b"ScrollBars".into(), text_box_prop.scroll_bars);

        text_box_prop.tab_index = prop.get_i32(b"TabIndex".into(), text_box_prop.tab_index);

        text_box_prop.tab_stop = prop.get_bool(b"TabStop".into(), text_box_prop.tab_stop);

        text_box_prop.text = match prop.get("Text".into()) {
            Some(text) => text.into(),
            None => "".into(),
        };
        text_box_prop.tool_tip_text = match prop.get(b"ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        text_box_prop.top = prop.get_i32(b"Top".into(), text_box_prop.top);

        text_box_prop.visible = prop.get_property(b"Visible".into(), text_box_prop.visible);

        text_box_prop.whats_this_help_id =
            prop.get_i32(b"WhatsThisHelpID".into(), text_box_prop.whats_this_help_id);

        text_box_prop.width = prop.get_i32(b"Width".into(), text_box_prop.width);

        text_box_prop
    }
}
