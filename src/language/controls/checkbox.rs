use crate::{
    language::{
        controls::{
            Activation, Appearance, DragMode, JustifyAlignment, MousePointer, OLEDropMode, Style,
            TabStop, TextDirection, Visibility,
        },
        VB6Color,
    },
    parsers::Properties,
};

use bstr::BString;
use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum CheckBoxValue {
    #[default]
    Unchecked = 0,
    Checked = 1,
    Grayed = 2,
}

/// Properties for a `CheckBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::CheckBox`](crate::language::controls::VB6ControlKind::CheckBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct CheckBoxProperties {
    pub alignment: JustifyAlignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub caption: BString,
    pub causes_validation: bool,
    pub data_field: BString,
    pub data_format: BString,
    pub data_member: BString,
    pub data_source: BString,
    pub disabled_picture: Option<DynamicImage>,
    pub down_picture: Option<DynamicImage>,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: Activation,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub mask_color: VB6Color,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub right_to_left: TextDirection,
    pub style: Style,
    pub tab_index: i32,
    pub tab_stop: TabStop,
    pub tool_tip_text: BString,
    pub top: i32,
    pub use_mask_color: bool,
    pub value: CheckBoxValue,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for CheckBoxProperties {
    fn default() -> Self {
        CheckBoxProperties {
            alignment: JustifyAlignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            caption: "".into(),
            causes_validation: true,
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            disabled_picture: None,
            down_picture: None,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            left: 30,
            mask_color: VB6Color::from_hex("&H00C0C0C0&").unwrap(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            right_to_left: TextDirection::LeftToRight,
            style: Style::Standard,
            tab_index: 0,
            tab_stop: TabStop::Included,
            tool_tip_text: "".into(),
            top: 30,
            use_mask_color: false,
            value: CheckBoxValue::Unchecked,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for CheckBoxProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::ser::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("CheckBoxProperties", 29)?;
        state.serialize_field("alignment", &self.alignment)?;
        state.serialize_field("appearance", &self.appearance)?;
        state.serialize_field("back_color", &self.back_color)?;
        state.serialize_field("caption", &self.caption)?;
        state.serialize_field("causes_validation", &self.causes_validation)?;
        state.serialize_field("data_field", &self.data_field)?;
        state.serialize_field("data_format", &self.data_format)?;
        state.serialize_field("data_member", &self.data_member)?;
        state.serialize_field("data_source", &self.data_source)?;

        let option_text = self.disabled_picture.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("disabled_picture", &option_text)?;

        let option_text = self.down_picture.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("down_picture", &option_text)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("drag_icon", &option_text)?;
        state.serialize_field("drag_mode", &self.drag_mode)?;
        state.serialize_field("enabled", &self.enabled)?;
        state.serialize_field("fore_color", &self.fore_color)?;
        state.serialize_field("height", &self.height)?;
        state.serialize_field("help_context_id", &self.help_context_id)?;
        state.serialize_field("left", &self.left)?;
        state.serialize_field("mask_color", &self.mask_color)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("mouse_icon", &option_text)?;
        state.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        state.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;

        let option_text = self.picture.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("picture", &option_text)?;
        state.serialize_field("right_to_left", &self.right_to_left)?;
        state.serialize_field("style", &self.style)?;
        state.serialize_field("tab_index", &self.tab_index)?;
        state.serialize_field("tab_stop", &self.tab_stop)?;
        state.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        state.serialize_field("top", &self.top)?;
        state.serialize_field("use_mask_color", &self.use_mask_color)?;
        state.serialize_field("value", &self.value)?;
        state.serialize_field("visible", &self.visible)?;
        state.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        state.serialize_field("width", &self.width)?;

        state.end()
    }
}

impl<'a> From<Properties<'a>> for CheckBoxProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut checkbox_prop = CheckBoxProperties::default();

        checkbox_prop.alignment = prop.get_property(b"Alignment".into(), checkbox_prop.alignment);
        checkbox_prop.appearance =
            prop.get_property(b"Appearance".into(), checkbox_prop.appearance);
        checkbox_prop.back_color = prop.get_color(b"BackColor".into(), checkbox_prop.back_color);
        checkbox_prop.caption = match prop.get("Caption".into()) {
            Some(caption) => caption.into(),
            None => checkbox_prop.caption,
        };
        checkbox_prop.causes_validation =
            prop.get_bool(b"CausesValidation".into(), checkbox_prop.causes_validation);
        checkbox_prop.data_field = match prop.get(b"DataField".into()) {
            Some(data_field) => data_field.into(),
            None => checkbox_prop.data_field,
        };
        checkbox_prop.data_format = match prop.get("DataFormat".into()) {
            Some(data_format) => data_format.into(),
            None => checkbox_prop.data_format,
        };
        checkbox_prop.data_member = match prop.get("DataMember".into()) {
            Some(data_member) => data_member.into(),
            None => checkbox_prop.data_member,
        };
        checkbox_prop.data_source = match prop.get("DataSource".into()) {
            Some(data_source) => data_source.into(),
            None => checkbox_prop.data_source,
        };
        //DisabledPicture
        //DownPicture
        //DragIcon

        checkbox_prop.drag_mode = prop.get_property(b"DragMode".into(), checkbox_prop.drag_mode);
        checkbox_prop.enabled = prop.get_property(b"Enabled".into(), checkbox_prop.enabled);
        checkbox_prop.fore_color = prop.get_color(b"ForeColor".into(), checkbox_prop.fore_color);
        checkbox_prop.height = prop.get_i32(b"Height".into(), checkbox_prop.height);
        checkbox_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), checkbox_prop.help_context_id);
        checkbox_prop.left = prop.get_i32(b"Left".into(), checkbox_prop.left);
        checkbox_prop.mask_color = prop.get_color(b"MaskColor".into(), checkbox_prop.mask_color);

        //MouseIcon

        checkbox_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), checkbox_prop.mouse_pointer);
        checkbox_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), checkbox_prop.ole_drop_mode);

        //Picture

        checkbox_prop.right_to_left =
            prop.get_property(b"RightToLeft".into(), checkbox_prop.right_to_left);
        checkbox_prop.style = prop.get_property(b"Style".into(), checkbox_prop.style);
        checkbox_prop.tab_index = prop.get_i32(b"TabIndex".into(), checkbox_prop.tab_index);
        checkbox_prop.tab_stop = prop.get_property(b"TabStop".into(), checkbox_prop.tab_stop);
        checkbox_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => checkbox_prop.tool_tip_text,
        };
        checkbox_prop.top = prop.get_i32(b"Top".into(), checkbox_prop.top);
        checkbox_prop.use_mask_color =
            prop.get_bool(b"UseMaskColor".into(), checkbox_prop.use_mask_color);
        checkbox_prop.value = prop.get_property(b"Value".into(), checkbox_prop.value);
        checkbox_prop.visible = prop.get_property(b"Visible".into(), checkbox_prop.visible);
        checkbox_prop.whats_this_help_id =
            prop.get_i32(b"WhatsThisHelp".into(), checkbox_prop.whats_this_help_id);
        checkbox_prop.width = prop.get_i32(b"Width".into(), checkbox_prop.width);

        checkbox_prop
    }
}
