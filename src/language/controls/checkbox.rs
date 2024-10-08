use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{
    Appearance, DragMode, JustifyAlignment, MousePointer, OLEDropMode, Style,
};
use crate::language::VB6Color;
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};

use bstr::BStr;
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
pub struct CheckBoxProperties<'a> {
    pub alignment: JustifyAlignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub caption: &'a BStr,
    pub causes_validation: bool,
    pub data_field: &'a BStr,
    pub data_format: &'a BStr,
    pub data_member: &'a BStr,
    pub data_source: &'a BStr,
    pub disabled_picture: Option<DynamicImage>,
    pub down_picture: Option<DynamicImage>,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub mask_color: VB6Color,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub right_to_left: bool,
    pub style: Style,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub use_mask_color: bool,
    pub value: CheckBoxValue,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for CheckBoxProperties<'_> {
    fn default() -> Self {
        CheckBoxProperties {
            alignment: JustifyAlignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            caption: BStr::new("Check1"),
            causes_validation: true,
            data_field: BStr::new(""),
            data_format: BStr::new(""),
            data_member: BStr::new(""),
            data_source: BStr::new(""),
            disabled_picture: None,
            down_picture: None,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            left: 30,
            mask_color: VB6Color::from_hex("&H00C0C0C0&").unwrap(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            right_to_left: false,
            style: Style::Standard,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: BStr::new(""),
            top: 30,
            use_mask_color: false,
            value: CheckBoxValue::Unchecked,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for CheckBoxProperties<'_> {
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

impl<'a> CheckBoxProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut checkbox_properties = CheckBoxProperties::default();

        checkbox_properties.alignment = build_property(properties, b"Alignment");
        checkbox_properties.appearance = build_property(properties, b"Appearance");
        checkbox_properties.back_color =
            build_color_property(properties, b"BackColor", checkbox_properties.back_color);
        checkbox_properties.caption = properties
            .get(BStr::new("Caption"))
            .unwrap_or(&checkbox_properties.caption);
        checkbox_properties.causes_validation = build_bool_property(
            properties,
            b"CausesValidation",
            checkbox_properties.causes_validation,
        );
        checkbox_properties.data_field = properties
            .get(BStr::new("DataField"))
            .unwrap_or(&checkbox_properties.data_field);
        checkbox_properties.data_format = properties
            .get(BStr::new("DataFormat"))
            .unwrap_or(&checkbox_properties.data_format);
        checkbox_properties.data_member = properties
            .get(BStr::new("DataMember"))
            .unwrap_or(&checkbox_properties.data_member);
        checkbox_properties.data_source = properties
            .get(BStr::new("DataSource"))
            .unwrap_or(&checkbox_properties.data_source);

        //DisabledPicture
        //DownPicture
        //DragIcon

        checkbox_properties.drag_mode = build_property(properties, b"DragMode");
        checkbox_properties.enabled =
            build_bool_property(properties, b"Enabled", checkbox_properties.enabled);
        checkbox_properties.fore_color =
            build_color_property(properties, b"ForeColor", checkbox_properties.fore_color);
        checkbox_properties.height =
            build_i32_property(properties, b"Height", checkbox_properties.height);
        checkbox_properties.help_context_id = build_i32_property(
            properties,
            b"HelpContextID",
            checkbox_properties.help_context_id,
        );
        checkbox_properties.left =
            build_i32_property(properties, b"Left", checkbox_properties.left);
        checkbox_properties.mask_color =
            build_color_property(properties, b"MaskColor", checkbox_properties.mask_color);

        //MouseIcon

        checkbox_properties.mouse_pointer = build_property(properties, b"MousePointer");
        checkbox_properties.ole_drop_mode = build_property(properties, b"OLEDropMode");

        //Picture

        checkbox_properties.right_to_left = build_bool_property(
            properties,
            b"RightToLeft",
            checkbox_properties.right_to_left,
        );
        checkbox_properties.style = build_property(properties, b"Style");
        checkbox_properties.tab_index =
            build_i32_property(properties, b"TabIndex", checkbox_properties.tab_index);
        checkbox_properties.tab_stop =
            build_bool_property(properties, b"TabStop", checkbox_properties.tab_stop);
        checkbox_properties.tool_tip_text = properties
            .get(BStr::new("ToolTipText"))
            .unwrap_or(&checkbox_properties.tool_tip_text);
        checkbox_properties.top = build_i32_property(properties, b"Top", checkbox_properties.top);
        checkbox_properties.use_mask_color = build_bool_property(
            properties,
            b"UseMaskColor",
            checkbox_properties.use_mask_color,
        );
        checkbox_properties.value = build_property(properties, b"Value");
        checkbox_properties.visible =
            build_bool_property(properties, b"Visible", checkbox_properties.visible);
        checkbox_properties.whats_this_help_id = build_i32_property(
            properties,
            b"WhatsThisHelp",
            checkbox_properties.whats_this_help_id,
        );
        checkbox_properties.width =
            build_i32_property(properties, b"Width", checkbox_properties.width);

        Ok(checkbox_properties)
    }
}
