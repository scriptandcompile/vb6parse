use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{Appearance, DragMode, MousePointer, OLEDragMode, OLEDropMode};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};
use crate::VB6Color;

use bstr::BStr;
use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum ComboBoxStyle {
    #[default]
    DropDownCombo = 0,
    SimpleCombo = 1,
    DropDownList = 2,
}

/// Properties for a `ComboBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::ComboBox`](crate::language::controls::VB6ControlKind::ComboBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ComboBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    pub data_field: &'a BStr,
    pub data_format: &'a BStr,
    pub data_member: &'a BStr,
    pub data_source: &'a BStr,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub integral_height: bool,
    //pub item_data: Vec<&'a BStr>,
    pub left: i32,
    // pub list: Vec<&'a BStr>,
    pub locked: bool,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub sorted: bool,
    pub style: ComboBoxStyle,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub text: &'a BStr,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ComboBoxProperties<'_> {
    fn default() -> Self {
        ComboBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H80000005&").unwrap(),
            causes_validation: true,
            data_field: BStr::new(""),
            data_format: BStr::new(""),
            data_member: BStr::new(""),
            data_source: BStr::new(""),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000008&").unwrap(),
            height: 30,
            help_context_id: 0,
            integral_height: true,
            //item_data:
            left: 30,
            //list:
            locked: false,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: false,
            sorted: false,
            style: ComboBoxStyle::DropDownCombo,
            tab_index: 0,
            tab_stop: true,
            text: BStr::new(""),
            tool_tip_text: BStr::new(""),
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for ComboBoxProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::ser::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("ComboBoxProperties", 26)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
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
        s.serialize_field("integral_height", &self.integral_height)?;
        //s.serialize_field("item_data", &self.item_data)?;
        s.serialize_field("left", &self.left)?;
        //s.serialize_field("list", &self.list)?;
        s.serialize_field("locked", &self.locked)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("sorted", &self.sorted)?;
        s.serialize_field("style", &self.style)?;
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

impl<'a> ComboBoxProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut combobox_properties = ComboBoxProperties::default();

        combobox_properties.appearance =
            build_property::<Appearance>(&properties, BStr::new("Appearance"));

        combobox_properties.back_color = build_color_property(
            &properties,
            BStr::new("BackColor"),
            combobox_properties.back_color,
        );

        combobox_properties.causes_validation = build_bool_property(
            &properties,
            BStr::new("CausesValidation"),
            combobox_properties.causes_validation,
        );
        let data_field_key = BStr::new("DataField");
        combobox_properties.data_field = properties
            .get(data_field_key)
            .unwrap_or(&combobox_properties.data_field);

        let data_format_key = BStr::new("DataFormat");
        combobox_properties.data_format = properties
            .get(data_format_key)
            .unwrap_or(&combobox_properties.data_format);

        let data_member_key = BStr::new("DataMember");
        combobox_properties.data_member = properties
            .get(data_member_key)
            .unwrap_or(&combobox_properties.data_member);

        let data_source_key = BStr::new("DataSource");
        combobox_properties.data_source = properties
            .get(data_source_key)
            .unwrap_or(&combobox_properties.data_source);

        // drag_icon

        combobox_properties.drag_mode =
            build_property::<DragMode>(&properties, BStr::new("DragMode"));

        combobox_properties.enabled = build_bool_property(
            &properties,
            BStr::new("Enabled"),
            combobox_properties.enabled,
        );

        combobox_properties.fore_color = build_color_property(
            &properties,
            BStr::new("ForeColor"),
            combobox_properties.fore_color,
        );

        combobox_properties.height =
            build_i32_property(&properties, BStr::new("Height"), combobox_properties.height);

        combobox_properties.help_context_id = build_i32_property(
            &properties,
            BStr::new("HelpContextID"),
            combobox_properties.help_context_id,
        );

        combobox_properties.integral_height = build_bool_property(
            &properties,
            BStr::new("IntegralHeight"),
            combobox_properties.integral_height,
        );

        combobox_properties.left =
            build_i32_property(&properties, BStr::new("Left"), combobox_properties.left);

        combobox_properties.locked =
            build_bool_property(&properties, BStr::new("Locked"), combobox_properties.locked);

        // mouse_icon

        combobox_properties.mouse_pointer =
            build_property::<MousePointer>(&properties, BStr::new("MousePointer"));

        combobox_properties.ole_drag_mode =
            build_property::<OLEDragMode>(&properties, BStr::new("OLEDragMode"));

        combobox_properties.ole_drop_mode =
            build_property::<OLEDropMode>(&properties, BStr::new("OLEDropMode"));

        combobox_properties.right_to_left = build_bool_property(
            &properties,
            BStr::new("RightToLeft"),
            combobox_properties.right_to_left,
        );

        combobox_properties.sorted =
            build_bool_property(&properties, BStr::new("Sorted"), combobox_properties.sorted);

        combobox_properties.style =
            build_property::<ComboBoxStyle>(&properties, BStr::new("Style"));

        combobox_properties.tab_index = build_i32_property(
            &properties,
            BStr::new("TabIndex"),
            combobox_properties.tab_index,
        );

        combobox_properties.tab_stop = build_bool_property(
            &properties,
            BStr::new("TabStop"),
            combobox_properties.tab_stop,
        );

        let text_key = BStr::new("Text");
        combobox_properties.text = properties
            .get(text_key)
            .unwrap_or(&combobox_properties.text);

        let tool_tip_text_key = BStr::new("ToolTipText");
        combobox_properties.tool_tip_text = properties
            .get(tool_tip_text_key)
            .unwrap_or(&combobox_properties.tool_tip_text);

        combobox_properties.top =
            build_i32_property(&properties, BStr::new("Top"), combobox_properties.top);

        combobox_properties.visible = build_bool_property(
            &properties,
            BStr::new("Visible"),
            combobox_properties.visible,
        );

        combobox_properties.whats_this_help_id = build_i32_property(
            &properties,
            BStr::new("WhatsThisHelp"),
            combobox_properties.whats_this_help_id,
        );

        combobox_properties.width =
            build_i32_property(&properties, BStr::new("Width"), combobox_properties.width);

        Ok(combobox_properties)
    }
}
