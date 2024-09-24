use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::color::VB6Color;
use crate::language::controls::{
    Appearance, DragMode, MousePointer, MultiSelect, OLEDragMode, OLEDropMode,
};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};

use bstr::BStr;
use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum ListBoxStyle {
    #[default]
    Standard = 0,
    Checkbox = 1,
}

/// Properties for a `ListBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::ListBox`](crate::language::controls::VB6ControlKind::ListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ListBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    pub columns: i32,
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
    // pub item_data: Vec<&'a BStr>,
    pub left: i32,
    // pub list: Vec<&'a BStr>,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub multi_select: MultiSelect,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub sorted: bool,
    pub style: ListBoxStyle,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ListBoxProperties<'_> {
    fn default() -> Self {
        ListBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            causes_validation: true,
            columns: 0,
            data_field: BStr::new(""),
            data_format: BStr::new(""),
            data_member: BStr::new(""),
            data_source: BStr::new(""),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            integral_height: true,
            left: 30,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            multi_select: MultiSelect::None,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: false,
            sorted: false,
            style: ListBoxStyle::Standard,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: BStr::new(""),
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for ListBoxProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("ListBoxProperties", 31)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;
        s.serialize_field("columns", &self.columns)?;
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
        s.serialize_field("left", &self.left)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("multi_select", &self.multi_select)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("sorted", &self.sorted)?;
        s.serialize_field("style", &self.style)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl<'a> ListBoxProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut list_box_properties = ListBoxProperties::default();

        list_box_properties.appearance =
            build_property::<Appearance>(properties, BStr::new("Appearance"));
        list_box_properties.back_color = build_color_property(
            properties,
            BStr::new("BackColor"),
            list_box_properties.back_color,
        );
        list_box_properties.causes_validation = build_bool_property(
            properties,
            BStr::new("CausesValidation"),
            list_box_properties.causes_validation,
        );
        list_box_properties.columns = build_i32_property(
            properties,
            BStr::new("Columns"),
            list_box_properties.columns,
        );
        list_box_properties.data_field = properties
            .get(BStr::new("DataField"))
            .unwrap_or(&list_box_properties.data_field);
        list_box_properties.data_format = properties
            .get(BStr::new("DataFormat"))
            .unwrap_or(&list_box_properties.data_format);
        list_box_properties.data_member = properties
            .get(BStr::new("DataMember"))
            .unwrap_or(&list_box_properties.data_member);
        list_box_properties.data_source = properties
            .get(BStr::new("DataSource"))
            .unwrap_or(&list_box_properties.data_source);

        // DragIcon

        list_box_properties.drag_mode =
            build_property::<DragMode>(properties, BStr::new("DragMode"));
        list_box_properties.enabled = build_bool_property(
            properties,
            BStr::new("Enabled"),
            list_box_properties.enabled,
        );
        list_box_properties.fore_color = build_color_property(
            properties,
            BStr::new("ForeColor"),
            list_box_properties.fore_color,
        );
        list_box_properties.height =
            build_i32_property(properties, BStr::new("Height"), list_box_properties.height);
        list_box_properties.help_context_id = build_i32_property(
            properties,
            BStr::new("HelpContextID"),
            list_box_properties.help_context_id,
        );
        list_box_properties.integral_height = build_bool_property(
            properties,
            BStr::new("IntegralHeight"),
            list_box_properties.integral_height,
        );
        list_box_properties.left =
            build_i32_property(properties, BStr::new("Left"), list_box_properties.left);

        // MouseIcon

        list_box_properties.mouse_pointer =
            build_property::<MousePointer>(properties, BStr::new("MousePointer"));
        list_box_properties.multi_select =
            build_property::<MultiSelect>(properties, BStr::new("MultiSelect"));
        list_box_properties.ole_drag_mode =
            build_property::<OLEDragMode>(properties, BStr::new("OLEDragMode"));
        list_box_properties.ole_drop_mode =
            build_property::<OLEDropMode>(properties, BStr::new("OLEDropMode"));
        list_box_properties.right_to_left = build_bool_property(
            properties,
            BStr::new("RightToLeft"),
            list_box_properties.right_to_left,
        );
        list_box_properties.sorted =
            build_bool_property(properties, BStr::new("Sorted"), list_box_properties.sorted);
        list_box_properties.style = build_property::<ListBoxStyle>(properties, BStr::new("Style"));
        list_box_properties.tab_index = build_i32_property(
            properties,
            BStr::new("TabIndex"),
            list_box_properties.tab_index,
        );
        list_box_properties.tab_stop = build_bool_property(
            properties,
            BStr::new("TabStop"),
            list_box_properties.tab_stop,
        );
        list_box_properties.tool_tip_text = properties
            .get(BStr::new("ToolTipText"))
            .unwrap_or(&list_box_properties.tool_tip_text);
        list_box_properties.top =
            build_i32_property(properties, BStr::new("Top"), list_box_properties.top);
        list_box_properties.visible = build_bool_property(
            properties,
            BStr::new("Visible"),
            list_box_properties.visible,
        );
        list_box_properties.whats_this_help_id = build_i32_property(
            properties,
            BStr::new("WhatsThisHelpID"),
            list_box_properties.whats_this_help_id,
        );
        list_box_properties.width =
            build_i32_property(properties, BStr::new("Width"), list_box_properties.width);

        Ok(list_box_properties)
    }
}
