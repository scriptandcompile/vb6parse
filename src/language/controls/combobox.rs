use crate::language::controls::{
    Activation, Appearance, CausesValidation, DragMode, MousePointer, OLEDragMode, OLEDropMode,
    TabStop, TextDirection, Visibility,
};
use crate::parsers::Properties;
use crate::VB6Color;

use bstr::{BStr, BString};
use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// The `ComboBoxStyle` enum represents the different styles of a `ComboBox` control.
/// It can be either a drop-down list, a simple list, or a drop-down
/// list with a text box.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum ComboBoxStyle {
    /// A drop-down combo box that allows the user to select an item from a list
    /// or enter a new value.
    ///
    /// This is the default style.
    #[default]
    DropDownCombo = 0,
    /// A simple combo box that allows the user to select an item from a list
    /// but does not allow the user to enter a new value.
    SimpleCombo = 1,
    /// A drop-down list that allows the user to select an item from a list
    /// but does not allow the user to enter a new value.
    DropDownList = 2,
}

/// Properties for a `ComboBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::ComboBox`](crate::language::controls::VB6ControlKind::ComboBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ComboBoxProperties {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: CausesValidation,
    pub data_field: BString,
    pub data_format: BString,
    pub data_member: BString,
    pub data_source: BString,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: Activation,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub integral_height: bool,
    //pub item_data: Vec<BString>,
    pub left: i32,
    // pub list: Vec<BString>,
    pub locked: bool,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: TextDirection,
    pub sorted: bool,
    pub style: ComboBoxStyle,
    pub tab_index: i32,
    pub tab_stop: TabStop,
    pub text: BString,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ComboBoxProperties {
    fn default() -> Self {
        ComboBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H80000005&").unwrap(),
            causes_validation: CausesValidation::Yes,
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
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
            right_to_left: TextDirection::LeftToRight,
            sorted: false,
            style: ComboBoxStyle::DropDownCombo,
            tab_index: 0,
            tab_stop: TabStop::Included,
            text: "".into(),
            tool_tip_text: "".into(),
            top: 30,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for ComboBoxProperties {
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

impl<'a> From<Properties<'a>> for ComboBoxProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut combobox_prop = ComboBoxProperties::default();

        combobox_prop.appearance =
            prop.get_property(b"Appearance".into(), combobox_prop.appearance);
        combobox_prop.back_color = prop.get_color(b"BackColor".into(), combobox_prop.back_color);
        combobox_prop.causes_validation =
            prop.get_property(b"CausesValidation".into(), combobox_prop.causes_validation);
        combobox_prop.data_field = match prop.get(b"DataField".into()) {
            Some(data_field) => data_field.into(),
            None => "".into(),
        };
        combobox_prop.data_format = match prop.get(b"DataFormat".into()) {
            Some(data_format) => data_format.into(),
            None => "".into(),
        };
        combobox_prop.data_member = match prop.get(b"DataMember".into()) {
            Some(data_member) => data_member.into(),
            None => "".into(),
        };
        combobox_prop.data_source = match prop.get(b"DataSource".into()) {
            Some(data_source) => data_source.into(),
            None => "".into(),
        };

        // drag_icon

        combobox_prop.drag_mode = prop.get_property(b"DragMode".into(), combobox_prop.drag_mode);
        combobox_prop.enabled = prop.get_property(b"Enabled".into(), combobox_prop.enabled);
        combobox_prop.fore_color = prop.get_color(b"ForeColor".into(), combobox_prop.fore_color);
        combobox_prop.height = prop.get_i32(b"Height".into(), combobox_prop.height);
        combobox_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), combobox_prop.help_context_id);
        combobox_prop.integral_height =
            prop.get_bool(b"IntegralHeight".into(), combobox_prop.integral_height);
        combobox_prop.left = prop.get_i32(b"Left".into(), combobox_prop.left);
        combobox_prop.locked = prop.get_bool(b"Locked".into(), combobox_prop.locked);

        // mouse_icon

        combobox_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), combobox_prop.mouse_pointer);
        combobox_prop.ole_drag_mode =
            prop.get_property(b"OLEDragMode".into(), combobox_prop.ole_drag_mode);
        combobox_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), combobox_prop.ole_drop_mode);
        combobox_prop.right_to_left =
            prop.get_property(b"RightToLeft".into(), combobox_prop.right_to_left);
        combobox_prop.sorted = prop.get_bool(b"Sorted".into(), combobox_prop.sorted);
        combobox_prop.style = prop.get_property(b"Style".into(), combobox_prop.style);
        combobox_prop.tab_index = prop.get_i32(b"TabIndex".into(), combobox_prop.tab_index);
        combobox_prop.tab_stop = prop.get_property(b"TabStop".into(), combobox_prop.tab_stop);
        combobox_prop.text = match prop.get(BStr::new("Text")) {
            Some(text) => text.into(),
            None => combobox_prop.text,
        };
        combobox_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => combobox_prop.tool_tip_text,
        };
        combobox_prop.top = prop.get_i32(b"Top".into(), combobox_prop.top);
        combobox_prop.visible = prop.get_property(b"Visible".into(), combobox_prop.visible);
        combobox_prop.whats_this_help_id =
            prop.get_i32(b"WhatsThisHelp".into(), combobox_prop.whats_this_help_id);
        combobox_prop.width = prop.get_i32(b"Width".into(), combobox_prop.width);

        combobox_prop
    }
}
