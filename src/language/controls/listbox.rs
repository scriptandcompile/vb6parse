use crate::{
    language::{
        color::{Color, VB_BUTTON_FACE, VB_BUTTON_TEXT},
        controls::{
            Activation, Appearance, CausesValidation, DragMode, MousePointer, MultiSelect,
            OLEDragMode, OLEDropMode, ReferenceOrValue, TabStop, TextDirection, Visibility,
        },
    },
    parsers::Properties,
};

use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// `ListBox` control style.
///
/// The `ListBoxStyle` enum represents the different styles that a `ListBox` control can have.
/// The `Standard` style is the default style, while the `Checkbox` style adds a check box
/// next to each item in the list box.
///
/// This enum is used in the `ListBoxProperties` struct to specify the style of the `ListBox` control.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum ListBoxStyle {
    /// Standard list box.
    ///
    /// This is the default style.
    #[default]
    Standard = 0,
    /// List box with a check box next to each item.
    Checkbox = 1,
}

/// Properties for a `ListBox` control.
///
/// This is used as an enum variant of
/// [`ControlKind::ListBox`](crate::language::controls::ControlKind::ListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ListBoxProperties {
    pub appearance: Appearance,
    pub back_color: Color,
    pub causes_validation: CausesValidation,
    pub columns: i32,
    pub data_field: String,
    pub data_format: String,
    pub data_member: String,
    pub data_source: String,
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    pub drag_mode: DragMode,
    pub enabled: Activation,
    pub fore_color: Color,
    pub height: i32,
    pub help_context_id: i32,
    pub integral_height: bool,
    // pub item_data: Vec<String>,
    pub left: i32,
    // pub list: Vec<String>,
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    pub mouse_pointer: MousePointer,
    pub multi_select: MultiSelect,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: TextDirection,
    pub sorted: bool,
    pub style: ListBoxStyle,
    pub tab_index: i32,
    pub tab_stop: TabStop,
    pub tool_tip_text: String,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ListBoxProperties {
    fn default() -> Self {
        ListBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB_BUTTON_FACE,
            causes_validation: CausesValidation::Yes,
            columns: 0,
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            fore_color: VB_BUTTON_TEXT,
            height: 30,
            help_context_id: 0,
            integral_height: true,
            left: 30,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            multi_select: MultiSelect::None,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: TextDirection::LeftToRight,
            sorted: false,
            style: ListBoxStyle::Standard,
            tab_index: 0,
            tab_stop: TabStop::Included,
            tool_tip_text: "".into(),
            top: 30,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for ListBoxProperties {
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

impl From<Properties> for ListBoxProperties {
    fn from(prop: Properties) -> Self {
        let mut list_box_prop = ListBoxProperties::default();

        list_box_prop.appearance = prop.get_property("Appearance", list_box_prop.appearance);
        list_box_prop.back_color = prop.get_color("BackColor", list_box_prop.back_color);
        list_box_prop.causes_validation =
            prop.get_property("CausesValidation", list_box_prop.causes_validation);
        list_box_prop.columns = prop.get_i32("Columns", list_box_prop.columns);
        list_box_prop.data_field = match prop.get("DataField") {
            Some(data_field) => data_field.into(),
            None => list_box_prop.data_field,
        };
        list_box_prop.data_format = match prop.get("DataFormat") {
            Some(data_format) => data_format.into(),
            None => list_box_prop.data_format,
        };
        list_box_prop.data_member = match prop.get("DataMember") {
            Some(data_member) => data_member.into(),
            None => list_box_prop.data_member,
        };
        list_box_prop.data_source = match prop.get("DataSource") {
            Some(data_source) => data_source.into(),
            None => list_box_prop.data_source,
        };

        // DragIcon

        list_box_prop.drag_mode = prop.get_property("DragMode", list_box_prop.drag_mode);
        list_box_prop.enabled = prop.get_property("Enabled", list_box_prop.enabled);
        list_box_prop.fore_color = prop.get_color("ForeColor", list_box_prop.fore_color);
        list_box_prop.height = prop.get_i32("Height", list_box_prop.height);
        list_box_prop.help_context_id =
            prop.get_i32("HelpContextID", list_box_prop.help_context_id);
        list_box_prop.integral_height =
            prop.get_bool("IntegralHeight", list_box_prop.integral_height);
        list_box_prop.left = prop.get_i32("Left", list_box_prop.left);

        // MouseIcon

        list_box_prop.mouse_pointer =
            prop.get_property("MousePointer", list_box_prop.mouse_pointer);
        list_box_prop.multi_select = prop.get_property("MultiSelect", list_box_prop.multi_select);
        list_box_prop.ole_drag_mode = prop.get_property("OLEDragMode", list_box_prop.ole_drag_mode);
        list_box_prop.ole_drop_mode = prop.get_property("OLEDropMode", list_box_prop.ole_drop_mode);
        list_box_prop.right_to_left = prop.get_property("RightToLeft", list_box_prop.right_to_left);
        list_box_prop.sorted = prop.get_bool("Sorted", list_box_prop.sorted);
        list_box_prop.style = prop.get_property("Style", list_box_prop.style);
        list_box_prop.tab_index = prop.get_i32("TabIndex", list_box_prop.tab_index);
        list_box_prop.tab_stop = prop.get_property("TabStop", list_box_prop.tab_stop);
        list_box_prop.tool_tip_text = match prop.get("ToolTipText") {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => list_box_prop.tool_tip_text,
        };
        list_box_prop.top = prop.get_i32("Top", list_box_prop.top);
        list_box_prop.visible = prop.get_property("Visible", list_box_prop.visible);
        list_box_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", list_box_prop.whats_this_help_id);
        list_box_prop.width = prop.get_i32("Width", list_box_prop.width);

        list_box_prop
    }
}
