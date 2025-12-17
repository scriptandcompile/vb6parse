//! Defines the properties and value enumeration for a CheckBox control in a VB6 form.
//! This includes the `CheckBoxProperties` struct which holds all configurable
//! properties of the CheckBox, as well as the `CheckBoxValue` enum which
//! represents the state of the CheckBox (Unchecked, Checked, Grayed).
//! These are used in the context of parsing and representing VB6 form controls.

use crate::{
    language::{
        controls::{
            Activation, Appearance, CausesValidation, DragMode, JustifyAlignment, MousePointer,
            OLEDropMode, ReferenceOrValue, Style, TabStop, TextDirection, UseMaskColor, Visibility,
        },
        Color, VB_BUTTON_FACE, VB_BUTTON_TEXT,
    },
    parsers::Properties,
};

use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// Represents the current state of a checkbox control.
///
/// This is used as a property of the [`CheckBoxProperties`](crate::language::controls::CheckBoxProperties)
/// struct.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum CheckBoxValue {
    /// The checkbox is unchecked.
    ///
    /// This is the default value.
    #[default]
    Unchecked = 0,
    /// The checkbox is checked.
    Checked = 1,
    /// The checkbox is grayed out and cannot be checked or unchecked.
    Grayed = 2,
}

/// Properties for a `CheckBox` control.
///
/// This is used as an enum variant of
/// [`ControlKind::CheckBox`](crate::language::controls::ControlKind::CheckBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct CheckBoxProperties {
    /// Justify alignment of the checkbox caption.
    pub alignment: JustifyAlignment,
    /// Appearance of the checkbox control.
    pub appearance: Appearance,
    /// Background color of the checkbox control.
    pub back_color: Color,
    /// Caption text of the checkbox control.
    pub caption: String,
    /// Whether the checkbox control causes validation.
    pub causes_validation: CausesValidation,
    /// Data field associated with the checkbox control.
    pub data_field: String,
    /// Data format for the checkbox control.
    pub data_format: String,
    /// Data member associated with the checkbox control.
    pub data_member: String,
    /// Data source associated with the checkbox control.
    pub data_source: String,
    /// Picture displayed when the checkbox is disabled.
    pub disabled_picture: Option<ReferenceOrValue<DynamicImage>>,
    /// Picture displayed when the checkbox is pressed down.
    pub down_picture: Option<ReferenceOrValue<DynamicImage>>,
    /// Icon used during drag operations.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Drag mode of the checkbox control.
    pub drag_mode: DragMode,
    /// Whether the checkbox control is enabled.
    pub enabled: Activation,
    /// Foreground color of the checkbox control.
    pub fore_color: Color,
    /// Height of the checkbox control.
    pub height: i32,
    /// Help context ID associated with the checkbox control.
    pub help_context_id: i32,
    /// Left position of the checkbox control.
    pub left: i32,
    /// Mask color used for transparency.
    pub mask_color: Color,
    /// Icon displayed when the mouse is over the checkbox control.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Mouse pointer style when hovering over the checkbox control.
    pub mouse_pointer: MousePointer,
    /// OLE drop mode of the checkbox control.
    pub ole_drop_mode: OLEDropMode,
    /// Picture displayed on the checkbox control.
    pub picture: Option<ReferenceOrValue<DynamicImage>>,
    /// Text direction of the checkbox control.
    pub right_to_left: TextDirection,
    /// Style of the checkbox control.
    pub style: Style,
    /// Tab index of the checkbox control.
    pub tab_index: i32,
    /// Whether the checkbox control is included in the tab order.
    pub tab_stop: TabStop,
    /// Tool tip text for the checkbox control.
    pub tool_tip_text: String,
    /// Top position of the checkbox control.
    pub top: i32,
    /// Whether to use the mask color for transparency.
    pub use_mask_color: UseMaskColor,
    /// Current value/state of the checkbox control.
    pub value: CheckBoxValue,
    /// Visibility of the checkbox control.
    pub visible: Visibility,
    /// "What's This?" help ID associated with the checkbox control.
    pub whats_this_help_id: i32,
    /// Width of the checkbox control.
    pub width: i32,
}

impl Default for CheckBoxProperties {
    fn default() -> Self {
        CheckBoxProperties {
            alignment: JustifyAlignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB_BUTTON_FACE,
            caption: "".into(),
            causes_validation: CausesValidation::Yes,
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            disabled_picture: None,
            down_picture: None,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            fore_color: VB_BUTTON_TEXT,
            height: 30,
            help_context_id: 0,
            left: 30,
            mask_color: Color::new(0xC0, 0xC0, 0xC0),
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
            use_mask_color: UseMaskColor::DoNotUseMaskColor,
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

impl From<Properties> for CheckBoxProperties {
    fn from(prop: Properties) -> Self {
        let mut checkbox_prop = CheckBoxProperties::default();

        checkbox_prop.alignment = prop.get_property("Alignment", checkbox_prop.alignment);
        checkbox_prop.appearance = prop.get_property("Appearance", checkbox_prop.appearance);
        checkbox_prop.back_color = prop.get_color("BackColor", checkbox_prop.back_color);
        checkbox_prop.caption = match prop.get("Caption".into()) {
            Some(caption) => caption.into(),
            None => checkbox_prop.caption,
        };
        checkbox_prop.causes_validation =
            prop.get_property("CausesValidation", checkbox_prop.causes_validation);
        checkbox_prop.data_field = match prop.get("DataField") {
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

        checkbox_prop.drag_mode = prop.get_property("DragMode", checkbox_prop.drag_mode);
        checkbox_prop.enabled = prop.get_property("Enabled", checkbox_prop.enabled);
        checkbox_prop.fore_color = prop.get_color("ForeColor", checkbox_prop.fore_color);
        checkbox_prop.height = prop.get_i32("Height", checkbox_prop.height);
        checkbox_prop.help_context_id =
            prop.get_i32("HelpContextID", checkbox_prop.help_context_id);
        checkbox_prop.left = prop.get_i32("Left", checkbox_prop.left);
        checkbox_prop.mask_color = prop.get_color("MaskColor", checkbox_prop.mask_color);

        //MouseIcon

        checkbox_prop.mouse_pointer =
            prop.get_property("MousePointer", checkbox_prop.mouse_pointer);
        checkbox_prop.ole_drop_mode = prop.get_property("OLEDropMode", checkbox_prop.ole_drop_mode);

        //Picture

        checkbox_prop.right_to_left = prop.get_property("RightToLeft", checkbox_prop.right_to_left);
        checkbox_prop.style = prop.get_property("Style", checkbox_prop.style);
        checkbox_prop.tab_index = prop.get_i32("TabIndex", checkbox_prop.tab_index);
        checkbox_prop.tab_stop = prop.get_property("TabStop", checkbox_prop.tab_stop);
        checkbox_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => checkbox_prop.tool_tip_text,
        };
        checkbox_prop.top = prop.get_i32("Top", checkbox_prop.top);
        checkbox_prop.use_mask_color =
            prop.get_property("UseMaskColor", checkbox_prop.use_mask_color);
        checkbox_prop.value = prop.get_property("Value", checkbox_prop.value);
        checkbox_prop.visible = prop.get_property("Visible", checkbox_prop.visible);
        checkbox_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelp", checkbox_prop.whats_this_help_id);
        checkbox_prop.width = prop.get_i32("Width", checkbox_prop.width);

        checkbox_prop
    }
}
