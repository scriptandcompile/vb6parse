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

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum OptionButtonValue {
    #[default]
    UnSelected = 0,
    Selected = 1,
}

/// Properties for a `OptionButton` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::OptionButton`](crate::language::controls::VB6ControlKind::OptionButton).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct OptionButtonProperties<'a> {
    pub alignment: JustifyAlignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub caption: &'a BStr,
    pub causes_validation: bool,
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
    pub value: OptionButtonValue,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for OptionButtonProperties<'_> {
    fn default() -> Self {
        OptionButtonProperties {
            alignment: JustifyAlignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            caption: BStr::new("Option1"),
            causes_validation: true,
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
            value: OptionButtonValue::UnSelected,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for OptionButtonProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("OptionButtonProperties", 29)?;
        s.serialize_field("alignment", &self.alignment)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("caption", &self.caption)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;

        let option_text = self.disabled_picture.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("disabled_picture", &option_text)?;

        let option_text = self.down_picture.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("down_picture", &option_text)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("left", &self.left)?;
        s.serialize_field("mask_color", &self.mask_color)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;

        let option_text = self.picture.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("picture", &option_text)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("style", &self.style)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("use_mask_color", &self.use_mask_color)?;
        s.serialize_field("value", &self.value)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl<'a> OptionButtonProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut option_button_properties = OptionButtonProperties::default();

        option_button_properties.alignment = build_property(properties, BStr::new("Alignment"));
        option_button_properties.appearance = build_property(properties, BStr::new("Appearance"));
        option_button_properties.back_color = build_color_property(
            properties,
            BStr::new("BackColor"),
            option_button_properties.back_color,
        );
        option_button_properties.caption = properties
            .get(&BStr::new("Caption"))
            .unwrap_or(&option_button_properties.caption);
        option_button_properties.causes_validation = build_bool_property(
            properties,
            BStr::new("CausesValidation"),
            option_button_properties.causes_validation,
        );

        // DisabledPicture
        // DownPicture
        // DragIcon

        option_button_properties.drag_mode = build_property(properties, BStr::new("DragMode"));
        option_button_properties.enabled = build_bool_property(
            properties,
            BStr::new("Enabled"),
            option_button_properties.enabled,
        );
        option_button_properties.fore_color = build_color_property(
            properties,
            BStr::new("ForeColor"),
            option_button_properties.fore_color,
        );
        option_button_properties.height = build_i32_property(
            properties,
            BStr::new("Height"),
            option_button_properties.height,
        );
        option_button_properties.help_context_id = build_i32_property(
            properties,
            BStr::new("HelpContextID"),
            option_button_properties.help_context_id,
        );
        option_button_properties.left =
            build_i32_property(properties, BStr::new("Left"), option_button_properties.left);
        option_button_properties.mask_color = build_color_property(
            properties,
            BStr::new("MaskColor"),
            option_button_properties.mask_color,
        );

        // MouseIcon

        option_button_properties.mouse_pointer =
            build_property(properties, BStr::new("MousePointer"));
        option_button_properties.ole_drop_mode =
            build_property(properties, BStr::new("OLEDropMode"));

        // Picture

        option_button_properties.right_to_left = build_bool_property(
            properties,
            BStr::new("RightToLeft"),
            option_button_properties.right_to_left,
        );
        option_button_properties.style = build_property(properties, BStr::new("Style"));
        option_button_properties.tab_index = build_i32_property(
            properties,
            BStr::new("TabIndex"),
            option_button_properties.tab_index,
        );
        option_button_properties.tab_stop = build_bool_property(
            properties,
            BStr::new("TabStop"),
            option_button_properties.tab_stop,
        );
        option_button_properties.tool_tip_text = properties
            .get(&BStr::new("ToolTipText"))
            .unwrap_or(&option_button_properties.tool_tip_text);
        option_button_properties.top =
            build_i32_property(properties, BStr::new("Top"), option_button_properties.top);
        option_button_properties.use_mask_color = build_bool_property(
            properties,
            BStr::new("UseMaskColor"),
            option_button_properties.use_mask_color,
        );
        option_button_properties.value = build_property(properties, BStr::new("Value"));
        option_button_properties.visible = build_bool_property(
            properties,
            BStr::new("Visible"),
            option_button_properties.visible,
        );
        option_button_properties.whats_this_help_id = build_i32_property(
            properties,
            BStr::new("WhatsThisHelpID"),
            option_button_properties.whats_this_help_id,
        );
        option_button_properties.width = build_i32_property(
            properties,
            BStr::new("Width"),
            option_button_properties.width,
        );

        Ok(option_button_properties)
    }
}
