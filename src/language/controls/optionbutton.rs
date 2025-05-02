use crate::language::controls::{
    Activation, Appearance, CausesValidation, DragMode, JustifyAlignment, MousePointer,
    OLEDropMode, Style, TabStop, TextDirection, UseMaskColor, Visibility,
};
use crate::language::VB6Color;
use crate::parsers::Properties;

use bstr::BString;
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
pub struct OptionButtonProperties {
    pub alignment: JustifyAlignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub caption: BString,
    pub causes_validation: CausesValidation,
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
    pub use_mask_color: UseMaskColor,
    pub value: OptionButtonValue,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for OptionButtonProperties {
    fn default() -> Self {
        OptionButtonProperties {
            alignment: JustifyAlignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            caption: "".into(),
            causes_validation: CausesValidation::Yes,
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
            use_mask_color: UseMaskColor::DoNotUseMaskColor,
            value: OptionButtonValue::UnSelected,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for OptionButtonProperties {
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

impl<'a> From<Properties<'a>> for OptionButtonProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut option_button_prop = OptionButtonProperties::default();

        option_button_prop.alignment =
            prop.get_property(b"Alignment".into(), option_button_prop.alignment);
        option_button_prop.appearance =
            prop.get_property(b"Appearance".into(), option_button_prop.appearance);
        option_button_prop.back_color =
            prop.get_color(b"BackColor".into(), option_button_prop.back_color);
        option_button_prop.caption = match prop.get(b"Caption".into()) {
            Some(caption) => caption.into(),
            None => option_button_prop.caption,
        };
        option_button_prop.causes_validation = prop.get_property(
            b"CausesValidation".into(),
            option_button_prop.causes_validation,
        );

        // DisabledPicture
        // DownPicture
        // DragIcon

        option_button_prop.drag_mode =
            prop.get_property(b"DragMode".into(), option_button_prop.drag_mode);
        option_button_prop.enabled =
            prop.get_property(b"Enabled".into(), option_button_prop.enabled);
        option_button_prop.fore_color =
            prop.get_color(b"ForeColor".into(), option_button_prop.fore_color);
        option_button_prop.height = prop.get_i32(b"Height".into(), option_button_prop.height);
        option_button_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), option_button_prop.help_context_id);
        option_button_prop.left = prop.get_i32(b"Left".into(), option_button_prop.left);
        option_button_prop.mask_color =
            prop.get_color(b"MaskColor".into(), option_button_prop.mask_color);

        // MouseIcon

        option_button_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), option_button_prop.mouse_pointer);
        option_button_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), option_button_prop.ole_drop_mode);

        // Picture

        option_button_prop.right_to_left =
            prop.get_property(b"RightToLeft".into(), option_button_prop.right_to_left);
        option_button_prop.style = prop.get_property(b"Style".into(), option_button_prop.style);
        option_button_prop.tab_index =
            prop.get_i32(b"TabIndex".into(), option_button_prop.tab_index);
        option_button_prop.tab_stop =
            prop.get_property(b"TabStop".into(), option_button_prop.tab_stop);
        option_button_prop.tool_tip_text = match prop.get(b"ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => option_button_prop.tool_tip_text,
        };
        option_button_prop.top = prop.get_i32(b"Top".into(), option_button_prop.top);
        option_button_prop.use_mask_color =
            prop.get_property(b"UseMaskColor".into(), option_button_prop.use_mask_color);
        option_button_prop.value = prop.get_property(b"Value".into(), option_button_prop.value);
        option_button_prop.visible =
            prop.get_property(b"Visible".into(), option_button_prop.visible);
        option_button_prop.whats_this_help_id = prop.get_i32(
            b"WhatsThisHelpID".into(),
            option_button_prop.whats_this_help_id,
        );
        option_button_prop.width = prop.get_i32(b"Width".into(), option_button_prop.width);

        option_button_prop
    }
}
