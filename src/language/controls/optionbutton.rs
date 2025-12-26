//! Properties for an `OptionButton` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::OptionButton`](crate::language::controls::ControlKind::OptionButton).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use std::convert::{From, TryFrom};
use std::fmt::{Display, Formatter};
use std::str::FromStr;

use crate::errors::FormErrorKind;
use crate::language::controls::{
    Activation, Appearance, CausesValidation, DragMode, JustifyAlignment, MousePointer,
    OLEDropMode, ReferenceOrValue, Style, TabStop, TextDirection, UseMaskColor, Visibility,
};
use crate::language::{Color, VB_3D_FACE, VB_BUTTON_TEXT};
use crate::files::common::Properties;

use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// Value of an `OptionButton` control.
///
/// This is used as the `Value` property of an `OptionButton` control.
/// Either, `UnSelected` (0) or `Selected` (1).
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum OptionButtonValue {
    /// The option button is not selected.
    #[default]
    UnSelected = 0,
    /// The option button is selected.
    Selected = 1,
}

impl From<bool> for OptionButtonValue {
    fn from(value: bool) -> Self {
        if value {
            OptionButtonValue::Selected
        } else {
            OptionButtonValue::UnSelected
        }
    }
}

impl TryFrom<&str> for OptionButtonValue {
    type Error = FormErrorKind;

    fn try_from(value: &str) -> Result<Self, FormErrorKind> {
        match value {
            "0" => Ok(OptionButtonValue::UnSelected),
            "1" => Ok(OptionButtonValue::Selected),
            _ => Err(FormErrorKind::InvalidOptionButtonValue(value.to_string())),
        }
    }
}

impl FromStr for OptionButtonValue {
    type Err = crate::errors::FormErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        OptionButtonValue::try_from(s)
    }
}

impl Display for OptionButtonValue {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            OptionButtonValue::UnSelected => "UnSelected",
            OptionButtonValue::Selected => "Selected",
        };
        write!(f, "{text}")
    }
}

/// Properties for a `OptionButton` control.
///
/// This is used as an enum variant of
/// [`ControlKind::OptionButton`](crate::language::controls::ControlKind::OptionButton).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct OptionButtonProperties {
    /// Alignment of the option button.
    pub alignment: JustifyAlignment,
    /// Appearance of the option button.
    pub appearance: Appearance,
    /// Background color of the option button.
    pub back_color: Color,
    /// Caption of the option button.
    pub caption: String,
    /// Causes validation setting of the option button.
    pub causes_validation: CausesValidation,
    /// Disabled picture of the option button.
    pub disabled_picture: Option<ReferenceOrValue<DynamicImage>>,
    /// Down picture of the option button.
    pub down_picture: Option<ReferenceOrValue<DynamicImage>>,
    /// Drag icon of the option button.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Drag mode of the option button.
    pub drag_mode: DragMode,
    /// Enabled state of the option button.
    pub enabled: Activation,
    /// Foreground color of the option button.
    pub fore_color: Color,
    /// Height of the option button.
    pub height: i32,
    /// Help context ID of the option button.
    pub help_context_id: i32,
    /// Left position of the option button.
    pub left: i32,
    /// Mask color of the option button.
    pub mask_color: Color,
    /// Mouse icon of the option button.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Mouse pointer of the option button.
    pub mouse_pointer: MousePointer,
    /// OLE drop mode of the option button.
    pub ole_drop_mode: OLEDropMode,
    /// Picture of the option button.
    pub picture: Option<ReferenceOrValue<DynamicImage>>,
    /// Right-to-left text direction of the option button.
    pub right_to_left: TextDirection,
    /// Style of the option button.
    pub style: Style,
    /// Tab index of the option button.
    pub tab_index: i32,
    /// Tab stop setting of the option button.
    pub tab_stop: TabStop,
    /// Tool tip text of the option button.
    pub tool_tip_text: String,
    /// Top position of the option button.
    pub top: i32,
    /// Use mask color setting of the option button.
    pub use_mask_color: UseMaskColor,
    /// Value of the option button.
    pub value: OptionButtonValue,
    /// Visibility of the option button.
    pub visible: Visibility,
    /// What's this help ID of the option button.
    pub whats_this_help_id: i32,
    /// Width of the option button.
    pub width: i32,
}

impl Default for OptionButtonProperties {
    fn default() -> Self {
        OptionButtonProperties {
            alignment: JustifyAlignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB_3D_FACE,
            caption: String::new(),
            causes_validation: CausesValidation::Yes,
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
            tool_tip_text: String::new(),
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

impl From<Properties> for OptionButtonProperties {
    fn from(prop: Properties) -> Self {
        let mut option_button_prop = OptionButtonProperties::default();

        option_button_prop.alignment = prop.get_property("Alignment", option_button_prop.alignment);
        option_button_prop.appearance =
            prop.get_property("Appearance", option_button_prop.appearance);
        option_button_prop.back_color = prop.get_color("BackColor", option_button_prop.back_color);
        option_button_prop.caption = match prop.get("Caption") {
            Some(caption) => caption.into(),
            None => option_button_prop.caption,
        };
        option_button_prop.causes_validation =
            prop.get_property("CausesValidation", option_button_prop.causes_validation);

        // TODO: process DisabledPicture
        // DisabledPicture

        // TODO: process DownPicture
        // DownPicture

        // TODO: process DragIcon
        // DragIcon

        option_button_prop.drag_mode = prop.get_property("DragMode", option_button_prop.drag_mode);
        option_button_prop.enabled = prop.get_property("Enabled", option_button_prop.enabled);
        option_button_prop.fore_color = prop.get_color("ForeColor", option_button_prop.fore_color);
        option_button_prop.height = prop.get_i32("Height", option_button_prop.height);
        option_button_prop.help_context_id =
            prop.get_i32("HelpContextID", option_button_prop.help_context_id);
        option_button_prop.left = prop.get_i32("Left", option_button_prop.left);
        option_button_prop.mask_color = prop.get_color("MaskColor", option_button_prop.mask_color);

        // TODO: process MouseIcon
        // MouseIcon

        option_button_prop.mouse_pointer =
            prop.get_property("MousePointer", option_button_prop.mouse_pointer);
        option_button_prop.ole_drop_mode =
            prop.get_property("OLEDropMode", option_button_prop.ole_drop_mode);

        // TODO: process Picture
        // Picture

        option_button_prop.right_to_left =
            prop.get_property("RightToLeft", option_button_prop.right_to_left);
        option_button_prop.style = prop.get_property("Style", option_button_prop.style);
        option_button_prop.tab_index = prop.get_i32("TabIndex", option_button_prop.tab_index);
        option_button_prop.tab_stop = prop.get_property("TabStop", option_button_prop.tab_stop);
        option_button_prop.tool_tip_text = match prop.get("ToolTipText") {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => option_button_prop.tool_tip_text,
        };
        option_button_prop.top = prop.get_i32("Top", option_button_prop.top);
        option_button_prop.use_mask_color =
            prop.get_property("UseMaskColor", option_button_prop.use_mask_color);
        option_button_prop.value = prop.get_property("Value", option_button_prop.value);
        option_button_prop.visible = prop.get_property("Visible", option_button_prop.visible);
        option_button_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", option_button_prop.whats_this_help_id);
        option_button_prop.width = prop.get_i32("Width", option_button_prop.width);

        option_button_prop
    }
}
