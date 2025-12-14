use crate::language::controls::{
    Activation, CausesValidation, DragMode, MousePointer, ReferenceOrValue, TabStop, TextDirection,
    Visibility,
};
use crate::parsers::Properties;

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `ScrollBar` control.
///
/// This is used as an enum variant of
/// either a [`ControlKind::HScrollBar`](crate::language::controls::ControlKind::HScrollBar)
/// or a [`ControlKind::VScrollBar`](crate::language::controls::ControlKind::VScrollBar).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ScrollBarProperties {
    pub causes_validation: CausesValidation,
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    pub drag_mode: DragMode,
    pub enabled: Activation,
    pub height: i32,
    pub help_context_id: i32,
    pub large_change: i32,
    pub left: i32,
    pub max: i32,
    pub min: i32,
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    pub mouse_pointer: MousePointer,
    pub right_to_left: TextDirection,
    pub small_change: i32,
    pub tab_index: i32,
    pub tab_stop: TabStop,
    pub top: i32,
    pub value: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ScrollBarProperties {
    fn default() -> Self {
        ScrollBarProperties {
            causes_validation: CausesValidation::Yes,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            height: 30,
            help_context_id: 0,
            large_change: 1,
            left: 30,
            max: 32767,
            min: 0,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            right_to_left: TextDirection::LeftToRight,
            small_change: 1,
            tab_index: 0,
            tab_stop: TabStop::Included,
            top: 30,
            value: 0,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for ScrollBarProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("ScrollBarProperties", 20)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("large_change", &self.large_change)?;
        s.serialize_field("left", &self.left)?;
        s.serialize_field("max", &self.max)?;
        s.serialize_field("min", &self.min)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("small_change", &self.small_change)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("top", &self.top)?;

        s.serialize_field("value", &self.value)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl From<Properties> for ScrollBarProperties {
    fn from(prop: Properties) -> Self {
        let mut scroll_bar_prop = ScrollBarProperties::default();

        scroll_bar_prop.causes_validation =
            prop.get_property("CausesValidation", scroll_bar_prop.causes_validation);

        // DragIcon

        scroll_bar_prop.drag_mode = prop.get_property("DragMode", scroll_bar_prop.drag_mode);
        scroll_bar_prop.enabled = prop.get_property("Enabled", scroll_bar_prop.enabled);
        scroll_bar_prop.height = prop.get_i32("Height", scroll_bar_prop.height);
        scroll_bar_prop.help_context_id =
            prop.get_i32("HelpContextID", scroll_bar_prop.help_context_id);
        scroll_bar_prop.large_change = prop.get_i32("LargeChange", scroll_bar_prop.large_change);
        scroll_bar_prop.left = prop.get_i32("Left", scroll_bar_prop.left);
        scroll_bar_prop.max = prop.get_i32("Max", scroll_bar_prop.max);
        scroll_bar_prop.min = prop.get_i32("Min", scroll_bar_prop.min);

        // MouseIcon

        scroll_bar_prop.mouse_pointer =
            prop.get_property("MousePointer", scroll_bar_prop.mouse_pointer);
        scroll_bar_prop.right_to_left =
            prop.get_property("RightToLeft", scroll_bar_prop.right_to_left);
        scroll_bar_prop.small_change = prop.get_i32("SmallChange", scroll_bar_prop.small_change);
        scroll_bar_prop.tab_index = prop.get_i32("TabIndex", scroll_bar_prop.tab_index);
        scroll_bar_prop.tab_stop = prop.get_property("TabStop", scroll_bar_prop.tab_stop);
        scroll_bar_prop.top = prop.get_i32("Top", scroll_bar_prop.top);
        scroll_bar_prop.value = prop.get_i32("Value", scroll_bar_prop.value);
        scroll_bar_prop.visible = prop.get_property("Visible", scroll_bar_prop.visible);
        scroll_bar_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", scroll_bar_prop.whats_this_help_id);
        scroll_bar_prop.width = prop.get_i32("Width", scroll_bar_prop.width);

        scroll_bar_prop
    }
}
