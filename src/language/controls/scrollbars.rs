use std::collections::HashMap;
use std::hash::RandomState;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{DragMode, MousePointer};
use crate::parsers::form::{build_bool_property, build_i32_property, build_property};

use bstr::BStr;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `ScrollBar` control.
///
/// This is used as an enum variant of
/// either a [`VB6ControlKind::HScrollBar`](crate::language::controls::VB6ControlKind::HScrollBar)
/// or a [`VB6ControlKind::VScrollBar`](crate::language::controls::VB6ControlKind::VScrollBar).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct ScrollBarProperties {
    pub causes_validation: bool,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub large_change: i32,
    pub left: i32,
    pub max: i32,
    pub min: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub right_to_left: bool,
    pub small_change: i32,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub top: i32,
    pub value: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for ScrollBarProperties {
    fn default() -> Self {
        ScrollBarProperties {
            causes_validation: true,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            height: 30,
            help_context_id: 0,
            large_change: 1,
            left: 30,
            max: 32767,
            min: 0,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            right_to_left: false,
            small_change: 1,
            tab_index: 0,
            tab_stop: true,
            top: 30,
            value: 0,
            visible: true,
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

impl ScrollBarProperties {
    pub fn construct_control(properties: &HashMap<&BStr, &BStr>) -> Result<Self, VB6ErrorKind> {
        let mut scroll_bar_properties = ScrollBarProperties::default();

        scroll_bar_properties.causes_validation = build_bool_property(
            properties,
            BStr::new("CausesValidation"),
            scroll_bar_properties.causes_validation,
        );

        // DragIcon

        scroll_bar_properties.drag_mode =
            build_property::<DragMode, RandomState>(properties, BStr::new("DragMode"));
        scroll_bar_properties.enabled = build_bool_property(
            properties,
            BStr::new("Enabled"),
            scroll_bar_properties.enabled,
        );
        scroll_bar_properties.height = build_i32_property(
            properties,
            BStr::new("Height"),
            scroll_bar_properties.height,
        );
        scroll_bar_properties.help_context_id = build_i32_property(
            properties,
            BStr::new("HelpContextID"),
            scroll_bar_properties.help_context_id,
        );
        scroll_bar_properties.large_change = build_i32_property(
            properties,
            BStr::new("LargeChange"),
            scroll_bar_properties.large_change,
        );
        scroll_bar_properties.left =
            build_i32_property(properties, BStr::new("Left"), scroll_bar_properties.left);
        scroll_bar_properties.max =
            build_i32_property(properties, BStr::new("Max"), scroll_bar_properties.max);
        scroll_bar_properties.min =
            build_i32_property(properties, BStr::new("Min"), scroll_bar_properties.min);

        // MouseIcon

        scroll_bar_properties.mouse_pointer =
            build_property::<MousePointer, RandomState>(properties, BStr::new("MousePointer"));
        scroll_bar_properties.right_to_left = build_bool_property(
            properties,
            BStr::new("RightToLeft"),
            scroll_bar_properties.right_to_left,
        );
        scroll_bar_properties.small_change = build_i32_property(
            properties,
            BStr::new("SmallChange"),
            scroll_bar_properties.small_change,
        );
        scroll_bar_properties.tab_index = build_i32_property(
            properties,
            BStr::new("TabIndex"),
            scroll_bar_properties.tab_index,
        );
        scroll_bar_properties.tab_stop = build_bool_property(
            properties,
            BStr::new("TabStop"),
            scroll_bar_properties.tab_stop,
        );
        scroll_bar_properties.top =
            build_i32_property(properties, BStr::new("Top"), scroll_bar_properties.top);
        scroll_bar_properties.value =
            build_i32_property(properties, BStr::new("Value"), scroll_bar_properties.value);
        scroll_bar_properties.visible = build_bool_property(
            properties,
            BStr::new("Visible"),
            scroll_bar_properties.visible,
        );
        scroll_bar_properties.whats_this_help_id = build_i32_property(
            properties,
            BStr::new("WhatsThisHelpID"),
            scroll_bar_properties.whats_this_help_id,
        );
        scroll_bar_properties.width =
            build_i32_property(properties, BStr::new("Width"), scroll_bar_properties.width);

        Ok(scroll_bar_properties)
    }
}
