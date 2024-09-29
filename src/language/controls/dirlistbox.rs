use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{Appearance, DragMode, MousePointer, OLEDragMode, OLEDropMode};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};
use crate::VB6Color;

use bstr::BStr;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `DirListBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::DirListBox`](crate::language::controls::VB6ControlKind::DirListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct DirListBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for DirListBoxProperties<'_> {
    fn default() -> Self {
        DirListBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::System { index: 5 },
            causes_validation: true,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::System { index: 8 },
            height: 3195,
            help_context_id: 0,
            left: 720,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: BStr::new(""),
            top: 720,
            visible: true,
            whats_this_help_id: 0,
            width: 975,
        }
    }
}

impl Serialize for DirListBoxProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("DirListBoxProperties", 20)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("left", &self.left)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
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

impl<'a> DirListBoxProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut dir_list_box_properties = DirListBoxProperties::default();

        dir_list_box_properties.appearance = build_property(properties, b"Appearance");
        dir_list_box_properties.back_color =
            build_color_property(properties, b"BackColor", dir_list_box_properties.back_color);
        dir_list_box_properties.causes_validation = build_bool_property(
            properties,
            b"CausesValidation",
            dir_list_box_properties.causes_validation,
        );

        // DragIcon

        dir_list_box_properties.drag_mode = build_property(properties, b"DragMode");
        dir_list_box_properties.enabled =
            build_bool_property(properties, b"Enabled", dir_list_box_properties.enabled);
        dir_list_box_properties.fore_color =
            build_color_property(properties, b"ForeColor", dir_list_box_properties.fore_color);
        dir_list_box_properties.height =
            build_i32_property(properties, b"Height", dir_list_box_properties.height);
        dir_list_box_properties.help_context_id = build_i32_property(
            properties,
            b"HelpContextID",
            dir_list_box_properties.help_context_id,
        );
        dir_list_box_properties.left =
            build_i32_property(properties, b"Left", dir_list_box_properties.left);
        dir_list_box_properties.mouse_pointer = build_property(properties, b"MousePointer");
        dir_list_box_properties.ole_drag_mode = build_property(properties, b"OLEDragMode");
        dir_list_box_properties.ole_drop_mode = build_property(properties, b"OLEDropMode");
        dir_list_box_properties.tab_index =
            build_i32_property(properties, b"TabIndex", dir_list_box_properties.tab_index);
        dir_list_box_properties.tab_stop =
            build_bool_property(properties, b"TabStop", dir_list_box_properties.tab_stop);
        dir_list_box_properties.tool_tip_text = properties
            .get(BStr::new("ToolTipText"))
            .unwrap_or(&BStr::new(""));
        dir_list_box_properties.top =
            build_i32_property(properties, b"Top", dir_list_box_properties.top);
        dir_list_box_properties.visible =
            build_bool_property(properties, b"Visible", dir_list_box_properties.visible);
        dir_list_box_properties.whats_this_help_id = build_i32_property(
            properties,
            b"WhatsThisHelpID",
            dir_list_box_properties.whats_this_help_id,
        );
        dir_list_box_properties.width =
            build_i32_property(properties, b"Width", dir_list_box_properties.width);

        Ok(dir_list_box_properties)
    }
}
