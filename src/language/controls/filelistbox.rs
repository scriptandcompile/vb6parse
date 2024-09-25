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
use serde::Serialize;

/// Properties for a `FileListBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::FileListBox`](crate::language::controls::VB6ControlKind::FileListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FileListBoxProperties<'a> {
    pub appearance: Appearance,
    pub archive: bool,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub hidden: bool,
    pub left: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub multi_select: MultiSelect,
    pub normal: bool,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub pattern: &'a BStr,
    pub read_only: bool,
    pub system: bool,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for FileListBoxProperties<'_> {
    fn default() -> Self {
        FileListBoxProperties {
            appearance: Appearance::ThreeD,
            archive: true,
            back_color: VB6Color::System { index: 5 },
            causes_validation: true,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::System { index: 8 },
            height: 1260,
            help_context_id: 0,
            hidden: false,
            left: 710,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            multi_select: MultiSelect::None,
            normal: true,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            pattern: BStr::new("*.*"),
            read_only: true,
            system: false,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: BStr::new(""),
            top: 480,
            visible: true,
            whats_this_help_id: 0,
            width: 735,
        }
    }
}

impl Serialize for FileListBoxProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("FileListBoxProperties", 27)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("archive", &self.archive)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("hidden", &self.hidden)?;
        s.serialize_field("left", &self.left)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("multi_select", &self.multi_select)?;
        s.serialize_field("normal", &self.normal)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("pattern", &self.pattern)?;
        s.serialize_field("read_only", &self.read_only)?;
        s.serialize_field("system", &self.system)?;
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

impl<'a> FileListBoxProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut file_list_box_properties = FileListBoxProperties::default();

        file_list_box_properties.appearance = build_property(properties, BStr::new("Appearance"));
        file_list_box_properties.archive = build_bool_property(
            properties,
            BStr::new("Archive"),
            file_list_box_properties.archive,
        );
        file_list_box_properties.back_color = build_color_property(
            properties,
            BStr::new("BackColor"),
            file_list_box_properties.back_color,
        );
        file_list_box_properties.causes_validation = build_bool_property(
            properties,
            BStr::new("CausesValidation"),
            file_list_box_properties.causes_validation,
        );
        file_list_box_properties.drag_mode = build_property(properties, BStr::new("DragMode"));
        file_list_box_properties.enabled = build_bool_property(
            properties,
            BStr::new("Enabled"),
            file_list_box_properties.enabled,
        );
        file_list_box_properties.fore_color = build_color_property(
            properties,
            BStr::new("ForeColor"),
            file_list_box_properties.fore_color,
        );
        file_list_box_properties.height = build_i32_property(
            properties,
            BStr::new("Height"),
            file_list_box_properties.height,
        );
        file_list_box_properties.help_context_id = build_i32_property(
            properties,
            BStr::new("HelpContextID"),
            file_list_box_properties.help_context_id,
        );
        file_list_box_properties.hidden = build_bool_property(
            properties,
            BStr::new("Hidden"),
            file_list_box_properties.hidden,
        );
        file_list_box_properties.left =
            build_i32_property(properties, BStr::new("Left"), file_list_box_properties.left);
        file_list_box_properties.mouse_pointer =
            build_property(properties, BStr::new("MousePointer"));
        file_list_box_properties.multi_select =
            build_property(properties, BStr::new("MultiSelect"));
        file_list_box_properties.normal = build_bool_property(
            properties,
            BStr::new("Normal"),
            file_list_box_properties.normal,
        );
        file_list_box_properties.ole_drag_mode =
            build_property(properties, BStr::new("OLEDragMode"));
        file_list_box_properties.ole_drop_mode =
            build_property(properties, BStr::new("OLEDropMode"));
        file_list_box_properties.pattern = properties
            .get(&BStr::new("Pattern"))
            .unwrap_or(&BStr::new("*.*"));
        file_list_box_properties.read_only = build_bool_property(
            properties,
            BStr::new("ReadOnly"),
            file_list_box_properties.read_only,
        );
        file_list_box_properties.system = build_bool_property(
            properties,
            BStr::new("System"),
            file_list_box_properties.system,
        );
        file_list_box_properties.tab_index = build_i32_property(
            properties,
            BStr::new("TabIndex"),
            file_list_box_properties.tab_index,
        );
        file_list_box_properties.tab_stop = build_bool_property(
            properties,
            BStr::new("TabStop"),
            file_list_box_properties.tab_stop,
        );
        file_list_box_properties.tool_tip_text = properties
            .get(&BStr::new("ToolTipText"))
            .unwrap_or(&BStr::new(""));
        file_list_box_properties.top =
            build_i32_property(properties, BStr::new("Top"), file_list_box_properties.top);
        file_list_box_properties.visible = build_bool_property(
            properties,
            BStr::new("Visible"),
            file_list_box_properties.visible,
        );
        file_list_box_properties.whats_this_help_id = build_i32_property(
            properties,
            BStr::new("WhatsThisHelpID"),
            file_list_box_properties.whats_this_help_id,
        );
        file_list_box_properties.width = build_i32_property(
            properties,
            BStr::new("Width"),
            file_list_box_properties.width,
        );

        Ok(file_list_box_properties)
    }
}
