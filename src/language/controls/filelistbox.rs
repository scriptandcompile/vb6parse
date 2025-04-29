use crate::language::color::VB6Color;
use crate::language::controls::{
    Appearance, DragMode, MousePointer, MultiSelect, OLEDragMode, OLEDropMode, Visibility,
};
use crate::parsers::Properties;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `FileListBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::FileListBox`](crate::language::controls::VB6ControlKind::FileListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FileListBoxProperties {
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
    pub pattern: BString,
    pub read_only: bool,
    pub system: bool,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for FileListBoxProperties {
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
            pattern: "*.*".into(), // Default pattern for file selection
            read_only: true,
            system: false,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: "".into(),
            top: 480,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 735,
        }
    }
}

impl Serialize for FileListBoxProperties {
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

impl<'a> From<Properties<'a>> for FileListBoxProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut file_list_box_prop = FileListBoxProperties::default();

        file_list_box_prop.appearance =
            prop.get_property(b"Appearance".into(), file_list_box_prop.appearance);
        file_list_box_prop.archive = prop.get_bool(b"Archive".into(), file_list_box_prop.archive);
        file_list_box_prop.back_color =
            prop.get_color(b"BackColor".into(), file_list_box_prop.back_color);
        file_list_box_prop.causes_validation = prop.get_bool(
            b"CausesValidation".into(),
            file_list_box_prop.causes_validation,
        );
        file_list_box_prop.drag_mode =
            prop.get_property(b"DragMode".into(), file_list_box_prop.drag_mode);
        file_list_box_prop.enabled = prop.get_bool(b"Enabled".into(), file_list_box_prop.enabled);
        file_list_box_prop.fore_color =
            prop.get_color(b"ForeColor".into(), file_list_box_prop.fore_color);
        file_list_box_prop.height = prop.get_i32(b"Height".into(), file_list_box_prop.height);
        file_list_box_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), file_list_box_prop.help_context_id);
        file_list_box_prop.hidden = prop.get_bool(b"Hidden".into(), file_list_box_prop.hidden);
        file_list_box_prop.left = prop.get_i32(b"Left".into(), file_list_box_prop.left);
        file_list_box_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), file_list_box_prop.mouse_pointer);
        file_list_box_prop.multi_select =
            prop.get_property(b"MultiSelect".into(), file_list_box_prop.multi_select);
        file_list_box_prop.normal = prop.get_bool(b"Normal".into(), file_list_box_prop.normal);
        file_list_box_prop.ole_drag_mode =
            prop.get_property(b"OLEDragMode".into(), file_list_box_prop.ole_drag_mode);
        file_list_box_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), file_list_box_prop.ole_drop_mode);
        file_list_box_prop.pattern = match prop.get(b"Pattern".into()) {
            Some(pattern) => pattern.into(),
            None => file_list_box_prop.pattern,
        };
        file_list_box_prop.read_only =
            prop.get_bool(b"ReadOnly".into(), file_list_box_prop.read_only);
        file_list_box_prop.system = prop.get_bool(b"System".into(), file_list_box_prop.system);
        file_list_box_prop.tab_index =
            prop.get_i32(b"TabIndex".into(), file_list_box_prop.tab_index);
        file_list_box_prop.tab_stop = prop.get_bool(b"TabStop".into(), file_list_box_prop.tab_stop);
        file_list_box_prop.tool_tip_text = match prop.get("ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => file_list_box_prop.tool_tip_text,
        };
        file_list_box_prop.top = prop.get_i32(b"Top".into(), file_list_box_prop.top);
        file_list_box_prop.visible =
            prop.get_property(b"Visible".into(), file_list_box_prop.visible);
        file_list_box_prop.whats_this_help_id = prop.get_i32(
            b"WhatsThisHelpID".into(),
            file_list_box_prop.whats_this_help_id,
        );
        file_list_box_prop.width = prop.get_i32(b"Width".into(), file_list_box_prop.width);

        file_list_box_prop
    }
}
