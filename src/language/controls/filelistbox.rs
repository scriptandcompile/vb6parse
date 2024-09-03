use crate::language::color::VB6Color;
use crate::language::controls::{
    Appearance, DragMode, MousePointer, MultiSelect, OLEDragMode, OLEDropMode,
};

use image::DynamicImage;
use serde::Serialize;

/// Properties for a FileListBox control. This is used as an enum variant of
/// [VB6ControlKind::FileListBox](crate::language::controls::VB6ControlKind::FileListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [VB6Control](crate::language::controls::VB6Control) struct.
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
    pub pattern: &'a str,
    pub read_only: bool,
    pub system: bool,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
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
            ole_drop_mode: OLEDropMode::None,
            pattern: "*.*",
            read_only: true,
            system: false,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: "",
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

        let option_text = match &self.drag_icon {
            Some(_) => "Some(DynamicImage)",
            None => "None",
        };

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("hidden", &self.hidden)?;
        s.serialize_field("left", &self.left)?;

        let option_text = match &self.mouse_icon {
            Some(_) => "Some(DynamicImage)",
            None => "None",
        };

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
