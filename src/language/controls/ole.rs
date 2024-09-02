use crate::language::controls::{
    Appearance, BackStyle, BorderStyle, DragMode, MousePointer, SizeMode,
};
use crate::language::VB6Color;

use image::DynamicImage;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum OLETypeAllowed {
    Link = 0,
    Embedded = 1,
    Either = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum UpdateOptions {
    Automatic = 0,
    Frozen = 1,
    Manual = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum AutoActivate {
    Manual = 0,
    GetFocus = 1,
    DoubleClick = 2,
    Automatic = 3,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum DisplayType {
    Content = 0,
    Icon = 1,
}

#[derive(Debug, PartialEq, Clone)]
pub struct OLEProperties<'a> {
    pub appearance: Appearance,
    pub auto_activate: AutoActivate,
    pub auto_verb_menu: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub class: Option<&'a str>,
    pub data_field: &'a str,
    pub data_source: &'a str,
    pub display_type: DisplayType,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub host_name: &'a str,
    pub left: i32,
    pub misc_flags: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_allowed: bool,
    pub ole_type_allowed: OLETypeAllowed,
    pub size_mode: SizeMode,
    //pub source_doc: &'a str,
    //pub source_item: &'a str,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub top: i32,
    pub update_options: UpdateOptions,
    pub verb: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Serialize for OLEProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("OLEProperties", 31)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("auto_activate", &self.auto_activate)?;
        s.serialize_field("auto_verb_menu", &self.auto_verb_menu)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("back_style", &self.back_style)?;
        s.serialize_field("border_style", &self.border_style)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;
        s.serialize_field("class", &self.class)?;
        s.serialize_field("data_field", &self.data_field)?;
        s.serialize_field("data_source", &self.data_source)?;
        s.serialize_field("display_type", &self.display_type)?;

        let option_text = match &self.drag_icon {
            Some(_) => "Some(DynamicImage)",
            None => "None",
        };

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("host_name", &self.host_name)?;
        s.serialize_field("left", &self.left)?;
        s.serialize_field("misc_flags", &self.misc_flags)?;

        let option_text = match &self.mouse_icon {
            Some(_) => "Some(DynamicImage)",
            None => "None",
        };

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drop_allowed", &self.ole_drop_allowed)?;
        s.serialize_field("ole_type_allowed", &self.ole_type_allowed)?;
        s.serialize_field("size_mode", &self.size_mode)?;
        //s.serialize_field("source_doc", &self.source_doc)?;
        //s.serialize_field("source_item", &self.source_item)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("update_options", &self.update_options)?;
        s.serialize_field("verb", &self.verb)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}
