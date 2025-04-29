use crate::language::controls::{
    Appearance, BackStyle, BorderStyle, DragMode, MousePointer, SizeMode, Visibility,
};
use crate::language::VB6Color;
use crate::parsers::Properties;

use bstr::BString;
use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum OLETypeAllowed {
    Link = 0,
    Embedded = 1,
    #[default]
    Either = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum UpdateOptions {
    #[default]
    Automatic = 0,
    Frozen = 1,
    Manual = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum AutoActivate {
    Manual = 0,
    GetFocus = 1,
    #[default]
    DoubleClick = 2,
    Automatic = 3,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum DisplayType {
    #[default]
    Content = 0,
    Icon = 1,
}

/// Properties for a `OLE` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Ole`](crate::language::controls::VB6ControlKind::Ole).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct OLEProperties {
    pub appearance: Appearance,
    pub auto_activate: AutoActivate,
    pub auto_verb_menu: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub class: Option<BString>,
    pub data_field: BString,
    pub data_source: BString,
    pub display_type: DisplayType,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub host_name: BString,
    pub left: i32,
    pub misc_flags: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_allowed: bool,
    pub ole_type_allowed: OLETypeAllowed,
    pub size_mode: SizeMode,
    //pub source_doc: BString,
    //pub source_item: BString,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub top: i32,
    pub update_options: UpdateOptions,
    pub verb: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for OLEProperties {
    fn default() -> Self {
        OLEProperties {
            appearance: Appearance::ThreeD,
            auto_activate: AutoActivate::DoubleClick,
            auto_verb_menu: true,
            back_color: VB6Color::System { index: 5 },
            back_style: BackStyle::Opaque,
            border_style: BorderStyle::FixedSingle,
            causes_validation: true,
            class: None,
            data_field: "".into(),
            data_source: "".into(),
            display_type: DisplayType::Content,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            height: 375,
            help_context_id: 0,
            host_name: "".into(),
            left: 600,
            misc_flags: 0,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_allowed: false,
            ole_type_allowed: OLETypeAllowed::Either,
            size_mode: SizeMode::Clip,
            //source_doc: "".into(),
            //source_item: "".into(),
            tab_index: 0,
            tab_stop: true,
            top: 1200,
            update_options: UpdateOptions::Automatic,
            verb: 0,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 1335,
        }
    }
}

impl Serialize for OLEProperties {
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

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("host_name", &self.host_name)?;
        s.serialize_field("left", &self.left)?;
        s.serialize_field("misc_flags", &self.misc_flags)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

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

impl<'a> From<Properties<'a>> for OLEProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut ole_prop = OLEProperties::default();

        ole_prop.appearance = prop.get_property(b"Appearance".into(), ole_prop.appearance);
        ole_prop.auto_activate = prop.get_property(b"AutoActivate".into(), ole_prop.auto_activate);
        ole_prop.auto_verb_menu = prop.get_bool(b"AutoVerbMenu".into(), ole_prop.auto_verb_menu);
        ole_prop.back_color = prop.get_color(b"BackColor".into(), ole_prop.back_color);
        ole_prop.back_style = prop.get_property(b"BackStyle".into(), ole_prop.back_style);
        ole_prop.border_style = prop.get_property(b"BorderStyle".into(), ole_prop.border_style);
        ole_prop.causes_validation =
            prop.get_bool(b"CausesValidation".into(), ole_prop.causes_validation);

        // Class

        ole_prop.data_field = match prop.get(b"DataField".into()) {
            Some(data_field) => data_field.into(),
            None => ole_prop.data_field,
        };
        ole_prop.data_source = match prop.get(b"DataSource".into()) {
            Some(data_source) => data_source.into(),
            None => ole_prop.data_source,
        };
        ole_prop.display_type = prop.get_property(b"DisplayType".into(), ole_prop.display_type);

        // DragIcon

        ole_prop.drag_mode = prop.get_property(b"DragMode".into(), ole_prop.drag_mode);
        ole_prop.enabled = prop.get_bool(b"Enabled".into(), ole_prop.enabled);
        ole_prop.height = prop.get_i32(b"Height".into(), ole_prop.height);
        ole_prop.help_context_id = prop.get_i32(b"HelpContextID".into(), ole_prop.help_context_id);
        ole_prop.host_name = match prop.get(b"HostName".into()) {
            Some(host_name) => host_name.into(),
            None => ole_prop.host_name,
        };
        ole_prop.left = prop.get_i32(b"Left".into(), ole_prop.left);
        ole_prop.misc_flags = prop.get_i32(b"MiscFlags".into(), ole_prop.misc_flags);
        ole_prop.mouse_pointer = prop.get_property(b"MousePointer".into(), ole_prop.mouse_pointer);
        ole_prop.ole_drop_allowed =
            prop.get_bool(b"OLEDropAllowed".into(), ole_prop.ole_drop_allowed);
        ole_prop.ole_type_allowed =
            prop.get_property(b"OLETypeAllowed".into(), ole_prop.ole_type_allowed);
        ole_prop.size_mode = prop.get_property(b"SizeMode".into(), ole_prop.size_mode);
        ole_prop.tab_index = prop.get_i32(b"TabIndex".into(), ole_prop.tab_index);
        ole_prop.tab_stop = prop.get_bool(b"TabStop".into(), ole_prop.tab_stop);
        ole_prop.top = prop.get_i32(b"Top".into(), ole_prop.top);
        ole_prop.update_options =
            prop.get_property(b"UpdateOptions".into(), ole_prop.update_options);
        ole_prop.verb = prop.get_i32(b"Verb".into(), ole_prop.verb);
        ole_prop.visible = prop.get_property(b"Visible".into(), ole_prop.visible);
        ole_prop.whats_this_help_id =
            prop.get_i32(b"WhatsThisHelpID".into(), ole_prop.whats_this_help_id);
        ole_prop.width = prop.get_i32(b"Width".into(), ole_prop.width);

        ole_prop
    }
}
