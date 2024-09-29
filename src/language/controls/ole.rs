use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{
    Appearance, BackStyle, BorderStyle, DragMode, MousePointer, SizeMode,
};
use crate::language::VB6Color;
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};

use bstr::BStr;
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
pub struct OLEProperties<'a> {
    pub appearance: Appearance,
    pub auto_activate: AutoActivate,
    pub auto_verb_menu: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub class: Option<&'a BStr>,
    pub data_field: &'a BStr,
    pub data_source: &'a BStr,
    pub display_type: DisplayType,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub host_name: &'a BStr,
    pub left: i32,
    pub misc_flags: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_allowed: bool,
    pub ole_type_allowed: OLETypeAllowed,
    pub size_mode: SizeMode,
    //pub source_doc: &'a BStr,
    //pub source_item: &'a BStr,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub top: i32,
    pub update_options: UpdateOptions,
    pub verb: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for OLEProperties<'_> {
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
            data_field: BStr::new(""),
            data_source: BStr::new(""),
            display_type: DisplayType::Content,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            height: 375,
            help_context_id: 0,
            host_name: BStr::new(""),
            left: 600,
            misc_flags: 0,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_allowed: false,
            ole_type_allowed: OLETypeAllowed::Either,
            size_mode: SizeMode::Clip,
            //source_doc: BStr::new(""),
            //source_item: BStr::new(""),
            tab_index: 0,
            tab_stop: true,
            top: 1200,
            update_options: UpdateOptions::Automatic,
            verb: 0,
            visible: true,
            whats_this_help_id: 0,
            width: 1335,
        }
    }
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

impl<'a> OLEProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut ole_properties = OLEProperties::default();

        ole_properties.appearance = build_property(properties, b"Appearance");
        ole_properties.auto_activate = build_property(properties, b"AutoActivate");
        ole_properties.auto_verb_menu =
            build_bool_property(properties, b"AutoVerbMenu", ole_properties.auto_verb_menu);
        ole_properties.back_color =
            build_color_property(properties, b"BackColor", ole_properties.back_color);
        ole_properties.back_style = build_property(properties, b"BackStyle");
        ole_properties.border_style = build_property(properties, b"BorderStyle");
        ole_properties.causes_validation = build_bool_property(
            properties,
            b"CausesValidation",
            ole_properties.causes_validation,
        );

        // Class

        ole_properties.data_field = properties
            .get(&BStr::new("DataField"))
            .unwrap_or(&ole_properties.data_field);
        ole_properties.data_source = properties
            .get(&BStr::new("DataSource"))
            .unwrap_or(&ole_properties.data_source);
        ole_properties.display_type = build_property(properties, b"DisplayType");

        // DragIcon

        ole_properties.drag_mode = build_property(properties, b"DragMode");
        ole_properties.enabled =
            build_bool_property(properties, b"Enabled", ole_properties.enabled);
        ole_properties.height = build_i32_property(properties, b"Height", ole_properties.height);
        ole_properties.help_context_id =
            build_i32_property(properties, b"HelpContextID", ole_properties.help_context_id);
        ole_properties.host_name = properties
            .get(&BStr::new("HostName"))
            .unwrap_or(&ole_properties.host_name);
        ole_properties.left = build_i32_property(properties, b"Left", ole_properties.left);
        ole_properties.misc_flags =
            build_i32_property(properties, b"MiscFlags", ole_properties.misc_flags);
        ole_properties.mouse_pointer = build_property(properties, b"MousePointer");
        ole_properties.ole_drop_allowed = build_bool_property(
            properties,
            b"OLEDropAllowed",
            ole_properties.ole_drop_allowed,
        );
        ole_properties.ole_type_allowed = build_property(properties, b"OLETypeAllowed");
        ole_properties.size_mode = build_property(properties, b"SizeMode");
        ole_properties.tab_index =
            build_i32_property(properties, b"TabIndex", ole_properties.tab_index);
        ole_properties.tab_stop =
            build_bool_property(properties, b"TabStop", ole_properties.tab_stop);
        ole_properties.top = build_i32_property(properties, b"Top", ole_properties.top);
        ole_properties.update_options = build_property(properties, b"UpdateOptions");
        ole_properties.verb = build_i32_property(properties, b"Verb", ole_properties.verb);
        ole_properties.visible =
            build_bool_property(properties, b"Visible", ole_properties.visible);
        ole_properties.whats_this_help_id = build_i32_property(
            properties,
            b"WhatsThisHelpID",
            ole_properties.whats_this_help_id,
        );
        ole_properties.width = build_i32_property(properties, b"Width", ole_properties.width);

        Ok(ole_properties)
    }
}
