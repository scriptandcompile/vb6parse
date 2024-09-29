use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::{
    controls::{Appearance, FormLinkMode, MousePointer, OLEDropMode, StartUpPosition, WindowState},
    VB6Color,
};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property, VB6PropertyGroup,
};

use bstr::BStr;
use image::DynamicImage;
use serde::Serialize;

#[derive(Debug, PartialEq, Clone)]
pub struct MDIFormProperties<'a> {
    pub appearance: Appearance,
    pub auto_show_children: bool,
    pub back_color: VB6Color,
    pub caption: &'a BStr,
    pub client_height: i32,
    pub client_left: i32,
    pub client_top: i32,
    pub client_width: i32,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub icon: Option<DynamicImage>,
    pub left: i32,
    pub link_mode: FormLinkMode,
    pub link_topic: &'a BStr,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub moveable: bool,
    pub negotiate_toolbars: bool,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub right_to_left: bool,
    pub scroll_bars: bool,
    pub start_up_position: StartUpPosition,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help: bool,
    pub width: i32,
    pub window_state: WindowState,
}

impl Serialize for MDIFormProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::ser::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("MDIFormProperties", 29)?;
        state.serialize_field("appearance", &self.appearance)?;
        state.serialize_field("auto_show_children", &self.auto_show_children)?;
        state.serialize_field("back_color", &self.back_color)?;
        state.serialize_field("caption", &self.caption)?;
        state.serialize_field("client_height", &self.client_height)?;
        state.serialize_field("client_left", &self.client_left)?;
        state.serialize_field("client_top", &self.client_top)?;
        state.serialize_field("client_width", &self.client_width)?;
        state.serialize_field("enabled", &self.enabled)?;
        state.serialize_field("height", &self.height)?;
        state.serialize_field("help_context_id", &self.help_context_id)?;

        let option_text = self.icon.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("icon", &option_text)?;
        state.serialize_field("left", &self.left)?;
        state.serialize_field("link_mode", &self.link_mode)?;
        state.serialize_field("link_topic", &self.link_topic)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("mouse_icon", &option_text)?;
        state.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        state.serialize_field("moveable", &self.moveable)?;
        state.serialize_field("negotiate_toolbars", &self.negotiate_toolbars)?;
        state.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;

        let option_text = self.picture.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("picture", &option_text)?;
        state.serialize_field("right_to_left", &self.right_to_left)?;
        state.serialize_field("scroll_bars", &self.scroll_bars)?;
        state.serialize_field("start_up_position", &self.start_up_position)?;
        state.serialize_field("top", &self.top)?;
        state.serialize_field("visible", &self.visible)?;
        state.serialize_field("whats_this_help", &self.whats_this_help)?;
        state.serialize_field("width", &self.width)?;
        state.serialize_field("window_state", &self.window_state)?;

        state.end()
    }
}

impl Default for MDIFormProperties<'_> {
    fn default() -> Self {
        MDIFormProperties {
            appearance: Appearance::ThreeD,
            auto_show_children: true,
            back_color: VB6Color::from_hex("&H8000000C&").unwrap(),
            caption: BStr::new("MDIForm1"),
            client_height: 0,
            client_left: 0,
            client_top: 0,
            client_width: 0,
            enabled: true,
            //font
            height: 3600,
            help_context_id: 0,
            icon: None,
            left: 0,
            link_mode: FormLinkMode::None,
            link_topic: BStr::new("MDIForm1"),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            moveable: true,
            negotiate_toolbars: true,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            right_to_left: false,
            scroll_bars: true,
            start_up_position: StartUpPosition::WindowsDefault,
            top: 0,
            visible: true,
            whats_this_help: false,
            width: 4800,
            window_state: WindowState::Normal,
        }
    }
}

impl<'a> MDIFormProperties<'a> {
    pub fn construct_control(
        properties: HashMap<&'a BStr, &'a BStr>,
        _property_groups: Vec<VB6PropertyGroup<'a>>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut mdi_form_properties = MDIFormProperties::default();

        let appearance_key = BStr::new("Appearance");
        mdi_form_properties.appearance = build_property(&properties, appearance_key);

        let auto_show_children_key = BStr::new("AutoShowChildren");
        mdi_form_properties.auto_show_children = build_bool_property(
            &properties,
            auto_show_children_key,
            mdi_form_properties.auto_show_children,
        );

        let backcolor_key = BStr::new("BackColor");
        mdi_form_properties.back_color =
            build_color_property(&properties, backcolor_key, mdi_form_properties.back_color);

        let caption_key = BStr::new("Caption");
        mdi_form_properties.caption = properties
            .get(caption_key)
            .unwrap_or(&mdi_form_properties.caption);

        let client_height_key = BStr::new("ClientHeight");
        mdi_form_properties.client_height = build_i32_property(
            &properties,
            client_height_key,
            mdi_form_properties.client_height,
        );

        let client_left_key = BStr::new("ClientLeft");
        mdi_form_properties.client_left = build_i32_property(
            &properties,
            client_left_key,
            mdi_form_properties.client_left,
        );

        let client_top_key = BStr::new("ClientTop");
        mdi_form_properties.client_top =
            build_i32_property(&properties, client_top_key, mdi_form_properties.client_top);

        let client_width_key = BStr::new("ClientWidth");
        mdi_form_properties.client_width = build_i32_property(
            &properties,
            client_width_key,
            mdi_form_properties.client_width,
        );

        let enabled_key = BStr::new("Enabled");
        mdi_form_properties.enabled =
            build_bool_property(&properties, enabled_key, mdi_form_properties.enabled);

        // Font - group

        let height_key = BStr::new("Height");
        mdi_form_properties.height =
            build_i32_property(&properties, height_key, mdi_form_properties.height);

        let help_context_id_key = BStr::new("HelpContextID");
        mdi_form_properties.help_context_id = build_i32_property(
            &properties,
            help_context_id_key,
            mdi_form_properties.help_context_id,
        );

        // Icon

        let left_key = BStr::new("Left");
        mdi_form_properties.left =
            build_i32_property(&properties, left_key, mdi_form_properties.left);

        let link_mode_key = BStr::new("LinkMode");
        mdi_form_properties.link_mode = build_property(&properties, link_mode_key);

        let link_topic_key = BStr::new("LinkTopic");
        mdi_form_properties.link_topic = properties
            .get(link_topic_key)
            .unwrap_or(&mdi_form_properties.link_topic);

        // MouseIcon

        let mouse_pointer_key = BStr::new("MousePointer");
        mdi_form_properties.mouse_pointer = build_property(&properties, mouse_pointer_key);

        let moveable_key = BStr::new("Moveable");
        mdi_form_properties.moveable =
            build_bool_property(&properties, moveable_key, mdi_form_properties.moveable);

        let negotiate_toolbars_key = BStr::new("NegotiateToolbars");
        mdi_form_properties.negotiate_toolbars = build_bool_property(
            &properties,
            negotiate_toolbars_key,
            mdi_form_properties.negotiate_toolbars,
        );

        let ole_drop_mode_key = BStr::new("OLEDropMode");
        mdi_form_properties.ole_drop_mode = build_property(&properties, ole_drop_mode_key);

        // Picture

        mdi_form_properties.right_to_left = build_bool_property(
            &properties,
            BStr::new("RightToLeft"),
            mdi_form_properties.right_to_left,
        );

        mdi_form_properties.scroll_bars = build_bool_property(
            &properties,
            BStr::new("Scrollbars"),
            mdi_form_properties.scroll_bars,
        );

        mdi_form_properties.start_up_position =
            build_property(&properties, BStr::new("StartUpPosition"));

        mdi_form_properties.top =
            build_i32_property(&properties, BStr::new("Top"), mdi_form_properties.top);

        mdi_form_properties.visible = build_bool_property(
            &properties,
            BStr::new("Visible"),
            mdi_form_properties.visible,
        );

        mdi_form_properties.width =
            build_i32_property(&properties, BStr::new("Width"), mdi_form_properties.width);

        mdi_form_properties.window_state = build_property(&properties, BStr::new("WindowState"));

        Ok(mdi_form_properties)
    }
}
