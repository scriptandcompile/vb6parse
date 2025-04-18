use crate::language::{
    controls::{Appearance, FormLinkMode, MousePointer, OLEDropMode, StartUpPosition, WindowState},
    VB6Color,
};
use crate::parsers::Properties;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

#[derive(Debug, PartialEq, Clone)]
pub struct MDIFormProperties {
    pub appearance: Appearance,
    pub auto_show_children: bool,
    pub back_color: VB6Color,
    pub caption: BString,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub icon: Option<DynamicImage>,
    pub left: i32,
    pub link_mode: FormLinkMode,
    pub link_topic: BString,
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

impl Serialize for MDIFormProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::ser::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("MDIFormProperties", 25)?;
        state.serialize_field("appearance", &self.appearance)?;
        state.serialize_field("auto_show_children", &self.auto_show_children)?;
        state.serialize_field("back_color", &self.back_color)?;
        state.serialize_field("caption", &self.caption)?;
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

impl Default for MDIFormProperties {
    fn default() -> Self {
        MDIFormProperties {
            appearance: Appearance::ThreeD,
            auto_show_children: true,
            back_color: VB6Color::from_hex("&H8000000C&").unwrap(),
            caption: "".into(),
            enabled: true,
            //font
            height: 3600,
            help_context_id: 0,
            icon: None,
            left: 0,
            link_mode: FormLinkMode::None,
            link_topic: "".into(),
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

impl<'a> From<Properties<'a>> for MDIFormProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut mdi_form_prop = MDIFormProperties::default();

        mdi_form_prop.appearance =
            prop.get_property(b"Appearance".into(), mdi_form_prop.appearance);

        mdi_form_prop.auto_show_children =
            prop.get_bool(b"AutoShowChildren".into(), mdi_form_prop.auto_show_children);

        mdi_form_prop.back_color = prop.get_color(b"BackColor".into(), mdi_form_prop.back_color);

        mdi_form_prop.caption = match prop.get(b"Caption".into()) {
            Some(caption) => caption.into(),
            None => mdi_form_prop.caption,
        };
        mdi_form_prop.enabled = prop.get_bool(b"Enabled".into(), mdi_form_prop.enabled);

        // Font - group

        mdi_form_prop.height = prop.get_i32(b"Height".into(), mdi_form_prop.height);

        mdi_form_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), mdi_form_prop.help_context_id);

        // Icon

        mdi_form_prop.left = prop.get_i32(b"Left".into(), mdi_form_prop.left);

        mdi_form_prop.link_mode = prop.get_property(b"LinkMode".into(), mdi_form_prop.link_mode);

        mdi_form_prop.link_topic = match prop.get(b"LinkTopic".into()) {
            Some(link_topic) => link_topic.into(),
            None => mdi_form_prop.link_topic,
        };

        // MouseIcon

        mdi_form_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), mdi_form_prop.mouse_pointer);

        mdi_form_prop.moveable = prop.get_bool(b"Moveable".into(), mdi_form_prop.moveable);

        mdi_form_prop.negotiate_toolbars = prop.get_bool(
            b"NegotiateToolbars".into(),
            mdi_form_prop.negotiate_toolbars,
        );

        mdi_form_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), mdi_form_prop.ole_drop_mode);

        // Picture

        mdi_form_prop.right_to_left =
            prop.get_bool(b"RightToLeft".into(), mdi_form_prop.right_to_left);

        mdi_form_prop.scroll_bars = prop.get_bool(b"Scrollbars".into(), mdi_form_prop.scroll_bars);

        mdi_form_prop.start_up_position =
            prop.get_startup_position(b"StartUpPosition".into(), mdi_form_prop.start_up_position);

        mdi_form_prop.top = prop.get_i32(b"Top".into(), mdi_form_prop.top);

        mdi_form_prop.visible = prop.get_bool(b"Visible".into(), mdi_form_prop.visible);

        mdi_form_prop.width = prop.get_i32(b"Width".into(), mdi_form_prop.width);

        mdi_form_prop.window_state =
            prop.get_property(b"WindowState".into(), mdi_form_prop.window_state);

        mdi_form_prop
    }
}
