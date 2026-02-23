//! Properties for a `MDIForm` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::MDIForm`](crate::language::controls::ControlKind::MDIForm).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use crate::files::common::Properties;
use crate::language::{
    controls::{
        Activation, Appearance, Font, FormLinkMode, MousePointer, Movability, OLEDropMode,
        ReferenceOrValue, StartUpPosition, TextDirection, Visibility, WhatsThisHelp, WindowState,
    },
    Color, VB_APPLICATION_WORKSPACE,
};

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `MDIForm`.
///
/// This is used as an enum variant of `FormRoot`
/// [`FormRoot::MDIForm`](crate::language::controls::FormRoot).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`MDIForm`](crate::language::controls::MDIForm) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct MDIFormProperties {
    /// Appearance of the MDI form.
    pub appearance: Appearance,
    /// Auto show children setting of the MDI form.
    pub auto_show_children: bool,
    /// Background color of the MDI form.
    pub back_color: Color,
    /// Caption of the MDI form.
    pub caption: String,
    /// Enabled state of the MDI form.
    pub enabled: Activation,
    /// The font style for the form.
    pub font: Option<Font>,
    /// Height of the MDI form.
    pub height: i32,
    /// Help context ID of the MDI form.
    pub help_context_id: i32,
    /// Icon of the MDI form.
    pub icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Left position of the MDI form.
    pub left: i32,
    /// Link mode of the MDI form.
    pub link_mode: FormLinkMode,
    /// Link topic of the MDI form.
    pub link_topic: String,
    /// Mouse icon of the MDI form.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Mouse pointer of the MDI form.
    pub mouse_pointer: MousePointer,
    /// Movability of the MDI form.
    pub moveable: Movability,
    /// Negotiate toolbars setting of the MDI form.
    pub negotiate_toolbars: bool,
    /// OLE drop mode of the MDI form.
    pub ole_drop_mode: OLEDropMode,
    /// Picture of the MDI form.
    pub picture: Option<ReferenceOrValue<DynamicImage>>,
    /// Right to left setting of the MDI form.
    pub right_to_left: TextDirection,
    /// Scroll bars setting of the MDI form.
    pub scroll_bars: bool,
    /// Start up position of the MDI form.
    pub start_up_position: StartUpPosition,
    /// Top position of the MDI form.
    pub top: i32,
    /// Visibility of the MDI form.
    pub visible: Visibility,
    /// What's This help setting of the MDI form.
    pub whats_this_help: WhatsThisHelp,
    /// Width of the MDI form.
    pub width: i32,
    /// Window state of the MDI form.
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
            back_color: VB_APPLICATION_WORKSPACE,
            caption: String::new(),
            enabled: Activation::Enabled,
            font: Some(Font::default()),
            height: 3600,
            help_context_id: 0,
            icon: None,
            left: 0,
            link_mode: FormLinkMode::None,
            link_topic: String::new(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            moveable: Movability::Moveable,
            negotiate_toolbars: true,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            right_to_left: TextDirection::LeftToRight,
            scroll_bars: true,
            start_up_position: StartUpPosition::WindowsDefault,
            top: 0,
            visible: Visibility::Visible,
            whats_this_help: WhatsThisHelp::F1Help,
            width: 4800,
            window_state: WindowState::Normal,
        }
    }
}

impl From<Properties> for MDIFormProperties {
    fn from(prop: Properties) -> Self {
        let mut mdi_form_prop = MDIFormProperties::default();

        mdi_form_prop.appearance = prop.get_property("Appearance", mdi_form_prop.appearance);
        mdi_form_prop.auto_show_children =
            prop.get_bool("AutoShowChildren", mdi_form_prop.auto_show_children);
        mdi_form_prop.back_color = prop.get_color("BackColor", mdi_form_prop.back_color);
        mdi_form_prop.caption = match prop.get("Caption") {
            Some(caption) => caption.into(),
            None => mdi_form_prop.caption,
        };
        mdi_form_prop.enabled = prop.get_property("Enabled", mdi_form_prop.enabled);

        // Font - group

        mdi_form_prop.height = prop.get_i32("Height", mdi_form_prop.height);
        mdi_form_prop.help_context_id =
            prop.get_i32("HelpContextID", mdi_form_prop.help_context_id);

        // Icon

        mdi_form_prop.left = prop.get_i32("Left", mdi_form_prop.left);
        mdi_form_prop.link_mode = prop.get_property("LinkMode", mdi_form_prop.link_mode);
        mdi_form_prop.link_topic = match prop.get("LinkTopic") {
            Some(link_topic) => link_topic.into(),
            None => mdi_form_prop.link_topic,
        };

        // MouseIcon

        mdi_form_prop.mouse_pointer =
            prop.get_property("MousePointer", mdi_form_prop.mouse_pointer);
        mdi_form_prop.moveable = prop.get_property("Moveable", mdi_form_prop.moveable);
        mdi_form_prop.negotiate_toolbars =
            prop.get_bool("NegotiateToolbars", mdi_form_prop.negotiate_toolbars);
        mdi_form_prop.ole_drop_mode = prop.get_property("OLEDropMode", mdi_form_prop.ole_drop_mode);

        // Picture

        mdi_form_prop.right_to_left = prop.get_property("RightToLeft", mdi_form_prop.right_to_left);
        mdi_form_prop.scroll_bars = prop.get_bool("Scrollbars", mdi_form_prop.scroll_bars);
        mdi_form_prop.start_up_position =
            prop.get_startup_position("StartUpPosition", mdi_form_prop.start_up_position);
        mdi_form_prop.top = prop.get_i32("Top", mdi_form_prop.top);
        mdi_form_prop.visible = prop.get_property("Visible", mdi_form_prop.visible);
        mdi_form_prop.whats_this_help =
            prop.get_property("WhatsThisHelp", mdi_form_prop.whats_this_help);
        mdi_form_prop.width = prop.get_i32("Width", mdi_form_prop.width);
        mdi_form_prop.window_state = prop.get_property("WindowState", mdi_form_prop.window_state);

        mdi_form_prop
    }
}
