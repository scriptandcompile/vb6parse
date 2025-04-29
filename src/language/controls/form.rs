use crate::{
    language::{
        controls::{
            Appearance, AutoRedraw, ClipControls, DrawMode, DrawStyle, FillStyle, MousePointer,
            OLEDropMode, ScaleMode, StartUpPosition, Visibility, WindowState,
        },
        FormLinkMode, VB6Color,
    },
    parsers::Properties,
};

use bstr::BString;
use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum PaletteMode {
    #[default]
    HalfTone = 0,
    UseZOrder = 1,
    Custom = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Default, TryFromPrimitive, serde::Serialize)]
#[repr(i32)]
pub enum FormBorderStyle {
    None = 0,
    FixedSingle = 1,
    #[default]
    Sizable = 2,
    FixedDialog = 3,
    FixedToolWindow = 4,
    SizableToolWindow = 5,
}

/// Properties for a `Form` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Form`](crate::language::controls::VB6ControlKind::Form).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FormProperties {
    pub appearance: Appearance,
    pub auto_redraw: AutoRedraw,
    pub back_color: VB6Color,
    pub border_style: FormBorderStyle,
    pub caption: BString,
    pub client_height: i32,
    pub client_left: i32,
    pub client_top: i32,
    pub client_width: i32,
    pub clip_controls: ClipControls,
    pub control_box: bool,
    pub draw_mode: DrawMode,
    pub draw_style: DrawStyle,
    pub draw_width: i32,
    pub enabled: bool,
    pub fill_color: VB6Color,
    pub fill_style: FillStyle,
    pub font_transparent: bool,
    pub fore_color: VB6Color,
    pub has_dc: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub icon: Option<DynamicImage>,
    pub key_preview: bool,
    pub left: i32,
    pub link_mode: FormLinkMode,
    pub link_topic: BString,
    pub max_button: bool,
    pub mdi_child: bool,
    pub min_button: bool,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub moveable: bool,
    pub negotiate_menus: bool,
    pub ole_drop_mode: OLEDropMode,
    pub palette: Option<DynamicImage>,
    pub palette_mode: PaletteMode,
    pub picture: Option<DynamicImage>,
    pub right_to_left: bool,
    pub scale_height: i32,
    pub scale_left: i32,
    pub scale_mode: ScaleMode,
    pub scale_top: i32,
    pub scale_width: i32,
    pub show_in_taskbar: bool,
    pub start_up_position: StartUpPosition,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_button: bool,
    pub whats_this_help: bool,
    pub width: i32,
    pub window_state: WindowState,
}

impl Serialize for FormProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::ser::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("FormProperties", 42)?;
        state.serialize_field("appearance", &self.appearance)?;
        state.serialize_field("auto_redraw", &self.auto_redraw)?;
        state.serialize_field("back_color", &self.back_color)?;
        state.serialize_field("border_style", &self.border_style)?;
        state.serialize_field("caption", &self.caption)?;
        state.serialize_field("client_height", &self.client_height)?;
        state.serialize_field("client_left", &self.client_left)?;
        state.serialize_field("client_top", &self.client_top)?;
        state.serialize_field("client_width", &self.client_width)?;

        state.serialize_field("clip_controls", &self.clip_controls)?;
        state.serialize_field("control_box", &self.control_box)?;
        state.serialize_field("draw_mode", &self.draw_mode)?;
        state.serialize_field("draw_style", &self.draw_style)?;
        state.serialize_field("draw_width", &self.draw_width)?;
        state.serialize_field("enabled", &self.enabled)?;
        state.serialize_field("fill_color", &self.fill_color)?;
        state.serialize_field("fill_style", &self.fill_style)?;
        state.serialize_field("font_transparent", &self.font_transparent)?;
        state.serialize_field("fore_color", &self.fore_color)?;
        state.serialize_field("has_dc", &self.has_dc)?;
        state.serialize_field("height", &self.height)?;
        state.serialize_field("help_context_id", &self.help_context_id)?;

        let option_text = self.icon.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("icon", &option_text)?;
        state.serialize_field("key_preview", &self.key_preview)?;
        state.serialize_field("left", &self.left)?;
        state.serialize_field("link_mode", &self.link_mode)?;
        state.serialize_field("link_topic", &self.link_topic)?;
        state.serialize_field("max_button", &self.max_button)?;
        state.serialize_field("mdi_child", &self.mdi_child)?;
        state.serialize_field("min_button", &self.min_button)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("mouse_icon", &option_text)?;
        state.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        state.serialize_field("moveable", &self.moveable)?;
        state.serialize_field("negotiate_menus", &self.negotiate_menus)?;
        state.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;

        let option_text = self.palette.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("palette", &option_text)?;
        state.serialize_field("palette_mode", &self.palette_mode)?;

        let option_text = self.picture.as_ref().map(|_| "Some(DynamicImage)");

        state.serialize_field("picture", &option_text)?;
        state.serialize_field("right_to_left", &self.right_to_left)?;
        state.serialize_field("scale_height", &self.scale_height)?;
        state.serialize_field("scale_left", &self.scale_left)?;
        state.serialize_field("scale_mode", &self.scale_mode)?;
        state.serialize_field("scale_top", &self.scale_top)?;
        state.serialize_field("scale_width", &self.scale_width)?;
        state.serialize_field("show_in_taskbar", &self.show_in_taskbar)?;
        state.serialize_field("start_up_position", &self.start_up_position)?;
        state.serialize_field("top", &self.top)?;
        state.serialize_field("visible", &self.visible)?;
        state.serialize_field("whats_this_button", &self.whats_this_button)?;
        state.serialize_field("whats_this_help", &self.whats_this_help)?;
        state.serialize_field("width", &self.width)?;
        state.serialize_field("window_state", &self.window_state)?;

        state.end()
    }
}

impl Default for FormProperties {
    fn default() -> Self {
        FormProperties {
            appearance: Appearance::ThreeD,
            auto_redraw: AutoRedraw::Manual,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: FormBorderStyle::Sizable,
            caption: "Form1".into(),
            client_height: 200,
            client_left: 0,
            client_top: 0,
            client_width: 300,
            clip_controls: ClipControls::default(),
            control_box: true,
            draw_mode: DrawMode::CopyPen,
            draw_style: DrawStyle::Solid,
            draw_width: 1,
            enabled: true,
            fill_color: VB6Color::from_hex("&H00000000&").unwrap(),
            fill_style: FillStyle::Transparent,
            font_transparent: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            has_dc: true,
            height: 240,
            help_context_id: 0,
            icon: None,
            key_preview: false,
            left: 0,
            link_mode: FormLinkMode::None,
            link_topic: "".into(),
            max_button: true,
            mdi_child: false,
            min_button: true,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            moveable: true,
            negotiate_menus: true,
            ole_drop_mode: OLEDropMode::default(),
            palette: None,
            palette_mode: PaletteMode::HalfTone,
            picture: None,
            right_to_left: false,
            scale_height: 240,
            scale_left: 0,
            scale_mode: ScaleMode::Twip,
            scale_top: 0,
            scale_width: 240,
            show_in_taskbar: true,
            start_up_position: StartUpPosition::WindowsDefault,
            top: 0,
            visible: Visibility::Visible,
            whats_this_button: false,
            whats_this_help: false,
            width: 240,
            window_state: WindowState::Normal,
        }
    }
}

impl From<Properties<'_>> for FormProperties {
    fn from(prop: Properties) -> Self {
        let mut form_prop = FormProperties::default();

        form_prop.appearance = prop.get_property(b"Appearance".into(), form_prop.appearance);
        form_prop.auto_redraw = prop.get_property(b"AutoRedraw".into(), form_prop.auto_redraw);
        form_prop.back_color = prop.get_color(b"BackColor".into(), form_prop.back_color);
        form_prop.border_style = prop.get_property(b"BorderStyle".into(), form_prop.border_style);
        form_prop.caption = match prop.get(b"Caption".into()) {
            Some(caption) => caption.into(),
            None => form_prop.caption,
        };

        form_prop.client_height = prop.get_i32(b"ClientHeight".into(), form_prop.client_height);
        form_prop.client_left = prop.get_i32(b"ClientLeft".into(), form_prop.client_left);
        form_prop.client_top = prop.get_i32(b"ClientTop".into(), form_prop.client_top);
        form_prop.client_width = prop.get_i32(b"ClientWidth".into(), form_prop.client_width);

        form_prop.clip_controls =
            prop.get_property(b"ClipControls".into(), form_prop.clip_controls);
        form_prop.control_box = prop.get_bool(b"ControlBox".into(), form_prop.control_box);

        form_prop.draw_mode = prop.get_property(b"DrawMode".into(), form_prop.draw_mode);
        form_prop.draw_style = prop.get_property(b"DrawStyle".into(), form_prop.draw_style);
        form_prop.draw_width = prop.get_i32(b"DrawWidth".into(), form_prop.draw_width);

        form_prop.enabled = prop.get_bool(b"Enabled".into(), form_prop.enabled);

        form_prop.fill_color = prop.get_color(b"FillColor".into(), form_prop.fill_color);
        form_prop.fill_style = prop.get_property(b"FillStyle".into(), form_prop.fill_style);

        // Font - group

        form_prop.font_transparent =
            prop.get_bool(b"FontTransparent".into(), form_prop.font_transparent);
        form_prop.fore_color = prop.get_color(b"ForeColor".into(), form_prop.fore_color);
        form_prop.has_dc = prop.get_bool(b"HasDC".into(), form_prop.has_dc);
        form_prop.height = prop.get_i32(b"Height".into(), form_prop.height);
        form_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), form_prop.help_context_id);

        // Icon

        form_prop.key_preview = prop.get_bool(b"KeyPreview".into(), form_prop.key_preview);
        form_prop.left = prop.get_i32(b"Left".into(), form_prop.left);
        form_prop.link_mode = prop.get_property(b"LinkMode".into(), form_prop.link_mode);
        form_prop.link_topic = match prop.get(b"LinkTopic".into()) {
            Some(link_topic) => link_topic.into(),
            None => form_prop.link_topic,
        };
        form_prop.max_button = prop.get_bool(b"MaxButton".into(), form_prop.max_button);
        form_prop.mdi_child = prop.get_bool(b"MDIChild".into(), form_prop.mdi_child);
        form_prop.min_button = prop.get_bool(b"MinButton".into(), form_prop.min_button);

        // MouseIcon

        form_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), form_prop.mouse_pointer);
        form_prop.moveable = prop.get_bool(b"Moveable".into(), form_prop.moveable);
        form_prop.negotiate_menus =
            prop.get_bool(b"NegotiateMenus".into(), form_prop.negotiate_menus);
        form_prop.ole_drop_mode = prop.get_property(b"OLEDropMode".into(), form_prop.ole_drop_mode);

        // Palette

        form_prop.palette_mode = prop.get_property(b"PaletteMode".into(), form_prop.palette_mode);

        // Picture

        form_prop.right_to_left = prop.get_bool(b"RightToLeft".into(), form_prop.right_to_left);
        form_prop.scale_height = prop.get_i32(b"ScaleHeight".into(), form_prop.scale_height);
        form_prop.scale_left = prop.get_i32(b"ScaleLeft".into(), form_prop.scale_left);
        form_prop.scale_mode = prop.get_property(b"ScaleMode".into(), form_prop.scale_mode);
        form_prop.scale_top = prop.get_i32(b"ScaleTop".into(), form_prop.scale_top);
        form_prop.scale_width = prop.get_i32(b"ScaleWidth".into(), form_prop.scale_width);
        form_prop.show_in_taskbar =
            prop.get_bool(b"ShowInTaskbar".into(), form_prop.show_in_taskbar);
        form_prop.start_up_position =
            prop.get_startup_position(b"StartUpPosition".into(), form_prop.start_up_position);
        form_prop.top = prop.get_i32(b"Top".into(), form_prop.top);
        form_prop.visible = prop.get_property(b"Visible".into(), form_prop.visible);
        form_prop.whats_this_button =
            prop.get_bool(b"WhatsThisButton".into(), form_prop.whats_this_button);
        form_prop.width = prop.get_i32(b"Width".into(), form_prop.width);
        form_prop.window_state = prop.get_property(b"WindowState".into(), form_prop.window_state);

        form_prop
    }
}
