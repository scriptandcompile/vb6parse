use crate::language::controls::{
    Appearance, DrawMode, DrawStyle, FillStyle, MousePointer, OLEDropMode, ScaleMode,
};
use crate::VB6Color;

use image::DynamicImage;
use serde::Serialize;

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub enum FormLinkMode {
    None = 0,
    Source = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub enum PaletteMode {
    HalfTone = 0,
    UseZOrder = 1,
    Custom = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub enum StartUpPosition {
    Manual = 0,
    CenterOwner = 1,
    CenterScreen = 2,
    WindowsDefault = 3,
}

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub enum FormBorderStyle {
    None = 0,
    FixedSingle = 1,
    Sizable = 2,
    FixedDialog = 3,
    FixedToolWindow = 4,
    SizableToolWindow = 5,
}

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub enum WindowState {
    Normal = 0,
    Minimized = 1,
    Maximized = 2,
}

/// Properties for a Form control. This is used as an enum variant of
/// [VB6ControlKind::Form](crate::language::controls::VB6ControlKind::Form).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [VB6Control](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FormProperties<'a> {
    pub appearance: Appearance,
    /// Determines if the output from a graphics method is to a persistent bitmap
    /// which acts as a double buffer.
    pub auto_redraw: bool,
    pub back_color: VB6Color,
    pub border_style: FormBorderStyle,
    pub caption: &'a str,
    pub clip_controls: bool,
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
    pub link_topic: &'a str,
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
    pub visible: bool,
    pub whats_this_button: bool,
    pub whats_this_help: bool,
    pub width: i32,
    pub window_state: WindowState,
}

impl Serialize for FormProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::ser::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("FormProperties", 38)?;
        state.serialize_field("appearance", &self.appearance)?;
        state.serialize_field("auto_redraw", &self.auto_redraw)?;
        state.serialize_field("back_color", &self.back_color)?;
        state.serialize_field("border_style", &self.border_style)?;
        state.serialize_field("caption", &self.caption)?;
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

        let option_text = match self.icon {
            Some(_) => Some("Some(DynamicImage)"),
            None => None,
        };

        state.serialize_field("icon", &option_text)?;
        state.serialize_field("key_preview", &self.key_preview)?;
        state.serialize_field("left", &self.left)?;
        state.serialize_field("link_mode", &self.link_mode)?;
        state.serialize_field("link_topic", &self.link_topic)?;
        state.serialize_field("max_button", &self.max_button)?;
        state.serialize_field("mdi_child", &self.mdi_child)?;
        state.serialize_field("min_button", &self.min_button)?;

        let option_text = match self.mouse_icon {
            Some(_) => Some("Some(DynamicImage)"),
            None => None,
        };

        state.serialize_field("mouse_icon", &option_text)?;
        state.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        state.serialize_field("moveable", &self.moveable)?;
        state.serialize_field("negotiate_menus", &self.negotiate_menus)?;
        state.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;

        let option_text = match self.palette {
            Some(_) => Some("Some(DynamicImage)"),
            None => None,
        };

        state.serialize_field("palette", &option_text)?;
        state.serialize_field("palette_mode", &self.palette_mode)?;

        let option_text = match self.palette {
            Some(_) => Some("Some(DynamicImage)"),
            None => None,
        };

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

impl Default for FormProperties<'_> {
    fn default() -> Self {
        FormProperties {
            appearance: Appearance::ThreeD,
            auto_redraw: false,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: FormBorderStyle::Sizable,
            caption: "Form1",
            clip_controls: true,
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
            link_topic: "Form1",
            max_button: true,
            mdi_child: false,
            min_button: true,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            moveable: true,
            negotiate_menus: true,
            ole_drop_mode: OLEDropMode::None,
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
            visible: true,
            whats_this_button: false,
            whats_this_help: false,
            width: 240,
            window_state: WindowState::Normal,
        }
    }
}
