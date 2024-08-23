use crate::language::controls::{
    Appearance, DrawMode, DrawStyle, FillStyle, MousePointer, OLEDropMode, ScaleMode,
};
use crate::VB6Color;

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum FormLinkMode {
    None = 0,
    Source = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum PaletteMode {
    HalfTone = 0,
    UseZOrder = 1,
    Custom = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum StartUpPosition {
    Manual = 0,
    CenterOwner = 1,
    CenterScreen = 2,
    WindowsDefault = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum FormBorderStyle {
    None = 0,
    FixedSingle = 1,
    Sizable = 2,
    FixedDialog = 3,
    FixedToolWindow = 4,
    SizableToolWindow = 5,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum WindowState {
    Normal = 0,
    Minimized = 1,
    Maximized = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
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
    // pub icon: Option<ImageBuffer>,
    pub key_preview: bool,
    pub left: i32,
    pub link_mode: FormLinkMode,
    pub link_topic: &'a str,
    pub max_button: bool,
    pub mdi_child: bool,
    pub min_button: bool,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub moveable: bool,
    pub negotiate_menus: bool,
    pub ole_drop_mode: OLEDropMode,
    // pub palette: Option<ImageBuffer>,
    pub pallette_mode: PaletteMode,
    // pub picture: Option<ImageBuffer>,
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
            key_preview: false,
            left: 0,
            link_mode: FormLinkMode::None,
            link_topic: "Form1",
            max_button: true,
            mdi_child: false,
            min_button: true,
            mouse_pointer: MousePointer::Default,
            moveable: true,
            negotiate_menus: true,
            ole_drop_mode: OLEDropMode::None,
            pallette_mode: PaletteMode::HalfTone,
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
