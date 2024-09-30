use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::{
    controls::{
        Appearance, ClipControls, DrawMode, DrawStyle, FillStyle, MousePointer, OLEDropMode,
        ScaleMode, StartUpPosition, WindowState,
    },
    FormLinkMode, VB6Color,
};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
    build_startup_position_property, VB6PropertyGroup,
};

use bstr::BStr;
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
pub struct FormProperties<'a> {
    pub appearance: Appearance,
    /// Determines if the output from a graphics method is to a persistent bitmap
    /// which acts as a double buffer.
    pub auto_redraw: bool,
    pub back_color: VB6Color,
    pub border_style: FormBorderStyle,
    pub caption: &'a BStr,
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
    pub link_topic: &'a BStr,
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

impl Default for FormProperties<'_> {
    fn default() -> Self {
        FormProperties {
            appearance: Appearance::ThreeD,
            auto_redraw: false,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: FormBorderStyle::Sizable,
            caption: BStr::new("Form1"),
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
            link_topic: BStr::new("Form1"),
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
            visible: true,
            whats_this_button: false,
            whats_this_help: false,
            width: 240,
            window_state: WindowState::Normal,
        }
    }
}

impl<'a> FormProperties<'a> {
    pub fn construct_control(
        properties: HashMap<&'a BStr, &'a BStr>,
        _property_groups: Vec<VB6PropertyGroup<'a>>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut form_properties = FormProperties::default();

        form_properties.appearance = build_property(&properties, b"Appearance");
        form_properties.auto_redraw =
            build_bool_property(&properties, b"AutoRedraw", form_properties.auto_redraw);
        form_properties.back_color =
            build_color_property(&properties, b"BackColor", form_properties.back_color);
        form_properties.border_style = build_property(&properties, b"BorderStyle");
        form_properties.caption = properties
            .get(BStr::new("Caption"))
            .unwrap_or(&form_properties.caption);
        form_properties.clip_controls = build_property(&properties, b"ClipControls");
        form_properties.control_box =
            build_bool_property(&properties, b"ControlBox", form_properties.control_box);
        form_properties.draw_mode = build_property(&properties, b"DrawMode");
        form_properties.draw_style = build_property(&properties, b"DrawStyle");
        form_properties.draw_width =
            build_i32_property(&properties, b"DrawWidth", form_properties.draw_width);
        form_properties.enabled =
            build_bool_property(&properties, b"Enabled", form_properties.enabled);
        form_properties.fill_color =
            build_color_property(&properties, b"FillColor", form_properties.fill_color);
        form_properties.fill_style = build_property(&properties, b"FillStyle");

        // Font - group

        form_properties.font_transparent = build_bool_property(
            &properties,
            b"FontTransparent",
            form_properties.font_transparent,
        );
        form_properties.fore_color =
            build_color_property(&properties, b"ForeColor", form_properties.fore_color);
        form_properties.has_dc = build_bool_property(&properties, b"HasDC", form_properties.has_dc);
        form_properties.height = build_i32_property(&properties, b"Height", form_properties.height);
        form_properties.help_context_id = build_i32_property(
            &properties,
            b"HelpContextID",
            form_properties.help_context_id,
        );

        // Icon

        form_properties.key_preview =
            build_bool_property(&properties, b"KeyPreview", form_properties.key_preview);
        form_properties.left = build_i32_property(&properties, b"Left", form_properties.left);
        form_properties.link_mode = build_property(&properties, b"LinkMode");
        form_properties.link_topic = properties
            .get(BStr::new("LinkTopic"))
            .unwrap_or(&form_properties.link_topic);
        form_properties.max_button =
            build_bool_property(&properties, b"MaxButton", form_properties.max_button);
        form_properties.mdi_child =
            build_bool_property(&properties, b"MDIChild", form_properties.mdi_child);
        form_properties.min_button =
            build_bool_property(&properties, b"MinButton", form_properties.min_button);

        // MouseIcon

        form_properties.mouse_pointer = build_property(&properties, b"MousePointer");
        form_properties.moveable =
            build_bool_property(&properties, b"Moveable", form_properties.moveable);
        form_properties.negotiate_menus = build_bool_property(
            &properties,
            b"NegotiateMenus",
            form_properties.negotiate_menus,
        );
        form_properties.ole_drop_mode = build_property(&properties, b"OLEDropMode");

        // Palette

        form_properties.palette_mode = build_property(&properties, b"PaletteMode");

        // Picture

        form_properties.right_to_left =
            build_bool_property(&properties, b"RightToLeft", form_properties.right_to_left);
        form_properties.scale_height =
            build_i32_property(&properties, b"ScaleHeight", form_properties.scale_height);
        form_properties.scale_left =
            build_i32_property(&properties, b"ScaleLeft", form_properties.scale_left);
        form_properties.scale_mode = build_property(&properties, b"ScaleMode");
        form_properties.scale_top =
            build_i32_property(&properties, b"ScaleTop", form_properties.scale_top);
        form_properties.scale_width =
            build_i32_property(&properties, b"ScaleWidth", form_properties.scale_width);
        form_properties.show_in_taskbar = build_bool_property(
            &properties,
            b"ShowInTaskbar",
            form_properties.show_in_taskbar,
        );
        form_properties.start_up_position =
            build_startup_position_property(&properties, b"StartUpPosition");
        form_properties.top = build_i32_property(&properties, b"Top", form_properties.top);
        form_properties.visible =
            build_bool_property(&properties, b"Visible", form_properties.visible);
        form_properties.whats_this_button = build_bool_property(
            &properties,
            b"WhatsThisButton",
            form_properties.whats_this_button,
        );
        form_properties.width = build_i32_property(&properties, b"Width", form_properties.width);
        form_properties.window_state = build_property(&properties, b"WindowState");

        Ok(form_properties)
    }
}
