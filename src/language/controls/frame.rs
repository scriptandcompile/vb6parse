use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::color::VB6Color;
use crate::language::controls::{
    Appearance, BorderStyle, ClipControls, DragMode, MousePointer, OLEDropMode,
};
use crate::parsers::form::VB6PropertyGroup;

use bstr::{BStr, ByteSlice};
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `Frame` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Frame`](crate::language::controls::VB6ControlKind::Frame).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FrameProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub caption: &'a BStr,
    pub clip_controls: ClipControls,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub tab_index: i32,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for FrameProperties<'_> {
    fn default() -> Self {
        FrameProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            caption: BStr::new("Frame1"),
            clip_controls: ClipControls::True,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            left: 30,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::default(),
            right_to_left: false,
            tab_index: 0,
            tool_tip_text: BStr::new(""),
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for FrameProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("FrameProperties", 20)?;

        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("border_style", &self.border_style)?;
        s.serialize_field("caption", &self.caption)?;
        s.serialize_field("clip_controls", &self.clip_controls)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("left", &self.left)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl<'a> FrameProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
        _property_groups: &[VB6PropertyGroup<'a>],
    ) -> Result<Self, VB6ErrorKind> {
        // TODO: We are not correctly handling property assignment for each control.

        let mut frame_properties = FrameProperties::default();

        frame_properties.appearance = appearance_property(&properties)?;

        frame_properties.back_color = back_color_properties(properties)?;

        frame_properties.border_style = border_style_property(&properties)?;

        let caption_key = BStr::new("Caption");
        if properties.contains_key(caption_key) {
            frame_properties.caption = properties[caption_key];
        }

        frame_properties.clip_controls = clip_controls_property(&properties)?;

        // TODO: Implement loading drag_icon picture loading.
        // drag_icon: None,

        frame_properties.drag_mode = drag_mode_property(&properties)?;

        frame_properties.enabled = enabled_property(&properties)?;

        frame_properties.fore_color = fore_color_properties(properties)?;

        frame_properties.height = height_property(&properties)?;

        frame_properties.help_context_id = help_context_id_property(&properties)?;

        frame_properties.left = left_property(&properties)?;

        // TODO: Implement mouse_icon picture loading.
        // mouse_icon: None,
        // frame_properties.mouse_icon = None;

        frame_properties.mouse_pointer = mouse_pointer_property(&properties)?;

        frame_properties.ole_drop_mode = ole_drop_mode_property(&properties)?;

        frame_properties.right_to_left = right_to_left_property(&properties)?;

        frame_properties.tab_index = tab_index_property(&properties)?;

        let tooltiptext_key = BStr::new("ToolTipText");
        if properties.contains_key(tooltiptext_key) {
            frame_properties.tool_tip_text = properties[tooltiptext_key];
        }

        frame_properties.top = top_property(&properties)?;

        frame_properties.visible = visible_property(&properties)?;

        frame_properties.whats_this_help_id = whats_this_help_id_property(&properties)?;

        frame_properties.width = width_property(&properties)?;

        Ok(frame_properties)
    }
}

fn appearance_property(properties: &HashMap<&BStr, &BStr>) -> Result<Appearance, VB6ErrorKind> {
    let appearance_key = BStr::new("Appearance");
    if !properties.contains_key(appearance_key) {
        return Ok(Appearance::ThreeD);
    }

    let appearance = properties[appearance_key];

    match appearance.as_bytes() {
        b"0" => Ok(Appearance::Flat),
        b"1" => Ok(Appearance::ThreeD),
        _ => Err(VB6ErrorKind::AppearancePropertyInvalid),
    }
}

fn back_color_properties(properties: &HashMap<&BStr, &BStr>) -> Result<VB6Color, VB6ErrorKind> {
    let back_color_key = BStr::new("BackColor");
    if !properties.contains_key(back_color_key) {
        return Ok(VB6Color::from_hex("&H8000000F&").unwrap());
    }

    let color_ascii = properties[back_color_key];

    match VB6Color::from_hex(color_ascii.to_str().unwrap()) {
        Ok(color) => Ok(color),
        Err(_) => Err(VB6ErrorKind::HexColorParseError),
    }
}

fn border_style_property(properties: &HashMap<&BStr, &BStr>) -> Result<BorderStyle, VB6ErrorKind> {
    let border_style_key = BStr::new("BorderStyle");
    if !properties.contains_key(border_style_key) {
        return Ok(BorderStyle::FixedSingle);
    }

    let border_style = properties[border_style_key];

    match border_style.as_bytes() {
        b"0" => Ok(BorderStyle::None),
        b"1" => Ok(BorderStyle::FixedSingle),
        _ => Err(VB6ErrorKind::BorderStylePropertyInvalid),
    }
}

fn clip_controls_property(
    properties: &HashMap<&BStr, &BStr>,
) -> Result<ClipControls, VB6ErrorKind> {
    let clip_controls_key = BStr::new("ClipControls");
    if !properties.contains_key(clip_controls_key) {
        return Ok(ClipControls::default());
    }

    let clip_controls = properties[clip_controls_key];

    match clip_controls.as_bytes() {
        b"0" => Ok(ClipControls::False),
        b"1" => Ok(ClipControls::True),
        _ => Err(VB6ErrorKind::ClipControlsPropertyInvalid),
    }
}

fn drag_mode_property(properties: &HashMap<&BStr, &BStr>) -> Result<DragMode, VB6ErrorKind> {
    let drag_mode_key = BStr::new("DragMode");
    if !properties.contains_key(drag_mode_key) {
        return Ok(DragMode::Manual);
    }

    let drag_mode = properties[drag_mode_key];

    match drag_mode.as_bytes() {
        b"0" => Ok(DragMode::Manual),
        b"1" => Ok(DragMode::Automatic),
        _ => Err(VB6ErrorKind::DragModePropertyInvalid),
    }
}

fn enabled_property(properties: &HashMap<&BStr, &BStr>) -> Result<bool, VB6ErrorKind> {
    let enabled_key = BStr::new("Enabled");
    if !properties.contains_key(enabled_key) {
        return Ok(true);
    }

    let enabled = properties[enabled_key];

    match enabled.as_bytes() {
        b"0" => Ok(false),
        b"1" => Ok(true),
        _ => Err(VB6ErrorKind::EnabledPropertyInvalid),
    }
}

fn fore_color_properties(properties: &HashMap<&BStr, &BStr>) -> Result<VB6Color, VB6ErrorKind> {
    let fore_color_key = BStr::new("ForeColor");
    if !properties.contains_key(fore_color_key) {
        return Ok(VB6Color::from_hex("&H80000012&").unwrap());
    }

    let color_ascii = properties[fore_color_key];

    match VB6Color::from_hex(color_ascii.to_str().unwrap()) {
        Ok(color) => Ok(color),
        Err(_) => Err(VB6ErrorKind::HexColorParseError),
    }
}

fn height_property(properties: &HashMap<&BStr, &BStr>) -> Result<i32, VB6ErrorKind> {
    let height_key = BStr::new("Height");
    if !properties.contains_key(height_key) {
        return Ok(30);
    }

    let height = properties[height_key];

    match height.to_str().unwrap().parse::<i32>() {
        Ok(height) => Ok(height),
        Err(_) => Err(VB6ErrorKind::PropertyValueAsciiConversionError),
    }
}

fn help_context_id_property(properties: &HashMap<&BStr, &BStr>) -> Result<i32, VB6ErrorKind> {
    let help_context_id_key = BStr::new("HelpContextID");
    if !properties.contains_key(help_context_id_key) {
        return Ok(0);
    }

    let help_context_id = properties[help_context_id_key];

    match help_context_id.to_str().unwrap().parse::<i32>() {
        Ok(help_context_id) => Ok(help_context_id),
        Err(_) => Err(VB6ErrorKind::PropertyValueAsciiConversionError),
    }
}

fn left_property(properties: &HashMap<&BStr, &BStr>) -> Result<i32, VB6ErrorKind> {
    let left_key = BStr::new("Left");
    if !properties.contains_key(left_key) {
        return Ok(30);
    }

    let left = properties[left_key];

    match left.to_str().unwrap().parse::<i32>() {
        Ok(left) => Ok(left),
        Err(_) => Err(VB6ErrorKind::PropertyValueAsciiConversionError),
    }
}

fn mouse_pointer_property(
    properties: &HashMap<&BStr, &BStr>,
) -> Result<MousePointer, VB6ErrorKind> {
    let mouse_pointer_key = BStr::new("MousePointer");
    if !properties.contains_key(mouse_pointer_key) {
        return Ok(MousePointer::Default);
    }

    let mouse_pointer = properties[mouse_pointer_key];

    match mouse_pointer.as_bytes() {
        b"0" => Ok(MousePointer::Default),
        b"1" => Ok(MousePointer::Arrow),
        b"2" => Ok(MousePointer::Cross),
        b"3" => Ok(MousePointer::IBeam),
        b"6" => Ok(MousePointer::SizeNESW),
        b"7" => Ok(MousePointer::SizeNS),
        b"8" => Ok(MousePointer::SizeNWSE),
        b"9" => Ok(MousePointer::SizeWE),
        b"10" => Ok(MousePointer::UpArrow),
        b"11" => Ok(MousePointer::Hourglass),
        b"12" => Ok(MousePointer::NoDrop),
        b"13" => Ok(MousePointer::ArrowHourglass),
        b"14" => Ok(MousePointer::ArrowQuestion),
        b"15" => Ok(MousePointer::SizeAll),
        b"99" => Ok(MousePointer::Custom),
        _ => Err(VB6ErrorKind::MousePointerPropertyInvalid),
    }
}

fn ole_drop_mode_property(properties: &HashMap<&BStr, &BStr>) -> Result<OLEDropMode, VB6ErrorKind> {
    let ole_drop_mode_key = BStr::new("OLEDropMode");
    if !properties.contains_key(ole_drop_mode_key) {
        return Ok(OLEDropMode::default());
    }

    let ole_drop_mode = properties[ole_drop_mode_key];

    match ole_drop_mode.as_bytes() {
        b"0" => Ok(OLEDropMode::None),
        b"1" => Ok(OLEDropMode::Manual),
        _ => Err(VB6ErrorKind::OLEDropModePropertyInvalid),
    }
}

fn right_to_left_property(properties: &HashMap<&BStr, &BStr>) -> Result<bool, VB6ErrorKind> {
    let right_to_left_key = BStr::new("RightToLeft");
    if !properties.contains_key(right_to_left_key) {
        return Ok(false);
    }

    let right_to_left = properties[right_to_left_key];

    match right_to_left.as_bytes() {
        b"0" => Ok(false),
        b"1" => Ok(true),
        _ => Err(VB6ErrorKind::RightToLeftPropertyInvalid),
    }
}

fn tab_index_property(properties: &HashMap<&BStr, &BStr>) -> Result<i32, VB6ErrorKind> {
    let tab_index_key = BStr::new("TabIndex");
    if !properties.contains_key(tab_index_key) {
        return Ok(0);
    }

    let tab_index = properties[tab_index_key];

    match tab_index.to_str().unwrap().parse::<i32>() {
        Ok(tab_index) => Ok(tab_index),
        Err(_) => Err(VB6ErrorKind::PropertyValueAsciiConversionError),
    }
}

fn top_property(properties: &HashMap<&BStr, &BStr>) -> Result<i32, VB6ErrorKind> {
    let top_key = BStr::new("Top");
    if !properties.contains_key(top_key) {
        return Ok(0);
    }

    let top = properties[top_key];

    match top.to_str().unwrap().parse::<i32>() {
        Ok(top) => Ok(top),
        Err(_) => Err(VB6ErrorKind::PropertyValueAsciiConversionError),
    }
}

fn visible_property(properties: &HashMap<&BStr, &BStr>) -> Result<bool, VB6ErrorKind> {
    let visible_key = BStr::new("Visible");
    if !properties.contains_key(visible_key) {
        return Ok(true);
    }

    let visible = properties[visible_key];

    match visible.as_bytes() {
        b"0" => Ok(false),
        b"1" => Ok(true),
        _ => Err(VB6ErrorKind::VisiblePropertyInvalid),
    }
}

fn whats_this_help_id_property(properties: &HashMap<&BStr, &BStr>) -> Result<i32, VB6ErrorKind> {
    let help_id_key = BStr::new("WhatsThisHelpID");
    if !properties.contains_key(help_id_key) {
        return Ok(0);
    }

    let help_id = properties[help_id_key];

    match help_id.to_str().unwrap().parse::<i32>() {
        Ok(help_id) => Ok(help_id),
        Err(_) => return Err(VB6ErrorKind::PropertyValueAsciiConversionError),
    }
}

fn width_property(properties: &HashMap<&BStr, &BStr>) -> Result<i32, VB6ErrorKind> {
    let width_key = BStr::new("Width");
    if !properties.contains_key(width_key) {
        return Ok(100);
    }

    let width = properties[width_key];

    match width.to_str().unwrap().parse::<i32>() {
        Ok(width) => Ok(width),
        Err(_) => Err(VB6ErrorKind::PropertyValueAsciiConversionError),
    }
}
