use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::color::VB6Color;
use crate::language::controls::{Appearance, BorderStyle, DragMode, MousePointer, OLEDropMode};
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
    pub clip_controls: bool,
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
            clip_controls: true,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            left: 30,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::None,
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

        let appearance_key = BStr::new("Appearance");
        if properties.contains_key(appearance_key) {
            let appearance = properties[appearance_key];

            frame_properties.appearance = match appearance.as_bytes() {
                b"0" => Appearance::Flat,
                b"1" => Appearance::ThreeD,
                _ => {
                    return Err(VB6ErrorKind::AppearancePropertyInvalid);
                }
            };
        }

        let backcolor_key = BStr::new("BackColor");
        if properties.contains_key(backcolor_key) {
            let color_ascii = properties[backcolor_key].to_str().unwrap();

            let Ok(back_color) = VB6Color::from_hex(color_ascii) else {
                return Err(VB6ErrorKind::HexColorParseError);
            };
            frame_properties.back_color = back_color;
        }

        let borderstyle_key = BStr::new("BorderStyle");
        if properties.contains_key(borderstyle_key) {
            let border_style = properties[borderstyle_key];

            frame_properties.border_style = match border_style.as_bytes() {
                b"0" => BorderStyle::None,
                b"1" => BorderStyle::FixedSingle,
                _ => {
                    return Err(VB6ErrorKind::BorderStylePropertyInvalid);
                }
            };
        }

        let caption_key = BStr::new("Caption");
        if properties.contains_key(caption_key) {
            frame_properties.caption = properties[caption_key];
        }

        let clipcontrols_key = BStr::new("ClipControls");
        if properties.contains_key(clipcontrols_key) {
            let clip_controls = properties[clipcontrols_key];

            frame_properties.clip_controls = match clip_controls.as_bytes() {
                b"0" => false,
                b"1" => true,
                _ => {
                    return Err(VB6ErrorKind::ClipControlsPropertyInvalid);
                }
            };
        }

        // TODO: Implement loading drag_icon picture loading.
        // drag_icon: None,

        let dragmode_key = BStr::new("DragMode");
        if properties.contains_key(dragmode_key) {
            let drag_mode = properties[dragmode_key];

            frame_properties.drag_mode = match drag_mode.as_bytes() {
                b"0" => DragMode::Manual,
                b"1" => DragMode::Automatic,
                _ => {
                    return Err(VB6ErrorKind::DragModePropertyInvalid);
                }
            };
        }

        let enabled_key = BStr::new("Enabled");
        if properties.contains_key(enabled_key) {
            let enabled = properties[enabled_key];

            frame_properties.enabled = match enabled.as_bytes() {
                b"0" => false,
                b"1" => true,
                _ => {
                    return Err(VB6ErrorKind::EnabledPropertyInvalid);
                }
            };
        }

        let forecolor_key = BStr::new("ForeColor");
        if properties.contains_key(forecolor_key) {
            let color_ascii = properties[forecolor_key];

            let Ok(fore_color) = VB6Color::from_hex(color_ascii.to_str().unwrap()) else {
                return Err(VB6ErrorKind::HexColorParseError);
            };
            frame_properties.fore_color = fore_color;
        }

        let height_key = BStr::new("Height");
        if properties.contains_key(height_key) {
            let height = properties[height_key];

            let Ok(height) = height.to_str().unwrap().parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.height = height;
        }

        let helpcontextid_key = BStr::new("HelpContextID");
        if properties.contains_key(helpcontextid_key) {
            let help_context_id = properties[helpcontextid_key];

            let Ok(help_context_id) = help_context_id.to_str().unwrap().parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.help_context_id = help_context_id;
        }

        let left_key = BStr::new("Left");
        if properties.contains_key(left_key) {
            let left = properties[left_key];

            let Ok(left) = left.to_str().unwrap().parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.left = left;
        }

        // TODO: Implement mouse_icon picture loading.
        // mouse_icon: None,
        let mousepointer_key = BStr::new("MousePointer");
        if properties.contains_key(mousepointer_key) {
            let mouse_pointer = properties[mousepointer_key];

            frame_properties.mouse_pointer = match mouse_pointer.as_bytes() {
                b"0" => MousePointer::Default,
                b"1" => MousePointer::Arrow,
                b"2" => MousePointer::Cross,
                b"3" => MousePointer::IBeam,
                b"6" => MousePointer::SizeNESW,
                b"7" => MousePointer::SizeNS,
                b"8" => MousePointer::SizeNWSE,
                b"9" => MousePointer::SizeWE,
                b"10" => MousePointer::UpArrow,
                b"11" => MousePointer::Hourglass,
                b"12" => MousePointer::NoDrop,
                b"13" => MousePointer::ArrowHourglass,
                b"14" => MousePointer::ArrowQuestion,
                b"15" => MousePointer::SizeAll,
                b"99" => MousePointer::Custom,
                _ => {
                    return Err(VB6ErrorKind::MousePointerPropertyInvalid);
                }
            };
        }

        let oledropmode_key = BStr::new("OLEDropMode");
        if properties.contains_key(oledropmode_key) {
            let ole_drop_mode = properties[oledropmode_key];

            frame_properties.ole_drop_mode = match ole_drop_mode.as_bytes() {
                b"0" => OLEDropMode::None,
                b"1" => OLEDropMode::Manual,
                _ => {
                    return Err(VB6ErrorKind::OLEDropModePropertyInvalid);
                }
            };
        }

        frame_properties.right_to_left = right_to_left_property(&properties)?;

        let right_to_left_key = BStr::new("RightToLeft");
        if properties.contains_key(right_to_left_key) {
            let right_to_left = properties[right_to_left_key];

            frame_properties.right_to_left = match right_to_left.as_bytes() {
                b"0" => false,
                b"1" => true,
                _ => {
                    return Err(VB6ErrorKind::RightToLeftPropertyInvalid);
                }
            };
        }

        let tabindex_key = BStr::new("TabIndex");
        if properties.contains_key(tabindex_key) {
            let tab_index = properties[tabindex_key];

            let Ok(tab_index) = tab_index.to_str().unwrap().parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.tab_index = tab_index;
        }

        let tooltiptext_key = BStr::new("ToolTipText");
        if properties.contains_key(tooltiptext_key) {
            frame_properties.tool_tip_text = properties[tooltiptext_key];
        }

        let top_key = BStr::new("Top");
        if properties.contains_key(top_key) {
            let top = properties[top_key];

            let Ok(top) = top.to_str().unwrap().parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.top = top;
        }

        let visible_key = BStr::new("Visible");
        if properties.contains_key(visible_key) {
            let visible = properties[visible_key];

            frame_properties.visible = match visible.as_bytes() {
                b"0" => false,
                b"1" => true,
                _ => {
                    return Err(VB6ErrorKind::VisiblePropertyInvalid);
                }
            };
        }

        frame_properties.whats_this_help_id = whats_this_help_id_property(properties)?;

        frame_properties.width = width_property(properties)?;

        Ok(frame_properties)
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
