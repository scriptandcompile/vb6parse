use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::color::VB6Color;
use crate::language::controls::{Appearance, BorderStyle, DragMode, MousePointer, OLEDropMode};
use crate::parsers::form::VB6PropertyGroup;

use image::DynamicImage;

#[derive(Debug, PartialEq, Clone)]
pub struct FrameProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub caption: &'a str,
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
    pub tool_tip_text: &'a str,
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
            caption: "Frame1",
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
            tool_tip_text: "",
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl<'a> FrameProperties<'a> {
    pub fn construct_control(
        properties: HashMap<&'a str, &'a str>,
        _property_groups: Vec<VB6PropertyGroup<'a>>,
    ) -> Result<Self, VB6ErrorKind> {
        // TODO: We are not correctly handling property assignment for each control.

        let mut frame_properties = FrameProperties::default();

        if properties.contains_key("Appearance") {
            let appearance = properties["Appearance"];

            frame_properties.appearance = match appearance {
                "0" => Appearance::Flat,
                "1" => Appearance::ThreeD,
                _ => {
                    return Err(VB6ErrorKind::AppearancePropertyInvalid);
                }
            };
        }

        if properties.contains_key("BackColor") {
            let color_ascii = properties["BackColor"];

            let Ok(back_color) = VB6Color::from_hex(color_ascii) else {
                return Err(VB6ErrorKind::HexColorParseError);
            };
            frame_properties.back_color = back_color;
        }

        if properties.contains_key("BorderStyle") {
            let border_style = properties["BorderStyle"];

            frame_properties.border_style = match border_style {
                "0" => BorderStyle::None,
                "1" => BorderStyle::FixedSingle,
                _ => {
                    return Err(VB6ErrorKind::BorderStylePropertyInvalid);
                }
            };
        }

        if properties.contains_key("Caption") {
            frame_properties.caption = properties["Caption"];
        }

        if properties.contains_key("ClipControls") {
            let clip_controls = properties["ClipControls"];

            frame_properties.clip_controls = match clip_controls {
                "0" => false,
                "1" => true,
                _ => {
                    return Err(VB6ErrorKind::ClipControlsPropertyInvalid);
                }
            };
        }

        // TODO: Implement loading drag_icon picture loading.
        // drag_icon: None,

        if properties.contains_key("DragMode") {
            let drag_mode = properties["DragMode"];

            frame_properties.drag_mode = match drag_mode {
                "0" => DragMode::Manual,
                "1" => DragMode::Automatic,
                _ => {
                    return Err(VB6ErrorKind::DragModePropertyInvalid);
                }
            };
        }

        if properties.contains_key("Enabled") {
            let enabled = properties["Enabled"];

            frame_properties.enabled = match enabled {
                "0" => false,
                "1" => true,
                _ => {
                    return Err(VB6ErrorKind::EnabledPropertyInvalid);
                }
            };
        }

        if properties.contains_key("ForeColor") {
            let color_ascii = properties["ForeColor"];

            let Ok(fore_color) = VB6Color::from_hex(color_ascii) else {
                return Err(VB6ErrorKind::HexColorParseError);
            };
            frame_properties.fore_color = fore_color;
        }

        if properties.contains_key("Height") {
            let height = properties["Height"];

            let Ok(height) = height.parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.height = height;
        }

        if properties.contains_key("HelpContextID") {
            let help_context_id = properties["HelpContextID"];

            let Ok(help_context_id) = help_context_id.parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.help_context_id = help_context_id;
        }

        if properties.contains_key("Left") {
            let left = properties["Left"];

            let Ok(left) = left.parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.left = left;
        }

        // TODO: Implement mouse_icon picture loading.
        // mouse_icon: None,

        if properties.contains_key("MousePointer") {
            let mouse_pointer = properties["MousePointer"];

            frame_properties.mouse_pointer = match mouse_pointer {
                "0" => MousePointer::Default,
                "1" => MousePointer::Arrow,
                "2" => MousePointer::Cross,
                "3" => MousePointer::IBeam,
                "6" => MousePointer::SizeNESW,
                "7" => MousePointer::SizeNS,
                "8" => MousePointer::SizeNWSE,
                "9" => MousePointer::SizeWE,
                "10" => MousePointer::UpArrow,
                "11" => MousePointer::Hourglass,
                "12" => MousePointer::NoDrop,
                "13" => MousePointer::ArrowHourglass,
                "14" => MousePointer::ArrowQuestion,
                "15" => MousePointer::SizeAll,
                "99" => MousePointer::Custom,
                _ => {
                    return Err(VB6ErrorKind::MousePointerPropertyInvalid);
                }
            };
        }

        if properties.contains_key("OLEDropMode") {
            let ole_drop_mode = properties["OLEDropMode"];

            frame_properties.ole_drop_mode = match ole_drop_mode {
                "0" => OLEDropMode::None,
                "1" => OLEDropMode::Manual,
                _ => {
                    return Err(VB6ErrorKind::OLEDropModePropertyInvalid);
                }
            };
        }

        if properties.contains_key("RightToLeft") {
            let right_to_left = properties["RightToLeft"];

            frame_properties.right_to_left = match right_to_left {
                "0" => false,
                "1" => true,
                _ => {
                    return Err(VB6ErrorKind::RightToLeftPropertyInvalid);
                }
            };
        }

        if properties.contains_key("TabIndex") {
            let tab_index = properties["TabIndex"];

            let Ok(tab_index) = tab_index.parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.tab_index = tab_index;
        }

        if properties.contains_key("ToolTipText") {
            frame_properties.tool_tip_text = properties["ToolTipText"];
        }

        if properties.contains_key("Top") {
            let top = properties["Top"];

            let Ok(top) = top.parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.top = top;
        }

        if properties.contains_key("Visible") {
            let visible = properties["Visible"];

            frame_properties.visible = match visible {
                "0" => false,
                "1" => true,
                _ => {
                    return Err(VB6ErrorKind::VisiblePropertyInvalid);
                }
            };
        }

        if properties.contains_key("WhatsThisHelpID") {
            let whats_this_help_id = properties["WhatsThisHelpID"];

            let Ok(whats_this_help_id) = whats_this_help_id.parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.whats_this_help_id = whats_this_help_id;
        }

        if properties.contains_key("Width") {
            let width = properties["Width"];

            let Ok(width) = width.parse::<i32>() else {
                return Err(VB6ErrorKind::PropertyValueAsciiConversionError);
            };

            frame_properties.width = width;
        }

        Ok(frame_properties)
    }
}
