use crate::language::controls::{
    Align, Appearance, BorderStyle, DragMode, DrawMode, DrawStyle, FillStyle, LinkMode,
    MousePointer, OLEDragMode, OLEDropMode, ScaleMode,
};
use crate::VB6Color;

use image::DynamicImage;
use serde::Serialize;

/// Properties for a PictureBox control. This is used as an enum variant of
/// [VB6ControlKind::PictureBox](crate::language::controls::VB6ControlKind::PictureBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [VB6Control](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct PictureBoxProperties<'a> {
    pub align: Align,
    pub appearance: Appearance,
    /// Determines if the output from a graphics method is to a persistent bitmap
    /// which acts as a double buffer.
    pub auto_redraw: bool,
    pub auto_size: bool,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub clip_controls: bool,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
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
    pub left: i32,
    pub link_item: &'a str,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: &'a str,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub negotiate: bool,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub right_to_left: bool,
    pub scale_height: i32,
    pub scale_left: i32,
    pub scale_mode: ScaleMode,
    pub scale_top: i32,
    pub scale_width: i32,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for PictureBoxProperties<'_> {
    fn default() -> Self {
        PictureBoxProperties {
            align: Align::None,
            appearance: Appearance::ThreeD,
            auto_redraw: false,
            auto_size: false,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            causes_validation: true,
            clip_controls: true,
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_icon: None,
            drag_mode: DragMode::Manual,
            draw_mode: DrawMode::CopyPen,
            draw_style: DrawStyle::Solid,
            draw_width: 1,
            enabled: true,
            fill_color: VB6Color::from_hex("&H00000000&").unwrap(),
            fill_style: FillStyle::Solid,
            font_transparent: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            has_dc: true,
            height: 30,
            help_context_id: 0,
            left: 30,
            link_item: "",
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "",
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            negotiate: false,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::None,
            picture: None,
            right_to_left: false,
            scale_height: 100,
            scale_left: 0,
            scale_mode: ScaleMode::Twip,
            scale_top: 0,
            scale_width: 100,
            tab_index: 0,
            tab_stop: true,
            tool_tip_text: "",
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for PictureBoxProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("PictureBoxProperties", 39)?;
        s.serialize_field("align", &self.align)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("auto_redraw", &self.auto_redraw)?;
        s.serialize_field("auto_size", &self.auto_size)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("border_style", &self.border_style)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;
        s.serialize_field("clip_controls", &self.clip_controls)?;
        s.serialize_field("data_field", &self.data_field)?;
        s.serialize_field("data_format", &self.data_format)?;
        s.serialize_field("data_member", &self.data_member)?;
        s.serialize_field("data_source", &self.data_source)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("draw_mode", &self.draw_mode)?;
        s.serialize_field("draw_style", &self.draw_style)?;
        s.serialize_field("draw_width", &self.draw_width)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fill_color", &self.fill_color)?;
        s.serialize_field("fill_style", &self.fill_style)?;
        s.serialize_field("font_transparent", &self.font_transparent)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("has_dc", &self.has_dc)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("left", &self.left)?;
        s.serialize_field("link_item", &self.link_item)?;
        s.serialize_field("link_mode", &self.link_mode)?;
        s.serialize_field("link_timeout", &self.link_timeout)?;
        s.serialize_field("link_topic", &self.link_topic)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("negotiate", &self.negotiate)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;

        let option_text = self.picture.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("picture", &option_text)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("scale_height", &self.scale_height)?;
        s.serialize_field("scale_left", &self.scale_left)?;
        s.serialize_field("scale_mode", &self.scale_mode)?;
        s.serialize_field("scale_top", &self.scale_top)?;
        s.serialize_field("scale_width", &self.scale_width)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}
