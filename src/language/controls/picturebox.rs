use crate::language::controls::{
    Activation, Align, Appearance, AutoRedraw, AutoSize, BorderStyle, CausesValidation,
    ClipControls, DragMode, DrawMode, DrawStyle, FillStyle, FontTransparency, HasDeviceContext,
    LinkMode, MousePointer, OLEDragMode, OLEDropMode, ScaleMode, TabStop, TextDirection,
    Visibility,
};
use crate::parsers::Properties;
use crate::VB6Color;

use bstr::BString;
use image::DynamicImage;
use serde::Serialize;

/// Properties for a `PictureBox` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::PictureBox`](crate::language::controls::VB6ControlKind::PictureBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct PictureBoxProperties {
    pub align: Align,
    pub appearance: Appearance,
    pub auto_redraw: AutoRedraw,
    pub auto_size: AutoSize,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub causes_validation: CausesValidation,
    pub clip_controls: ClipControls,
    pub data_field: BString,
    pub data_format: BString,
    pub data_member: BString,
    pub data_source: BString,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub draw_mode: DrawMode,
    pub draw_style: DrawStyle,
    pub draw_width: i32,
    pub enabled: Activation,
    pub fill_color: VB6Color,
    pub fill_style: FillStyle,
    pub font_transparent: FontTransparency,
    pub fore_color: VB6Color,
    pub has_dc: HasDeviceContext,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    pub link_item: BString,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: BString,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub negotiate: bool,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub picture: Option<DynamicImage>,
    pub right_to_left: TextDirection,
    pub scale_height: i32,
    pub scale_left: i32,
    pub scale_mode: ScaleMode,
    pub scale_top: i32,
    pub scale_width: i32,
    pub tab_index: i32,
    pub tab_stop: TabStop,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: Visibility,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for PictureBoxProperties {
    fn default() -> Self {
        PictureBoxProperties {
            align: Align::None,
            appearance: Appearance::ThreeD,
            auto_redraw: AutoRedraw::Manual,
            auto_size: AutoSize::Fixed,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            causes_validation: CausesValidation::Yes,
            clip_controls: ClipControls::default(),
            data_field: "".into(),
            data_format: "".into(),
            data_member: "".into(),
            data_source: "".into(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            draw_mode: DrawMode::CopyPen,
            draw_style: DrawStyle::Solid,
            draw_width: 1,
            enabled: Activation::Enabled,
            fill_color: VB6Color::from_hex("&H00000000&").unwrap(),
            fill_style: FillStyle::Transparent,
            font_transparent: FontTransparency::Transparent,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            has_dc: HasDeviceContext::Yes,
            height: 30,
            help_context_id: 0,
            left: 30,
            link_item: "".into(),
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "".into(),
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            negotiate: false,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            picture: None,
            right_to_left: TextDirection::LeftToRight,
            scale_height: 100,
            scale_left: 0,
            scale_mode: ScaleMode::Twip,
            scale_top: 0,
            scale_width: 100,
            tab_index: 0,
            tab_stop: TabStop::Included,
            tool_tip_text: "".into(),
            top: 30,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Serialize for PictureBoxProperties {
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

impl<'a> From<Properties<'a>> for PictureBoxProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut picture_box_prop = PictureBoxProperties::default();

        picture_box_prop.align = prop.get_property(b"Align".into(), picture_box_prop.align);
        picture_box_prop.appearance =
            prop.get_property(b"Appearance".into(), picture_box_prop.appearance);
        picture_box_prop.auto_redraw =
            prop.get_property(b"AutoRedraw".into(), picture_box_prop.auto_redraw);
        picture_box_prop.auto_size =
            prop.get_property(b"AutoSize".into(), picture_box_prop.auto_size);
        picture_box_prop.back_color =
            prop.get_color(b"BackColor".into(), picture_box_prop.back_color);
        picture_box_prop.border_style =
            prop.get_property(b"BorderStyle".into(), picture_box_prop.border_style);
        picture_box_prop.causes_validation = prop.get_property(
            b"CausesValidation".into(),
            picture_box_prop.causes_validation,
        );
        picture_box_prop.clip_controls =
            prop.get_property(b"ClipControls".into(), picture_box_prop.clip_controls);
        picture_box_prop.data_field = match prop.get(b"DataField".into()) {
            Some(data_field) => data_field.into(),
            None => picture_box_prop.data_field,
        };
        picture_box_prop.data_format = match prop.get(b"DataFormat".into()) {
            Some(data_format) => data_format.into(),
            None => picture_box_prop.data_format,
        };
        picture_box_prop.data_member = match prop.get(b"DataMember".into()) {
            Some(data_member) => data_member.into(),
            None => picture_box_prop.data_member,
        };
        picture_box_prop.data_source = match prop.get(b"DataSource".into()) {
            Some(data_source) => data_source.into(),
            None => picture_box_prop.data_source,
        };

        // DragIcon

        picture_box_prop.drag_mode =
            prop.get_property(b"DragMode".into(), picture_box_prop.drag_mode);
        picture_box_prop.draw_mode =
            prop.get_property(b"DrawMode".into(), picture_box_prop.draw_mode);
        picture_box_prop.draw_style =
            prop.get_property(b"DrawStyle".into(), picture_box_prop.draw_style);
        picture_box_prop.draw_width =
            prop.get_i32(b"DrawWidth".into(), picture_box_prop.draw_width);
        picture_box_prop.enabled = prop.get_property(b"Enabled".into(), picture_box_prop.enabled);
        picture_box_prop.fill_color =
            prop.get_color(b"FillColor".into(), picture_box_prop.fill_color);
        picture_box_prop.fill_style =
            prop.get_property(b"FillStyle".into(), picture_box_prop.fill_style);
        picture_box_prop.font_transparent =
            prop.get_property(b"FontTransparent".into(), picture_box_prop.font_transparent);
        picture_box_prop.fore_color =
            prop.get_color(b"ForeColor".into(), picture_box_prop.fore_color);
        picture_box_prop.has_dc = prop.get_property(b"HasDC".into(), picture_box_prop.has_dc);
        picture_box_prop.height = prop.get_i32(b"Height".into(), picture_box_prop.height);
        picture_box_prop.help_context_id =
            prop.get_i32(b"HelpContextID".into(), picture_box_prop.help_context_id);
        picture_box_prop.left = prop.get_i32(b"Left".into(), picture_box_prop.left);
        picture_box_prop.link_item = match prop.get(b"LinkItem".into()) {
            Some(link_item) => link_item.into(),
            None => picture_box_prop.link_item,
        };
        picture_box_prop.link_mode =
            prop.get_property(b"LinkMode".into(), picture_box_prop.link_mode);
        picture_box_prop.link_timeout =
            prop.get_i32(b"LinkTimeout".into(), picture_box_prop.link_timeout);
        picture_box_prop.link_topic = match prop.get(b"LinkTopic".into()) {
            Some(link_topic) => link_topic.into(),
            None => picture_box_prop.link_topic,
        };

        // MouseIcon

        picture_box_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), picture_box_prop.mouse_pointer);
        picture_box_prop.negotiate = prop.get_bool(b"Negotiate".into(), picture_box_prop.negotiate);
        picture_box_prop.ole_drag_mode =
            prop.get_property(b"OLEDragMode".into(), picture_box_prop.ole_drag_mode);
        picture_box_prop.ole_drop_mode =
            prop.get_property(b"OLEDropMode".into(), picture_box_prop.ole_drop_mode);

        // Picture

        picture_box_prop.right_to_left =
            prop.get_property(b"RightToLeft".into(), picture_box_prop.right_to_left);
        picture_box_prop.scale_height =
            prop.get_i32(b"ScaleHeight".into(), picture_box_prop.scale_height);
        picture_box_prop.scale_left =
            prop.get_i32(b"ScaleLeft".into(), picture_box_prop.scale_left);
        picture_box_prop.scale_mode =
            prop.get_property(b"ScaleMode".into(), picture_box_prop.scale_mode);
        picture_box_prop.scale_top = prop.get_i32(b"ScaleTop".into(), picture_box_prop.scale_top);
        picture_box_prop.scale_width =
            prop.get_i32(b"ScaleWidth".into(), picture_box_prop.scale_width);
        picture_box_prop.tab_index = prop.get_i32(b"TabIndex".into(), picture_box_prop.tab_index);
        picture_box_prop.tab_stop = prop.get_property(b"TabStop".into(), picture_box_prop.tab_stop);
        picture_box_prop.tool_tip_text = match prop.get(b"ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        picture_box_prop.top = prop.get_i32(b"Top".into(), picture_box_prop.top);
        picture_box_prop.visible = prop.get_property(b"Visible".into(), picture_box_prop.visible);
        picture_box_prop.whats_this_help_id = prop.get_i32(
            b"WhatsThisHelpID".into(),
            picture_box_prop.whats_this_help_id,
        );
        picture_box_prop.width = prop.get_i32(b"Width".into(), picture_box_prop.width);

        picture_box_prop
    }
}
