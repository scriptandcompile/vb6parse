//! Properties for `PictureBox` controls.
//!
//! This is used as an enum variant of
//! [`ControlKind::PictureBox`](crate::language::controls::ControlKind::PictureBox).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use crate::{
    files::common::Properties,
    language::{
        color::{Color, VB_BUTTON_FACE, VB_BUTTON_TEXT, VB_SCROLL_BARS},
        controls::{
            Activation, Align, Appearance, AutoRedraw, AutoSize, BorderStyle, CausesValidation,
            ClipControls, DragMode, DrawMode, DrawStyle, FillStyle, Font, FontTransparency,
            HasDeviceContext, LinkMode, MousePointer, OLEDragMode, OLEDropMode, ReferenceOrValue,
            ScaleMode, TabStop, TextDirection, Visibility,
        },
    },
};

use image::DynamicImage;
use serde::Serialize;

/// Properties for a `PictureBox` control.
///
/// This is used as an enum variant of
/// [`ControlKind::PictureBox`](crate::language::controls::ControlKind::PictureBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct PictureBoxProperties {
    /// Alignment of the `PictureBox`.
    pub align: Align,
    /// Appearance of the `PictureBox`.
    pub appearance: Appearance,
    /// `AutoRedraw` setting of the `PictureBox`.
    pub auto_redraw: AutoRedraw,
    /// `AutoSize` setting of the `PictureBox`.
    pub auto_size: AutoSize,
    /// Background color of the `PictureBox`.
    pub back_color: Color,
    /// Border style of the `PictureBox`.
    pub border_style: BorderStyle,
    /// Indicates whether the `PictureBox` causes validation.
    pub causes_validation: CausesValidation,
    /// The `ClipControls` setting of the `PictureBox`.
    pub clip_controls: ClipControls,
    /// Data field associated with the `PictureBox`.
    pub data_field: String,
    /// Data format associated with the `PictureBox`.
    pub data_format: String,
    /// Data member associated with the `PictureBox`.
    pub data_member: String,
    /// Data source associated with the `PictureBox`.
    pub data_source: String,
    /// Drag icon for the `PictureBox`.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Drag mode of the `PictureBox`.
    pub drag_mode: DragMode,
    /// Draw mode of the `PictureBox`.
    pub draw_mode: DrawMode,
    /// Draw style of the `PictureBox`.
    pub draw_style: DrawStyle,
    /// Width of the drawing pen.
    pub draw_width: i32,
    /// Indicates whether the `PictureBox` is enabled.
    pub enabled: Activation,
    /// The font style for the form.
    pub font: Option<Font>,
    /// Fill color of the `PictureBox`.
    pub fill_color: Color,
    /// Fill style of the `PictureBox`.
    pub fill_style: FillStyle,
    /// Font transparency setting of the `PictureBox`.
    pub font_transparent: FontTransparency,
    /// Foreground color of the `PictureBox`.
    pub fore_color: Color,
    /// Indicates whether the `PictureBox` has a device context.
    pub has_dc: HasDeviceContext,
    /// Height of the `PictureBox`.
    pub height: i32,
    /// Help context ID of the `PictureBox`.
    pub help_context_id: i32,
    /// Left position of the `PictureBox`.
    pub left: i32,
    /// Link item associated with the `PictureBox`.
    pub link_item: String,
    /// Link mode of the `PictureBox`.
    pub link_mode: LinkMode,
    /// Link timeout of the `PictureBox`.
    pub link_timeout: i32,
    /// Link topic associated with the `PictureBox`.
    pub link_topic: String,
    /// Mouse icon for the `PictureBox`.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// Mouse pointer style of the `PictureBox`.
    pub mouse_pointer: MousePointer,
    /// Indicates whether the `PictureBox` negotiates OLE drag-and-drop operations.
    pub negotiate: bool,
    /// OLE drag mode of the `PictureBox`.
    pub ole_drag_mode: OLEDragMode,
    /// OLE drop mode of the `PictureBox`.
    pub ole_drop_mode: OLEDropMode,
    /// Picture displayed in the `PictureBox`.
    pub picture: Option<ReferenceOrValue<DynamicImage>>,
    /// Text direction of the `PictureBox`.
    pub right_to_left: TextDirection,
    /// Scale height of the `PictureBox`.
    pub scale_height: i32,
    /// Scale left position of the `PictureBox`.
    pub scale_left: i32,
    /// Scale mode of the `PictureBox`.
    pub scale_mode: ScaleMode,
    /// Scale top position of the `PictureBox`.
    pub scale_top: i32,
    /// Scale width of the `PictureBox`.
    pub scale_width: i32,
    /// Tab index of the `PictureBox`.
    pub tab_index: i32,
    /// Tab stop setting of the `PictureBox`.
    pub tab_stop: TabStop,
    /// Tool tip text of the `PictureBox`.
    pub tool_tip_text: String,
    /// Top position of the `PictureBox`.
    pub top: i32,
    /// Visibility of the `PictureBox`.
    pub visible: Visibility,
    /// "What's This?" help ID of the `PictureBox`.
    pub whats_this_help_id: i32,
    /// Width of the `PictureBox`.
    pub width: i32,
}

impl Default for PictureBoxProperties {
    fn default() -> Self {
        PictureBoxProperties {
            align: Align::None,
            appearance: Appearance::ThreeD,
            auto_redraw: AutoRedraw::Manual,
            auto_size: AutoSize::Fixed,
            back_color: VB_BUTTON_FACE,
            border_style: BorderStyle::FixedSingle,
            causes_validation: CausesValidation::Yes,
            clip_controls: ClipControls::default(),
            data_field: String::new(),
            data_format: String::new(),
            data_member: String::new(),
            data_source: String::new(),
            drag_icon: None,
            drag_mode: DragMode::Manual,
            draw_mode: DrawMode::CopyPen,
            draw_style: DrawStyle::Solid,
            draw_width: 1,
            enabled: Activation::Enabled,
            fill_color: VB_SCROLL_BARS,
            fill_style: FillStyle::Transparent,
            font: Some(Font::default()),
            font_transparent: FontTransparency::Transparent,
            fore_color: VB_BUTTON_TEXT,
            has_dc: HasDeviceContext::HasContext,
            height: 30,
            help_context_id: 0,
            left: 30,
            link_item: String::new(),
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: String::new(),
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
            tool_tip_text: String::new(),
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

impl From<Properties> for PictureBoxProperties {
    fn from(prop: Properties) -> Self {
        let mut picture_box_prop = PictureBoxProperties::default();

        picture_box_prop.align = prop.get_property("Align", picture_box_prop.align);
        picture_box_prop.appearance = prop.get_property("Appearance", picture_box_prop.appearance);
        picture_box_prop.auto_redraw =
            prop.get_property("AutoRedraw", picture_box_prop.auto_redraw);
        picture_box_prop.auto_size = prop.get_property("AutoSize", picture_box_prop.auto_size);
        picture_box_prop.back_color = prop.get_color("BackColor", picture_box_prop.back_color);
        picture_box_prop.border_style =
            prop.get_property("BorderStyle", picture_box_prop.border_style);
        picture_box_prop.causes_validation =
            prop.get_property("CausesValidation", picture_box_prop.causes_validation);
        picture_box_prop.clip_controls =
            prop.get_property("ClipControls", picture_box_prop.clip_controls);
        picture_box_prop.data_field = match prop.get("DataField") {
            Some(data_field) => data_field.into(),
            None => picture_box_prop.data_field,
        };
        picture_box_prop.data_format = match prop.get("DataFormat") {
            Some(data_format) => data_format.into(),
            None => picture_box_prop.data_format,
        };
        picture_box_prop.data_member = match prop.get("DataMember") {
            Some(data_member) => data_member.into(),
            None => picture_box_prop.data_member,
        };
        picture_box_prop.data_source = match prop.get("DataSource") {
            Some(data_source) => data_source.into(),
            None => picture_box_prop.data_source,
        };

        // DragIcon

        picture_box_prop.drag_mode = prop.get_property("DragMode", picture_box_prop.drag_mode);
        picture_box_prop.draw_mode = prop.get_property("DrawMode", picture_box_prop.draw_mode);
        picture_box_prop.draw_style = prop.get_property("DrawStyle", picture_box_prop.draw_style);
        picture_box_prop.draw_width = prop.get_i32("DrawWidth", picture_box_prop.draw_width);
        picture_box_prop.enabled = prop.get_property("Enabled", picture_box_prop.enabled);
        picture_box_prop.fill_color = prop.get_color("FillColor", picture_box_prop.fill_color);
        picture_box_prop.fill_style = prop.get_property("FillStyle", picture_box_prop.fill_style);
        picture_box_prop.font_transparent =
            prop.get_property("FontTransparent", picture_box_prop.font_transparent);
        picture_box_prop.fore_color = prop.get_color("ForeColor", picture_box_prop.fore_color);
        picture_box_prop.has_dc = prop.get_property("HasDC", picture_box_prop.has_dc);
        picture_box_prop.height = prop.get_i32("Height", picture_box_prop.height);
        picture_box_prop.help_context_id =
            prop.get_i32("HelpContextID", picture_box_prop.help_context_id);
        picture_box_prop.left = prop.get_i32("Left", picture_box_prop.left);
        picture_box_prop.link_item = match prop.get("LinkItem") {
            Some(link_item) => link_item.into(),
            None => picture_box_prop.link_item,
        };
        picture_box_prop.link_mode = prop.get_property("LinkMode", picture_box_prop.link_mode);
        picture_box_prop.link_timeout = prop.get_i32("LinkTimeout", picture_box_prop.link_timeout);
        picture_box_prop.link_topic = match prop.get("LinkTopic") {
            Some(link_topic) => link_topic.into(),
            None => picture_box_prop.link_topic,
        };

        // TODO: process MouseIcon
        // MouseIcon

        picture_box_prop.mouse_pointer =
            prop.get_property("MousePointer", picture_box_prop.mouse_pointer);
        picture_box_prop.negotiate = prop.get_bool("Negotiate", picture_box_prop.negotiate);
        picture_box_prop.ole_drag_mode =
            prop.get_property("OLEDragMode", picture_box_prop.ole_drag_mode);
        picture_box_prop.ole_drop_mode =
            prop.get_property("OLEDropMode", picture_box_prop.ole_drop_mode);

        // TODO: process Picture
        // Picture

        picture_box_prop.right_to_left =
            prop.get_property("RightToLeft", picture_box_prop.right_to_left);
        picture_box_prop.scale_height = prop.get_i32("ScaleHeight", picture_box_prop.scale_height);
        picture_box_prop.scale_left = prop.get_i32("ScaleLeft", picture_box_prop.scale_left);
        picture_box_prop.scale_mode = prop.get_property("ScaleMode", picture_box_prop.scale_mode);
        picture_box_prop.scale_top = prop.get_i32("ScaleTop", picture_box_prop.scale_top);
        picture_box_prop.scale_width = prop.get_i32("ScaleWidth", picture_box_prop.scale_width);
        picture_box_prop.tab_index = prop.get_i32("TabIndex", picture_box_prop.tab_index);
        picture_box_prop.tab_stop = prop.get_property("TabStop", picture_box_prop.tab_stop);
        picture_box_prop.tool_tip_text = match prop.get("ToolTipText") {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => String::new(),
        };
        picture_box_prop.top = prop.get_i32("Top", picture_box_prop.top);
        picture_box_prop.visible = prop.get_property("Visible", picture_box_prop.visible);
        picture_box_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", picture_box_prop.whats_this_help_id);
        picture_box_prop.width = prop.get_i32("Width", picture_box_prop.width);

        picture_box_prop
    }
}
