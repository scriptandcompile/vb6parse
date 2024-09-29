use std::collections::HashMap;

use crate::errors::VB6ErrorKind;
use crate::language::controls::{Align, Appearance, DragMode, MousePointer, OLEDropMode};
use crate::parsers::form::{
    build_bool_property, build_color_property, build_i32_property, build_property,
};
use crate::VB6Color;

use bstr::{BStr, ByteSlice};
use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// Properties for a `Data` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Data`](crate::language::controls::VB6ControlKind::Data).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct DataProperties<'a> {
    pub align: Align,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub bof_action: BOFAction,
    pub caption: &'a BStr,
    pub connection: Connection,
    pub database_name: &'a BStr,
    pub default_cursor_type: DefaultCursorType,
    pub default_type: DefaultType,
    pub drag_icon: Option<DynamicImage>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub eof_action: EOFAction,
    pub exclusive: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub left: i32,
    pub mouse_icon: Option<DynamicImage>,
    pub mouse_pointer: MousePointer,
    pub negotitate: bool,
    pub ole_drop_mode: OLEDropMode,
    pub options: i32,
    pub read_only: bool,
    pub record_set_type: RecordSetType,
    pub record_source: &'a BStr,
    pub right_to_left: bool,
    pub tool_tip_text: &'a BStr,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for DataProperties<'_> {
    fn default() -> Self {
        DataProperties {
            align: Align::None,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::System { index: 5 },
            bof_action: BOFAction::MoveFirst,
            caption: BStr::new("Data1"),
            connection: Connection::Access,
            database_name: BStr::new(""),
            default_cursor_type: DefaultCursorType::Default,
            default_type: DefaultType::UseJet,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: true,
            eof_action: EOFAction::MoveLast,
            exclusive: false,
            fore_color: VB6Color::System { index: 8 },
            height: 1215,
            left: 480,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            negotitate: false,
            ole_drop_mode: OLEDropMode::default(),
            options: 0,
            read_only: false,
            record_set_type: RecordSetType::Dynaset,
            record_source: BStr::new(""),
            right_to_left: false,
            tool_tip_text: BStr::new(""),
            top: 840,
            visible: true,
            whats_this_help_id: 0,
            width: 1140,
        }
    }
}

impl Serialize for DataProperties<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("DataProperties", 30)?;
        s.serialize_field("align", &self.align)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("bof_action", &self.bof_action)?;
        s.serialize_field("caption", &self.caption)?;
        s.serialize_field("connection", &self.connection)?;
        s.serialize_field("database_name", &self.database_name)?;
        s.serialize_field("default_cursor_type", &self.default_cursor_type)?;
        s.serialize_field("default_type", &self.default_type)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("eof_action", &self.eof_action)?;
        s.serialize_field("exclusive", &self.exclusive)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("left", &self.left)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("negotitate", &self.negotitate)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("options", &self.options)?;
        s.serialize_field("read_only", &self.read_only)?;
        s.serialize_field("record_set_type", &self.record_set_type)?;
        s.serialize_field("record_source", &self.record_source)?;
        s.serialize_field("right_to_left", &self.right_to_left)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl<'a> DataProperties<'a> {
    pub fn construct_control(
        properties: &HashMap<&'a BStr, &'a BStr>,
    ) -> Result<Self, VB6ErrorKind> {
        let mut data_properties = DataProperties::default();

        data_properties.align = build_property(properties, b"Align");
        data_properties.appearance = build_property(properties, b"Appearance");
        data_properties.back_color =
            build_color_property(properties, b"BackColor", data_properties.back_color);
        data_properties.bof_action = build_property(properties, b"BOFAction");
        data_properties.caption = properties
            .get(BStr::new("Caption"))
            .unwrap_or(&data_properties.caption);
        data_properties.connection = properties
            .get(BStr::new("Connection"))
            .map_or(Ok(Connection::Access), |v| {
                Connection::try_from(v.to_str().unwrap_or("Access"))
            })
            .unwrap();
        data_properties.database_name = properties
            .get(BStr::new("DatabaseName"))
            .unwrap_or(&data_properties.database_name);
        data_properties.default_cursor_type = build_property(properties, b"DefaultCursorType");
        data_properties.default_type = build_property(properties, b"DefaultType");

        // DragIcon

        data_properties.drag_mode = build_property(properties, b"DragMode");
        data_properties.enabled =
            build_bool_property(properties, b"Enabled", data_properties.enabled);
        data_properties.eof_action = build_property(properties, b"EOFAction");
        data_properties.exclusive =
            build_bool_property(properties, b"Exclusive", data_properties.exclusive);
        data_properties.fore_color =
            build_color_property(properties, b"ForeColor", data_properties.fore_color);
        data_properties.height = build_i32_property(properties, b"Height", data_properties.height);
        data_properties.left = build_i32_property(properties, b"Left", data_properties.left);
        data_properties.mouse_pointer = build_property(properties, b"MousePointer");
        data_properties.negotitate =
            build_bool_property(properties, b"Negotitate", data_properties.negotitate);
        data_properties.ole_drop_mode = build_property(properties, b"OLEDropMode");
        data_properties.options =
            build_i32_property(properties, b"Options", data_properties.options);
        data_properties.read_only =
            build_bool_property(properties, b"ReadOnly", data_properties.read_only);
        data_properties.record_set_type = build_property(properties, b"RecordsetType");
        data_properties.record_source = properties
            .get(BStr::new("RecordSource"))
            .unwrap_or(&data_properties.record_source);
        data_properties.right_to_left =
            build_bool_property(properties, b"RightToLeft", data_properties.right_to_left);
        data_properties.tool_tip_text = properties
            .get(BStr::new("ToolTipText"))
            .unwrap_or(&data_properties.tool_tip_text);
        data_properties.top = build_i32_property(properties, b"Top", data_properties.top);
        data_properties.visible =
            build_bool_property(properties, b"Visible", data_properties.visible);
        data_properties.whats_this_help_id = build_i32_property(
            properties,
            b"WhatsThisHelpID",
            data_properties.whats_this_help_id,
        );
        data_properties.width = build_i32_property(properties, b"Width", data_properties.width);

        Ok(data_properties)
    }
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum BOFAction {
    #[default]
    MoveFirst = 0,
    Bof = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum Connection {
    #[default]
    Access,
    DBaseIII,
    DBaseIV,
    DBase5_0,
    Excel3_0,
    Excel4_0,
    Excel5_0,
    Excel8_0,
    FoxPro2_0,
    FoxPro2_5,
    FoxPro2_6,
    FoxPro3_0,
    LotusWk1,
    LotusWk3,
    LotusWk4,
    Paradox3X,
    Paradox4X,
    Paradox5X,
    Text,
}

impl TryFrom<&str> for Connection {
    type Error = VB6ErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "Access" => Ok(Connection::Access),
            "dBase III" => Ok(Connection::DBaseIII),
            "dBase IV" => Ok(Connection::DBaseIV),
            "dBase 5.0" => Ok(Connection::DBase5_0),
            "Excel 3.0" => Ok(Connection::Excel3_0),
            "Excel 4.0" => Ok(Connection::Excel4_0),
            "Excel 5.0" => Ok(Connection::Excel5_0),
            "Excel 8.0" => Ok(Connection::Excel8_0),
            "FoxPro 2.0" => Ok(Connection::FoxPro2_0),
            "FoxPro 2.5" => Ok(Connection::FoxPro2_5),
            "FoxPro 2.6" => Ok(Connection::FoxPro2_6),
            "FoxPro 3.0" => Ok(Connection::FoxPro3_0),
            "Lotus WK1" => Ok(Connection::LotusWk1),
            "Lotus WK3" => Ok(Connection::LotusWk3),
            "Lotus WK4" => Ok(Connection::LotusWk4),
            "Paradox 3.X" => Ok(Connection::Paradox3X),
            "Paradox 4.X" => Ok(Connection::Paradox4X),
            "Paradox 5.X" => Ok(Connection::Paradox5X),
            "Text" => Ok(Connection::Text),
            _ => Err(VB6ErrorKind::ConnectionTypeUnparseable),
        }
    }
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum DefaultCursorType {
    #[default]
    Default = 0,
    Odbc = 1,
    ServerSide = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum DefaultType {
    UseODBC = 1,
    #[default]
    UseJet = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum EOFAction {
    #[default]
    MoveLast = 0,
    Eof = 1,
    AddNew = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum RecordSetType {
    Table = 0,
    #[default]
    Dynaset = 1,
    Snapshot = 2,
}
