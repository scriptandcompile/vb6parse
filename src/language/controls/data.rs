use crate::errors::VB6ErrorKind;
use crate::language::controls::{Align, Appearance, DragMode, MousePointer, OLEDropMode};
use crate::parsers::Properties;
use crate::VB6Color;

use bstr::{BString, ByteSlice};
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
pub struct DataProperties {
    pub align: Align,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub bof_action: BOFAction,
    pub caption: BString,
    pub connection: Connection,
    pub database_name: BString,
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
    pub record_source: BString,
    pub right_to_left: bool,
    pub tool_tip_text: BString,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

impl Default for DataProperties {
    fn default() -> Self {
        DataProperties {
            align: Align::None,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::System { index: 5 },
            bof_action: BOFAction::MoveFirst,
            caption: "".into(),
            connection: Connection::Access,
            database_name: "".into(),
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
            record_source: "".into(),
            right_to_left: false,
            tool_tip_text: "".into(),
            top: 840,
            visible: true,
            whats_this_help_id: 0,
            width: 1140,
        }
    }
}

impl Serialize for DataProperties {
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

impl<'a> From<Properties<'a>> for DataProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut data_prop = DataProperties::default();

        data_prop.align = prop.get_property(b"Align".into(), data_prop.align);
        data_prop.appearance = prop.get_property(b"Appearance".into(), data_prop.appearance);
        data_prop.back_color = prop.get_color(b"BackColor".into(), data_prop.back_color);
        data_prop.bof_action = prop.get_property(b"BOFAction".into(), data_prop.bof_action);
        data_prop.caption = match prop.get(b"Caption".into()) {
            Some(caption) => caption.into(),
            None => data_prop.caption,
        };
        data_prop.connection = prop
            .get(b"Connection".into())
            .map_or(Ok(Connection::Access), |v| {
                Connection::try_from(v.to_str().unwrap_or("Access"))
            })
            .unwrap();
        data_prop.database_name = match prop.get("DatabaseName".into()) {
            Some(database_name) => database_name.into(),
            None => "".into(),
        };
        data_prop.default_cursor_type =
            prop.get_property(b"DefaultCursorType".into(), data_prop.default_cursor_type);
        data_prop.default_type = prop.get_property(b"DefaultType".into(), data_prop.default_type);

        // DragIcon

        data_prop.drag_mode = prop.get_property(b"DragMode".into(), data_prop.drag_mode);
        data_prop.enabled = prop.get_bool(b"Enabled".into(), data_prop.enabled);
        data_prop.eof_action = prop.get_property(b"EOFAction".into(), data_prop.eof_action);
        data_prop.exclusive = prop.get_bool(b"Exclusive".into(), data_prop.exclusive);
        data_prop.fore_color = prop.get_color(b"ForeColor".into(), data_prop.fore_color);
        data_prop.height = prop.get_i32(b"Height".into(), data_prop.height);
        data_prop.left = prop.get_i32(b"Left".into(), data_prop.left);
        data_prop.mouse_pointer =
            prop.get_property(b"MousePointer".into(), data_prop.mouse_pointer);
        data_prop.negotitate = prop.get_bool(b"Negotitate".into(), data_prop.negotitate);
        data_prop.ole_drop_mode = prop.get_property(b"OLEDropMode".into(), data_prop.ole_drop_mode);
        data_prop.options = prop.get_i32(b"Options".into(), data_prop.options);
        data_prop.read_only = prop.get_bool(b"ReadOnly".into(), data_prop.read_only);
        data_prop.record_set_type =
            prop.get_property(b"RecordsetType".into(), data_prop.record_set_type);
        data_prop.record_source = match prop.get(b"RecordSource".into()) {
            Some(record_source) => record_source.into(),
            None => "".into(),
        };

        data_prop.right_to_left = prop.get_bool(b"RightToLeft".into(), data_prop.right_to_left);
        data_prop.tool_tip_text = match prop.get(b"ToolTipText".into()) {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => "".into(),
        };
        data_prop.top = prop.get_i32(b"Top".into(), data_prop.top);
        data_prop.visible = prop.get_bool(b"Visible".into(), data_prop.visible);
        data_prop.whats_this_help_id =
            prop.get_i32(b"WhatsThisHelpID".into(), data_prop.whats_this_help_id);
        data_prop.width = prop.get_i32(b"Width".into(), data_prop.width);

        data_prop
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
