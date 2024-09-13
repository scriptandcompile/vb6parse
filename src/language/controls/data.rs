use crate::language::controls::{Align, Appearance, DragMode, MousePointer, OLEDropMode};
use crate::VB6Color;

use image::DynamicImage;
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
    pub caption: &'a str,
    pub connection: Connection,
    pub database_name: &'a str,
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
    pub record_source: &'a str,
    pub right_to_left: bool,
    pub tool_tip_text: &'a str,
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
            caption: "Data1",
            connection: Connection::Access,
            database_name: "",
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
            record_source: "",
            right_to_left: false,
            tool_tip_text: "",
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

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum BOFAction {
    MoveFirst = 0,
    Bof = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum Connection {
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

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum DefaultCursorType {
    Default = 0,
    Odbc = 1,
    ServerSide = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum DefaultType {
    UseODBC = 1,
    UseJet = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum EOFAction {
    MoveLast = 0,
    Eof = 1,
    AddNew = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum RecordSetType {
    Table = 0,
    Dynaset = 1,
    Snapshot = 2,
}
