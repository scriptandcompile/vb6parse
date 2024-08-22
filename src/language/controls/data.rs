use crate::language::controls::{Align, Appearance, DragMode, MousePointer, OLEDropMode};
use crate::VB6Color;

#[derive(Debug, PartialEq, Eq, Clone)]
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
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub eof_action: EOFAction,
    pub exclusive: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub left: i32,
    //pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub options: i32,
    pub read_only: bool,
    pub record_set_type: RecordSetType,
    // pub record_source: &'a str,
    pub right_to_left: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum BOFAction {
    MoveFirst = 0,
    BOF = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
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

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DefaultCursorType {
    DefaultCursor = 0,
    ODBCCursor = 1,
    ServerSideCursor = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DefaultType {
    UseODBC = 1,
    UseJet = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum EOFAction {
    MoveLast = 0,
    EOF = 1,
    AddNew = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum RecordSetType {
    Table = 0,
    Dynaset = 1,
    Snapshot = 2,
}
