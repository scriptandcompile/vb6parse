pub mod checkbox;
pub mod combobox;
pub mod commandbutton;
pub mod data;
pub mod dirlistbox;
pub mod drivelistbox;
pub mod filelistbox;
pub mod form;
pub mod frame;
pub mod label;
pub mod line;
pub mod listbox;
pub mod menus;
pub mod ole;
pub mod picturebox;
pub mod scrollbars;
pub mod textbox;
pub mod timer;

use crate::language::{
    controls::checkbox::CheckBoxProperties,
    controls::combobox::ComboBoxProperties,
    controls::commandbutton::CommandButtonProperties,
    controls::data::DataProperties,
    controls::dirlistbox::DirListBoxProperties,
    controls::drivelistbox::DriveListBoxProperties,
    controls::filelistbox::FileListBoxProperties,
    controls::form::FormProperties,
    controls::frame::FrameProperties,
    controls::label::LabelProperties,
    controls::line::LineProperties,
    controls::listbox::ListBoxProperties,
    controls::menus::{MenuProperties, VB6MenuControl},
    controls::ole::OLEProperties,
    controls::picturebox::PictureBoxProperties,
    controls::scrollbars::ScrollBarProperties,
    controls::textbox::TextBoxProperties,
    controls::timer::TimerProperties,
    VB6Color,
};

/// Represents a VB6 control.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Control<'a> {
    pub name: &'a str,
    pub tag: &'a str,
    pub index: i32,
    pub kind: VB6ControlKind<'a>,
}

///
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Align {
    None = 0,
    Top = 1,
    Bottom = 2,
    Left = 3,
    Right = 4,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Alignment {
    LeftJustify = 0,
    RightJustify = 1,
    Center = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]

pub enum BackStyle {
    Transparent = 0,
    Opaque = 1,
}

/// Whether or not a control is painted at run time with 3D effects.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Appearance {
    /// The control is painted with a flat style.
    Flat = 0,
    /// The control is painted with a 3D style.
    ThreeD = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum BorderStyle {
    None = 0,
    FixedSingle = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DragMode {
    Manual = 0,
    Automatic = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DrawMode {
    Blackness = 1,
    NotMergePen = 2,
    MaskNotPen = 3,
    NotCopyPen = 4,
    MaskPenNot = 5,
    Invert = 6,
    XorPen = 7,
    NotMaskPen = 8,
    MaskPen = 9,
    NotXorPen = 10,
    Nop = 11,
    MergeNotPen = 12,
    CopyPen = 13,
    MergePenNot = 14,
    MergePen = 15,
    Whiteness = 16,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DrawStyle {
    Solid = 0,
    Dash = 1,
    Dot = 2,
    DashDot = 3,
    DashDotDot = 4,
    Transparent = 5,
    InsideSolid = 6,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum MousePointer {
    Default = 0,
    Arrow = 1,
    Cross = 2,
    IBeam = 3,
    Icon = 4,
    Size = 5,
    SizeNESW = 6,
    SizeNS = 7,
    SizeNWSE = 8,
    SizeWE = 9,
    UpArrow = 10,
    Hourglass = 11,
    NoDrop = 12,
    ArrowHourglass = 13,
    ArrowQuestion = 14,
    SizeAll = 15,
    Custom = 99,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum OLEDragMode {
    Manual = 0,
    Automatic = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum OLEDropMode {
    None = 0,
    Manual = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Style {
    Standard = 0,
    Graphical = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum FillStyle {
    Solid = 0,
    Transparent = 1,
    HorizontalLine = 2,
    VerticalLine = 3,
    UpwardDiagonal = 4,
    DownwardDiagonal = 5,
    Cross = 6,
    DiagonalCross = 7,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum LinkMode {
    None = 0,
    Automatic = 1,
    Manual = 2,
    Notify = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum MultiSelect {
    None = 0,
    Simple = 1,
    Extended = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum ScaleMode {
    User = 0,
    Twip = 1,
    Point = 2,
    Pixel = 3,
    Character = 4,
    Inches = 5,
    Millimeter = 6,
    Centimeter = 7,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum SizeMode {
    Clip = 0,
    Stretch = 1,
    AutoSize = 2,
    Zoom = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum NegotiatePosition {
    None = 0,
    Left = 1,
    Middle = 2,
    Right = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum ShortCut {
    CtrlA,
    CtrlB,
    CtrlC,
    CtrlD,
    CtrlE,
    CtrlF,
    CtrlG,
    CtrlH,
    CtrlI,
    CtrlJ,
    CtrlK,
    CtrlL,
    CtrlM,
    CtrlN,
    CtrlO,
    CtrlP,
    CtrlQ,
    CtrlR,
    CtrlS,
    CtrlT,
    CtrlU,
    CtrlV,
    CtrlW,
    CtrlX,
    CtrlY,
    CtrlZ,
    F1,
    F2,
    F3,
    F4,
    F5,
    F6,
    F7,
    F8,
    F9,
    // F10 is not included.
    F11,
    F12,
    CtrlF1,
    CtrlF2,
    CtrlF3,
    CtrlF4,
    CtrlF5,
    CtrlF6,
    CtrlF7,
    CtrlF8,
    CtrlF9,
    // CtrlF10 is not included.
    CtrlF11,
    CtrlF12,
    ShiftF1,
    ShiftF2,
    ShiftF3,
    ShiftF4,
    ShiftF5,
    ShiftF6,
    ShiftF7,
    ShiftF8,
    ShiftF9,
    // ShiftF10 is not included.
    ShiftF11,
    ShiftF12,
    ShiftCtrlF1,
    ShiftCtrlF2,
    ShiftCtrlF3,
    ShiftCtrlF4,
    ShiftCtrlF5,
    ShiftCtrlF6,
    ShiftCtrlF7,
    ShiftCtrlF8,
    ShiftCtrlF9,
    // ShiftCtrlF10 is not included.
    ShiftCtrlF11,
    ShiftCtrlF12,
    CtrlIns,
    ShiftIns,
    Del,
    ShiftDel,
    AltBKsp,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct ImageProperties<'a> {
    pub appearance: Appearance,
    pub border_style: BorderStyle,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    // pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub left: i32,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    // pub picture: Option<ImageBuffer>,
    pub stretch: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

/// Represents a VB6 control kind.
/// A VB6 control kind is an enumeration of the different kinds of
/// standard VB6 controls.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum VB6ControlKind<'a> {
    CommandButton {
        properties: CommandButtonProperties<'a>,
    },
    Data {
        properties: DataProperties<'a>,
    },
    TextBox {
        properties: TextBoxProperties<'a>,
    },
    CheckBox {
        properties: CheckBoxProperties<'a>,
    },
    Line {
        properties: LineProperties,
    },
    ListBox {
        properties: ListBoxProperties<'a>,
    },
    Timer {
        properties: TimerProperties,
    },
    Label {
        properties: LabelProperties<'a>,
    },
    Frame {
        properties: FrameProperties<'a>,
        controls: Vec<VB6Control<'a>>,
    },
    PictureBox {
        properties: PictureBoxProperties<'a>,
    },
    FileListBox {
        properties: FileListBoxProperties<'a>,
    },
    DriveListBox {
        properties: DriveListBoxProperties<'a>,
    },
    DirListBox {
        properties: DirListBoxProperties<'a>,
    },
    Ole {
        properties: OLEProperties<'a>,
    },
    Image {
        properties: ImageProperties<'a>,
    },
    ComboBox {
        properties: ComboBoxProperties<'a>,
    },
    HScrollBar {
        properties: ScrollBarProperties,
    },
    VScrollBar {
        properties: ScrollBarProperties,
    },
    Menu {
        properties: MenuProperties<'a>,
        sub_menus: Vec<VB6MenuControl<'a>>,
    },
    Form {
        properties: FormProperties<'a>,
        controls: Vec<VB6Control<'a>>,
        menus: Vec<VB6MenuControl<'a>>,
    },
}

impl<'a> VB6ControlKind<'a> {
    pub fn is_menu(&self) -> bool {
        match self {
            VB6ControlKind::Menu { .. } => true,
            _ => false,
        }
    }
}
