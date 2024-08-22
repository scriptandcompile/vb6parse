pub mod checkbox;
pub mod commandbutton;
pub mod listbox;
pub mod menus;
pub mod picturebox;

use crate::language::{
    controls::checkbox::CheckBoxProperties,
    controls::commandbutton::CommandButtonProperties,
    controls::listbox::ListBoxProperties,
    controls::menus::{MenuProperties, VB6MenuControl},
    controls::picturebox::PictureBoxProperties,
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
pub enum FormBorderStyle {
    None = 0,
    FixedSingle = 1,
    Sizable = 2,
    FixedDialog = 3,
    FixedToolWindow = 4,
    SizableToolWindow = 5,
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
pub enum ComboBoxStyle {
    DropDownCombo = 0,
    SimpleCombo = 1,
    DropDownList = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum BOFAction {
    MoveFirst = 0,
    BOF = 1,
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

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum LinkMode {
    None = 0,
    Automatic = 1,
    Manual = 2,
    Notify = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum FormLinkMode {
    None = 0,
    Source = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum ScrollBars {
    None = 0,
    Horizontal = 1,
    Vertical = 2,
    Both = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum MultiSelect {
    None = 0,
    Simple = 1,
    Extended = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum PaletteMode {
    HalfTone = 0,
    UseZOrder = 1,
    Custom = 2,
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
pub enum AutoActivate {
    Manual = 0,
    GetFocus = 1,
    DoubleClick = 2,
    Automatic = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DisplayType {
    Content = 0,
    Icon = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum OLETypeAllowed {
    Link = 0,
    Embedded = 1,
    Either = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum SizeMode {
    Clip = 0,
    Stretch = 1,
    AutoSize = 2,
    Zoom = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum UpdateOptions {
    Automatic = 0,
    Frozen = 1,
    Manual = 2,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum StartUpPosition {
    Manual = 0,
    CenterOwner = 1,
    CenterScreen = 2,
    WindowsDefault = 3,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum WindowState {
    Normal = 0,
    Minimized = 1,
    Maximized = 2,
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
pub struct TextBoxProperties<'a> {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub hide_selection: bool,
    pub left: i32,
    pub link_item: &'a str,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: &'a str,
    pub locked: bool,
    pub max_length: i32,
    //pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub multi_line: bool,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub password_char: Option<char>,
    pub right_to_left: bool,
    pub scroll_bars: ScrollBars,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub text: &'a str,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct LineProperties {
    pub border_color: VB6Color,
    pub border_style: DrawStyle,
    pub border_width: i32,
    pub draw_mode: DrawMode,
    pub visible: bool,
    pub x1: i32,
    pub y1: i32,
    pub x2: i32,
    pub y2: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct TimerProperties {
    pub enabled: bool,
    pub interval: i32,
    pub left: i32,
    pub top: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct LabelProperties<'a> {
    pub alignment: Alignment,
    pub appearance: Appearance,
    pub auto_size: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub caption: &'a str,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub left: i32,
    pub link_item: &'a str,
    pub link_mode: LinkMode,
    pub link_timeout: i32,
    pub link_topic: &'a str,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub tab_index: i32,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub use_mnemonic: bool,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
    pub word_wrap: bool,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct FrameProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub border_style: BorderStyle,
    pub caption: &'a str,
    pub clip_controls: bool,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    // pub mouse_icon: Option<ImageBuffer>,
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

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct FileListBoxProperties<'a> {
    pub appearance: Appearance,
    pub archive: bool,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub hidden: bool,
    pub left: i32,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub multi_select: MultiSelect,
    pub normal: bool,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub pattern: &'a str,
    pub read_only: bool,
    pub system: bool,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct DriveListBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_mode: OLEDropMode,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct DirListBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub left: i32,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct OLEProperties<'a> {
    pub appearance: Appearance,
    pub auto_activate: AutoActivate,
    pub auto_verb_menu: bool,
    pub back_color: VB6Color,
    pub back_style: BackStyle,
    pub border_style: BorderStyle,
    pub causes_validation: bool,
    pub class: Option<&'a str>,
    pub data_field: &'a str,
    pub data_source: &'a str,
    pub display_type: DisplayType,
    // pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub host_name: &'a str,
    pub left: i32,
    pub misc_flags: i32,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drop_allowed: bool,
    pub ole_type_allowed: OLETypeAllowed,
    pub size_mode: SizeMode,
    //pub source_doc: &'a str,
    //pub source_item: &'a str,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub top: i32,
    pub update_options: UpdateOptions,
    pub verb: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
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

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct ComboBoxProperties<'a> {
    pub appearance: Appearance,
    pub back_color: VB6Color,
    pub causes_validation: bool,
    pub data_field: &'a str,
    pub data_format: &'a str,
    pub data_member: &'a str,
    pub data_source: &'a str,
    // pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub fore_color: VB6Color,
    pub height: i32,
    pub help_context_id: i32,
    pub integral_height: bool,
    // pub item_data: Vec<&'a str>,
    pub left: i32,
    // pub list: Vec<&'a str>,
    pub locked: bool,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub ole_drag_mode: OLEDragMode,
    pub ole_drop_mode: OLEDropMode,
    pub right_to_left: bool,
    pub sorted: bool,
    pub style: ComboBoxStyle,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub text: &'a str,
    pub tool_tip_text: &'a str,
    pub top: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct ScrollBarProperties {
    pub causes_validation: bool,
    //pub drag_icon: Option<ImageBuffer>,
    pub drag_mode: DragMode,
    pub enabled: bool,
    pub height: i32,
    pub help_context_id: i32,
    pub large_change: i32,
    pub left: i32,
    pub max: i32,
    pub min: i32,
    //pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub right_to_left: bool,
    pub small_change: i32,
    pub tab_index: i32,
    pub tab_stop: bool,
    pub top: i32,
    pub value: i32,
    pub visible: bool,
    pub whats_this_help_id: i32,
    pub width: i32,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct FormProperties<'a> {
    pub appearance: Appearance,
    /// Determines if the output from a graphics method is to a persistent bitmap
    /// which acts as a double buffer.
    pub auto_redraw: bool,
    pub back_color: VB6Color,
    pub border_style: FormBorderStyle,
    pub caption: &'a str,
    pub clip_controls: bool,
    pub control_box: bool,
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
    // pub icon: Option<ImageBuffer>,
    pub key_preview: bool,
    pub left: i32,
    pub link_mode: FormLinkMode,
    pub link_topic: &'a str,
    pub max_button: bool,
    pub mdi_child: bool,
    pub min_button: bool,
    // pub mouse_icon: Option<ImageBuffer>,
    pub mouse_pointer: MousePointer,
    pub moveable: bool,
    pub negotiate_menus: bool,
    pub ole_drop_mode: OLEDropMode,
    // pub palette: Option<ImageBuffer>,
    pub pallette_mode: PaletteMode,
    // pub picture: Option<ImageBuffer>,
    pub right_to_left: bool,
    pub scale_height: i32,
    pub scale_left: i32,
    pub scale_mode: ScaleMode,
    pub scale_top: i32,
    pub scale_width: i32,
    pub show_in_taskbar: bool,
    pub start_up_position: StartUpPosition,
    pub top: i32,
    pub visible: bool,
    pub whats_this_button: bool,
    pub whats_this_help: bool,
    pub width: i32,
    pub window_state: WindowState,
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

impl Default for ComboBoxProperties<'_> {
    fn default() -> Self {
        ComboBoxProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H80000005&").unwrap(),
            causes_validation: true,
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000008&").unwrap(),
            height: 30,
            help_context_id: 0,
            integral_height: true,
            left: 30,
            locked: false,
            mouse_pointer: MousePointer::Default,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::None,
            right_to_left: false,
            sorted: false,
            style: ComboBoxStyle::DropDownCombo,
            tab_index: 0,
            tab_stop: true,
            text: "",
            tool_tip_text: "",
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Default for TextBoxProperties<'_> {
    fn default() -> Self {
        TextBoxProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H80000005&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            causes_validation: true,
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000008&").unwrap(),
            height: 30,
            help_context_id: 0,
            hide_selection: true,
            left: 30,
            link_item: "",
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "",
            locked: false,
            max_length: 0,
            mouse_pointer: MousePointer::Default,
            multi_line: false,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::None,
            password_char: None,
            right_to_left: false,
            scroll_bars: ScrollBars::None,
            tab_index: 0,
            tab_stop: true,
            text: "",
            tool_tip_text: "",
            top: 30,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Default for LineProperties {
    fn default() -> Self {
        LineProperties {
            border_color: VB6Color::from_hex("&H80000008&").unwrap(),
            border_style: DrawStyle::Solid,
            border_width: 1,
            draw_mode: DrawMode::CopyPen,
            visible: true,
            x1: 0,
            y1: 0,
            x2: 100,
            y2: 100,
        }
    }
}

impl Default for LabelProperties<'_> {
    fn default() -> Self {
        LabelProperties {
            alignment: Alignment::LeftJustify,
            appearance: Appearance::ThreeD,
            auto_size: false,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            back_style: BackStyle::Opaque,
            border_style: BorderStyle::None,
            caption: "Label1",
            data_field: "",
            data_format: "",
            data_member: "",
            data_source: "",
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            left: 30,
            link_item: "",
            link_mode: LinkMode::None,
            link_timeout: 50,
            link_topic: "",
            mouse_pointer: MousePointer::Default,
            ole_drop_mode: OLEDropMode::None,
            right_to_left: false,
            tab_index: 0,
            tool_tip_text: "",
            top: 30,
            use_mnemonic: true,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
            word_wrap: false,
        }
    }
}

impl Default for ScrollBarProperties {
    fn default() -> Self {
        ScrollBarProperties {
            causes_validation: true,
            drag_mode: DragMode::Manual,
            enabled: true,
            height: 30,
            help_context_id: 0,
            large_change: 1,
            left: 30,
            max: 32767,
            min: 0,
            mouse_pointer: MousePointer::Default,
            right_to_left: false,
            small_change: 1,
            tab_index: 0,
            tab_stop: true,
            top: 30,
            value: 0,
            visible: true,
            whats_this_help_id: 0,
            width: 100,
        }
    }
}

impl Default for TimerProperties {
    fn default() -> Self {
        TimerProperties {
            enabled: true,
            interval: 0,
            left: 0,
            top: 0,
        }
    }
}

impl Default for FrameProperties<'_> {
    fn default() -> Self {
        FrameProperties {
            appearance: Appearance::ThreeD,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: BorderStyle::FixedSingle,
            caption: "Frame1",
            clip_controls: true,
            drag_mode: DragMode::Manual,
            enabled: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            height: 30,
            help_context_id: 0,
            left: 30,
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

impl Default for FormProperties<'_> {
    fn default() -> Self {
        FormProperties {
            appearance: Appearance::ThreeD,
            auto_redraw: false,
            back_color: VB6Color::from_hex("&H8000000F&").unwrap(),
            border_style: FormBorderStyle::Sizable,
            caption: "Form1",
            clip_controls: true,
            control_box: true,
            draw_mode: DrawMode::CopyPen,
            draw_style: DrawStyle::Solid,
            draw_width: 1,
            enabled: true,
            fill_color: VB6Color::from_hex("&H00000000&").unwrap(),
            fill_style: FillStyle::Transparent,
            font_transparent: true,
            fore_color: VB6Color::from_hex("&H80000012&").unwrap(),
            has_dc: true,
            height: 240,
            help_context_id: 0,
            key_preview: false,
            left: 0,
            link_mode: FormLinkMode::None,
            link_topic: "Form1",
            max_button: true,
            mdi_child: false,
            min_button: true,
            mouse_pointer: MousePointer::Default,
            moveable: true,
            negotiate_menus: true,
            ole_drop_mode: OLEDropMode::None,
            pallette_mode: PaletteMode::HalfTone,
            right_to_left: false,
            scale_height: 240,
            scale_left: 0,
            scale_mode: ScaleMode::Twip,
            scale_top: 0,
            scale_width: 240,
            show_in_taskbar: true,
            start_up_position: StartUpPosition::WindowsDefault,
            top: 0,
            visible: true,
            whats_this_button: false,
            whats_this_help: false,
            width: 240,
            window_state: WindowState::Normal,
        }
    }
}
