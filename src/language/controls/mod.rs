pub mod checkbox;
pub mod combobox;
pub mod commandbutton;
pub mod data;
pub mod dirlistbox;
pub mod drivelistbox;
pub mod filelistbox;
pub mod form;
pub mod frame;
pub mod image;
pub mod label;
pub mod line;
pub mod listbox;
pub mod mdiform;
pub mod menus;
pub mod ole;
pub mod optionbutton;
pub mod picturebox;
pub mod scrollbars;
pub mod shape;
pub mod textbox;
pub mod timer;

use std::collections::HashMap;

use bstr::BStr;
use num_enum::TryFromPrimitive;
use serde::Serialize;

use crate::parsers::form::VB6PropertyGroup;

use crate::language::controls::{
    checkbox::CheckBoxProperties,
    combobox::ComboBoxProperties,
    commandbutton::CommandButtonProperties,
    data::DataProperties,
    dirlistbox::DirListBoxProperties,
    drivelistbox::DriveListBoxProperties,
    filelistbox::FileListBoxProperties,
    form::FormProperties,
    frame::FrameProperties,
    image::ImageProperties,
    label::LabelProperties,
    line::LineProperties,
    listbox::ListBoxProperties,
    mdiform::MDIFormProperties,
    menus::{MenuProperties, VB6MenuControl},
    ole::OLEProperties,
    optionbutton::OptionButtonProperties,
    picturebox::PictureBoxProperties,
    scrollbars::ScrollBarProperties,
    shape::ShapeProperties,
    textbox::TextBoxProperties,
    timer::TimerProperties,
};

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum FormLinkMode {
    #[default]
    None = 0,
    Source = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum WindowState {
    #[default]
    Normal = 0,
    Minimized = 1,
    Maximized = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default)]
pub enum StartUpPosition {
    /// 0
    Manual {
        client_height: i32,
        client_width: i32,
        client_top: i32,
        client_left: i32,
    },
    /// 1
    CenterOwner,
    /// 2
    CenterScreen,
    #[default]
    /// 3
    WindowsDefault,
}

/// Represents a VB6 control.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct VB6Control<'a> {
    pub name: &'a BStr,
    pub tag: &'a BStr,
    pub index: i32,
    pub kind: VB6ControlKind<'a>,
}

/// The `VB6ControlKind` determines the specific kind of control that the `VB6Control` represents.
///
/// Each variant contains the properties that are specific to that kind of control.
#[derive(Debug, PartialEq, Clone, Serialize)]
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
    Shape {
        properties: ShapeProperties,
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
    OptionButton {
        properties: OptionButtonProperties<'a>,
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
    MDIForm {
        properties: MDIFormProperties<'a>,
        controls: Vec<VB6Control<'a>>,
        menus: Vec<VB6MenuControl<'a>>,
    },
    Custom {
        properties: HashMap<&'a BStr, &'a BStr>,
        property_groups: Vec<VB6PropertyGroup<'a>>,
    },
}

impl<'a> VB6ControlKind<'a> {
    #[must_use]
    pub fn is_menu(&self) -> bool {
        matches!(self, VB6ControlKind::Menu { .. })
    }
}

/// Determines which side of the parent control to dock this control to.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum Align {
    /// The control is not docked to any side of the parent control.
    /// This is the default setting.
    #[default]
    None = 0,
    /// The control is docked to the top of the parent control.
    Top = 1,
    /// The control is docked to the bottom of the parent control.
    Bottom = 2,
    /// The control is docked to the left of the parent control.
    Left = 3,
    /// The control is docked to the right of the parent control.
    Right = 4,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum JustifyAlignment {
    #[default]
    LeftJustify = 0,
    RightJustify = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum Alignment {
    #[default]
    LeftJustify = 0,
    RightJustify = 1,
    Center = 2,
}

/// The back style determines whether the background of a control is opaque or transparent.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum BackStyle {
    /// The background of the control is transparent.
    Transparent = 0,
    /// The background of the control is opaque. (default)
    #[default]
    Opaque = 1,
}

/// The appearance determines whether or not a control is painted at run time
/// with 3D effects.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum Appearance {
    /// The control is painted with a flat style.
    Flat = 0,
    /// The control is painted with a 3D style.
    #[default]
    ThreeD = 1,
}

/// The border style determines the appearance of the border of a control.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum BorderStyle {
    /// The control has no border.
    None = 0,
    /// The control has a single-line border.
    #[default]
    FixedSingle = 1,
}

/// Determines the style of drag and drop operations.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum DragMode {
    /// The control does not support drag and drop operations until
    /// the program manually initiates the drag operation.
    #[default]
    Manual = 0,
    /// The control automatically initiates a drag operation when the
    /// user presses the mouse button on the control.
    Automatic = 1,
}

/// Specifies how the pen (the color used in drawing) interacts with the
/// background.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum DrawMode {
    /// Black pen color is applied over the background.
    Blackness = 1,
    /// Inversion is applied after the combination of the pen and the background color.
    NotMergePen = 2,
    /// The combination of the colors common to the background color and the inverse of the pen.
    MaskNotPen = 3,
    /// Inversion is applied to the pen color.
    NotCopyPen = 4,
    /// The combination of the colors common to the pen and the inverse of the background color.
    MaskPenNot = 5,
    /// Inversion is applied to the background color.
    Invert = 6,
    /// The combination of the colors common to the pen and the background color, but not in both (ie, XOR).
    XorPen = 7,
    /// Inversion is applied to the combination of the colors common to both the pen and the background color.
    NotMaskPen = 8,
    /// The combination of the colors common to the pen and the background color.
    MaskPen = 9,
    /// Inversion of the combinationfs of the colors in the pen and the background color but not in both (ie, NXOR).
    NotXorPen = 10,
    /// No operation is performed. The output remains unchanged. In effect, this turns drawing off (No Operation).
    Nop = 11,
    /// The combinaton of the display color and the inverse of the pen color.
    MergeNotPen = 12,
    /// The color specified by the `ForeColor` property is applied over the background.
    /// This is the default setting.
    #[default]
    CopyPen = 13,
    /// The combination of the pen color and inverse of the display color.
    MergePenNot = 14,
    /// the combination of the pen color and the display color.
    MergePen = 15,
    /// White pen color is applied over the background.
    Whiteness = 16,
}

/// Determines the line style of any drawing from any graphic method applied by the control.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum DrawStyle {
    /// A solid line. This is the default.
    #[default]
    Solid = 0,
    /// A dashed line.
    Dash = 1,
    /// A dotted line.
    Dot = 2,
    /// A line that alternates between dashes and dots.
    DashDot = 3,
    /// A line that alternates between dashes and double dots.
    DashDotDot = 4,
    /// Invisible line, transparent interior.
    Transparent = 5,
    /// Invisible line, solid interior.
    InsideSolid = 6,
}

/// Determines the appearance of the mouse pointer when it is over the control.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum MousePointer {
    /// Standard pointer. the image is determined by the object (default).
    #[default]
    Default = 0,
    /// Arrow pointer.
    Arrow = 1,
    /// Cross-hair pointer.
    Cross = 2,
    /// I-beam pointer.
    IBeam = 3,
    /// Icon pointer. The image is determined by the `MouseIcon` property.
    /// If the `MouseIcon` property is not set, the behavior is the same as the Default setting.
    /// This is a duplicate of Custom (99).
    Icon = 4,
    /// Size all cursor (arrows pointing north, south, east, and west).
    /// This cursor is used to indicate that the control can be resized in any direction.
    Size = 5,
    /// Double arrow pointing northeast and southwest.
    SizeNESW = 6,
    /// Double arrow pointing north and south.
    SizeNS = 7,
    /// Double arrow pointing northwest and southeast.
    SizeNWSE = 8,
    /// Double arrow pointing west and east.
    SizeWE = 9,
    /// Up arrow.
    UpArrow = 10,
    /// Hourglass or wait cursor.
    Hourglass = 11,
    /// "Not" symbol (circle with a diagonal line) on top of the object being dragged.
    /// Indicates an invalid drop target.
    NoDrop = 12,
    // Arrow with an hourglass.
    ArrowHourglass = 13,
    /// Arrow with a question mark.
    ArrowQuestion = 14,
    /// Size all cursor (arrows pointing north, south, east, and west).
    /// This cursor is used to indicate that the control can be resized in any direction.
    /// Duplicate of Size (5).
    SizeAll = 15,
    /// Uses the icon specified by the `MouseIcon` property.
    /// If the `MouseIcon` property is not set, the behavior is the same as the Default setting.
    /// This is a duplicate of Icon (4).
    Custom = 99,
}

/// Determines the style of drag and drop operations.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum OLEDragMode {
    /// The programmer handles all OLE drag/drop events manually. (default).
    #[default]
    Manual = 0,
    /// The control automatically handles all OLE drag/drop events.
    Automatic = 1,
}

/// Determines the style of drop operations.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum OLEDropMode {
    /// The control does not accept any OLE drop operations.
    #[default]
    None = 0,
    /// The programmer handles all OLE drop events manually.
    Manual = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum ClipControls {
    /// The controls are not clipped to the bounds of the parent control.
    False = 0,
    /// The controls are clipped to the bounds of the parent control.
    #[default]
    True = 1,
}

/// Determines if the control uses standard styling or if it uses graphical styling from it's
/// picture properties.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum Style {
    /// The control uses standard styling.
    #[default]
    Standard = 0,
    /// The control uses graphical styling using its appropriate picture properties.
    Graphical = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum FillStyle {
    Solid = 0,
    #[default]
    Transparent = 1,
    HorizontalLine = 2,
    VerticalLine = 3,
    UpwardDiagonal = 4,
    DownwardDiagonal = 5,
    Cross = 6,
    DiagonalCross = 7,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum LinkMode {
    #[default]
    None = 0,
    Automatic = 1,
    Manual = 2,
    Notify = 3,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum MultiSelect {
    #[default]
    None = 0,
    Simple = 1,
    Extended = 2,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum ScaleMode {
    User = 0,
    #[default]
    Twip = 1,
    Point = 2,
    Pixel = 3,
    Character = 4,
    Inches = 5,
    Millimeter = 6,
    Centimeter = 7,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum SizeMode {
    #[default]
    Clip = 0,
    Stretch = 1,
    AutoSize = 2,
    Zoom = 3,
}
