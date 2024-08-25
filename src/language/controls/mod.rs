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
pub mod menus;
pub mod ole;
pub mod optionbutton;
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
    controls::image::ImageProperties,
    controls::label::LabelProperties,
    controls::line::LineProperties,
    controls::listbox::ListBoxProperties,
    controls::menus::{MenuProperties, VB6MenuControl},
    controls::ole::OLEProperties,
    controls::optionbutton::OptionButtonProperties,
    controls::picturebox::PictureBoxProperties,
    controls::scrollbars::ScrollBarProperties,
    controls::textbox::TextBoxProperties,
    controls::timer::TimerProperties,
};

/// Represents a VB6 control.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Control<'a> {
    pub name: &'a str,
    pub tag: &'a str,
    pub index: i32,
    pub kind: VB6ControlKind<'a>,
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
}

impl<'a> VB6ControlKind<'a> {
    pub fn is_menu(&self) -> bool {
        match self {
            VB6ControlKind::Menu { .. } => true,
            _ => false,
        }
    }
}

/// Determines which side of the parent control to dock this control to.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Align {
    /// The control is not docked to any side of the parent control.
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

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum JustifyAlignment {
    LeftJustify = 0,
    RightJustify = 1,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Alignment {
    LeftJustify = 0,
    RightJustify = 1,
    Center = 2,
}

/// The back style determines whether the background of a control is opaque or transparent.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum BackStyle {
    /// The background of the control is transparent.
    Transparent = 0,
    /// The background of the control is opaque. (default)
    Opaque = 1,
}

/// The appearance determines whether or not a control is painted at run time
/// with 3D effects.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Appearance {
    /// The control is painted with a flat style.
    Flat = 0,
    /// The control is painted with a 3D style.
    ThreeD = 1,
}

/// The border style determines the appearance of the border of a control.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum BorderStyle {
    /// The control has no border.
    None = 0,
    /// The control has a single-line border.
    FixedSingle = 1,
}

/// Determines the style of drag and drop operations.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DragMode {
    /// The control does not support drag and drop operations until
    /// the program manually initiates the drag operation.
    Manual = 0,
    /// The control automatically initiates a drag operation when the
    /// user presses the mouse button on the control.
    Automatic = 1,
}

/// Specifies how the pen (the color used in drawing) interacts with the
/// background.
#[derive(Debug, PartialEq, Eq, Clone)]
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
    /// The color specified by the ForeColor property is applied over the background.
    /// This is the default setting.
    CopyPen = 13,
    /// The combination of the pen color and inverse of the display color.
    MergePenNot = 14,
    /// the combination of the pen color and the display color.
    MergePen = 15,
    /// White pen color is applied over the background.
    Whiteness = 16,
}

/// Determines the line style of any drawing from any graphic method applied by the control.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum DrawStyle {
    /// A solid line.
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
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum MousePointer {
    /// Standard pointer. the image is determined by the object (default).
    Default = 0,
    /// Arrow pointer.
    Arrow = 1,
    /// Cross-hair pointer.
    Cross = 2,
    /// I-beam pointer.
    IBeam = 3,
    /// Icon pointer. The image is determined by the MouseIcon property.
    /// If the MouseIcon property is not set, the behavior is the same as the Default setting.
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
    /// Uses the icon specified by the MouseIcon property.
    /// If the MouseIcon property is not set, the behavior is the same as the Default setting.
    /// This is a duplicate of Icon (4).
    Custom = 99,
}

/// Determines the style of drag and drop operations.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum OLEDragMode {
    /// The programmer handles all OLE drag/drop events manually. (default).
    Manual = 0,
    /// The control automatically handles all OLE drag/drop events.
    Automatic = 1,
}

/// Determines the style of drop operations.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum OLEDropMode {
    /// The control does not accept any OLE drop operations.
    None = 0,
    /// The programmer handles all OLE drop events manually.
    Manual = 1,
}

/// Determines if the control uses standard styling or if it uses graphical styling from it's
/// picture properties.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Style {
    /// The control uses standard styling.
    Standard = 0,
    /// The control uses graphical styling using its appropriate picture properties.
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
