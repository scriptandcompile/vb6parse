pub mod checkbox;
pub mod combobox;
pub mod commandbutton;
pub mod custom;
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

use bstr::BString;
use num_enum::TryFromPrimitive;
use serde::Serialize;

use crate::parsers::form::VB6PropertyGroup;

use crate::language::controls::{
    checkbox::CheckBoxProperties,
    combobox::ComboBoxProperties,
    commandbutton::CommandButtonProperties,
    custom::CustomControlProperties,
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

/// `AutoRedraw` determines if the control is redrawn automatically when something is
/// moved in front of it or if it is redrawn manually.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245029(v=vs.60))
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum AutoRedraw {
    /// Disables automatic repainting of an object and writes graphics or text
    /// only to the screen. Visual Basic invokes the object's `Paint` event when
    /// necessary to repaint the object.
    ///
    /// This is the default setting.
    #[default]
    Manual = 0,
    /// Enables automatic repainting of a `Form` object or `PictureBox` control.
    /// Graphics and text are written to the screen and to an image stored in memory.
    /// The object doesn't receive `Paint` events; it's repainted when necessary,
    /// using the image stored in memory.
    Automatic = -1,
}

/// `TextDirection` determines the direction in which text is displayed in the control.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa442921(v=vs.60))
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum TextDirection {
    /// The text is ordered from left to right.
    ///
    /// This is the default setting.
    #[default]
    LeftToRight = 0,
    /// The text is ordered from right to left.
    RightToLeft = -1,
}

/// `AutoSize` determines if the control is automatically resized to fit its contents.
/// This is used with the `Label` control and the `PictureBox` control.
///
/// In a `PictureBox`, this property is used to determine if the control is automatically resized
/// to fit the size of the picture. If set to `Fixed` the control is not resized and the picture
/// will be scaled or clipped depending on other properties like `SizeMode`.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245034(v=vs.60))
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum AutoSize {
    /// Keeps the size of the control constant. Contents are clipped when they
    /// exceed the area of the control.
    ///
    /// This is the default setting.
    #[default]
    Fixed = 0,
    /// Automatically resizes the control to display its entire contents.
    Resize = -1,
}

/// Determines if a control or form can respond to user-generated events.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267301(v=vs.60))
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum Activation {
    /// The control is disabled and will not respond to user-generated events.
    Disabled = 0,
    /// The control is enabled and will respond to user-generated events.
    ///
    /// This is the default setting.
    #[default]
    Enabled = -1,
}

/// `TabStop` determines if the control is included in the tab order.
/// In VB6, the `TabStop` property determines whether a control can receive focus
/// when the user navigates through controls using the Tab key.
///
/// When `TabStop` is set to `Included`, the control is included in the tab order
/// and can receive focus when the user presses the Tab key.
///
/// When `TabStop` is set to `ProgrammaticOnly`, the control is skipped in the
/// tab order and cannot receive focus via the Tab key.
/// However, it can still receive focus programmatically or through other user interactions.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445721(v=vs.60))
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum TabStop {
    /// Bypasses the object when the user is tabbing, although the object still
    /// holds its place in the actual tab order, as determined by the `TabIndex`
    /// property.
    ProgrammaticOnly = 0,
    /// Designates the object as a tab stop.
    ///
    /// This is the default setting.
    #[default]
    Included = -1,
}

/// Determines if the control is visible or hidden.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445768(v=vs.60))
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum Visibility {
    /// The control is not visible.
    Hidden = 0,
    /// The control is visible.
    ///
    /// This is the default setting.
    #[default]
    Visible = -1,
}

/// Determines if the control has a device context.
/// A device context is a Windows data structure that defines a set of graphic objects
/// and their associated attributes, and it defines a mapping between the logical
/// coordinates and device coordinates for a particular device, such as a display or printer.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum HasDeviceContext {
    /// The control does not have a device context.
    No = 0,
    /// The control has a device context.
    #[default]
    Yes = -1,
}

/// Determines if the control uses the `mask_color` property as the trnsparent color
/// on the control.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum UseMaskColor {
    /// The control does not use the mask color.
    DoNotUseMaskColor = 0,
    /// The control uses the mask color.
    #[default]
    UseMaskColor = -1,
}

/// Determines if the control causes validation.
/// In VB6, the `CausesValidation` property determines whether a control causes validation
/// to occur when the user attempts to move focus from the control.
/// If `CausesValidation` is set to `True`, validation occurs when the user attempts to move
/// focus from the control to another control.
/// If `CausesValidation` is set to `False`, validation does not occur when the user attempts
/// to move focus from the control to another control.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum CausesValidation {
    /// The control does not cause validation.
    No = 0,
    /// The control causes validation.
    #[default]
    Yes = -1,
}

/// The `Movability` property of a `Form` control determines whether the
/// form can be moved by the user. If the form is not moveable, the user cannot
/// move the form by dragging its title bar or by using the arrow keys.
/// If the form is moveable, the user can move the form by dragging its title
/// bar or by using the arrow keys.
#[derive(Debug, PartialEq, Eq, Clone, Default, TryFromPrimitive, serde::Serialize)]
#[repr(i32)]
pub enum Movability {
    /// The form is not moveable.
    Fixed = 0,
    /// The form is moveable.
    #[default]
    Moveable = -1,
}

/// The `FontTransparency` property of a `Form` or `PictureBox` control determines
/// whether the `Font` property is transparent or opaque.
#[derive(Debug, PartialEq, Eq, Clone, Default, TryFromPrimitive, serde::Serialize)]
#[repr(i32)]
pub enum FontTransparency {
    /// The font is not transparent.
    Opaque = 0,
    /// The font is transparent.
    #[default]
    Transparent = -1,
}

/// The `WhatsThisHelp` property of a `Form` control determines whether the
/// context-sensitive Help uses the pop-up window provided by Windows 95 Help
/// or the main Help window.
#[derive(Debug, PartialEq, Eq, Clone, Default, TryFromPrimitive, serde::Serialize)]
#[repr(i32)]
pub enum WhatsThisHelp {
    /// The application uses the F1 key to start Windows Help and load the topic
    /// identified by the `help_context_id` property.
    #[default]
    F1Help = 0,
    // The application uses one of the 'What's This' access techniques to start
    // Windows Help.
    WhatsThisHelp = -1,
}

#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum FormLinkMode {
    #[default]
    None = 0,
    Source = 1,
}

/// Controls the display state of a form from normal, minimized, or maximized.
/// This is used with the `Form` and `MDIForm` controls.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum WindowState {
    /// The form is in its normal state.
    #[default]
    Normal = 0,
    /// The form is minimized.
    Minimized = 1,
    /// The form is maximized.
    Maximized = 2,
}

/// The `StartUpPosition` property of a `Form` or `MDIForm` control determines
/// the initial position of the form when it is displayed.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default)]
pub enum StartUpPosition {
    /// The form is positioned based on the `client_height`, `client_width`,
    /// `client_top`, and `client_left` properties.
    ///
    /// The `Manual` variant is saved as a 0 in the VB6 file.
    Manual {
        /// The height of the client area of the form.
        client_height: i32,
        /// The width of the client area of the form.
        client_width: i32,
        /// The top position of the client area of the form.
        client_top: i32,
        /// The left position of the client area of the form.
        client_left: i32,
    },
    /// The form is centered in the parent window.
    ///
    /// The `CenterOwner` variant is saved as a 1 in the VB6 file.
    CenterOwner,
    /// The form is centered on the screen.
    ///
    /// The `CenterScreen` variant is saved as a 2 in the VB6 file.
    CenterScreen,
    #[default]
    /// The form requests the operating system to position the form.
    ///
    /// The `WindowsDefault` variant is saved as a 3 in the VB6 file.
    WindowsDefault,
}

/// Represents a VB6 control.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct VB6Control {
    pub name: BString,
    pub tag: BString,
    pub index: i32,
    pub kind: VB6ControlKind,
}

/// The `VB6ControlKind` determines the specific kind of control that the `VB6Control` represents.
///
/// Each variant contains the properties that are specific to that kind of control.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub enum VB6ControlKind {
    CommandButton {
        properties: CommandButtonProperties,
    },
    Data {
        properties: DataProperties,
    },
    TextBox {
        properties: TextBoxProperties,
    },
    CheckBox {
        properties: CheckBoxProperties,
    },
    Line {
        properties: LineProperties,
    },
    Shape {
        properties: ShapeProperties,
    },
    ListBox {
        properties: ListBoxProperties,
    },
    Timer {
        properties: TimerProperties,
    },
    Label {
        properties: LabelProperties,
    },
    Frame {
        properties: FrameProperties,
        controls: Vec<VB6Control>,
    },
    PictureBox {
        properties: PictureBoxProperties,
    },
    FileListBox {
        properties: FileListBoxProperties,
    },
    DriveListBox {
        properties: DriveListBoxProperties,
    },
    DirListBox {
        properties: DirListBoxProperties,
    },
    Ole {
        properties: OLEProperties,
    },
    OptionButton {
        properties: OptionButtonProperties,
    },
    Image {
        properties: ImageProperties,
    },
    ComboBox {
        properties: ComboBoxProperties,
    },
    HScrollBar {
        properties: ScrollBarProperties,
    },
    VScrollBar {
        properties: ScrollBarProperties,
    },
    Menu {
        properties: MenuProperties,
        sub_menus: Vec<VB6MenuControl>,
    },
    Form {
        properties: FormProperties,
        controls: Vec<VB6Control>,
        menus: Vec<VB6MenuControl>,
    },
    MDIForm {
        properties: MDIFormProperties,
        controls: Vec<VB6Control>,
        menus: Vec<VB6MenuControl>,
    },
    Custom {
        properties: CustomControlProperties,
        property_groups: Vec<VB6PropertyGroup>,
    },
}

impl VB6ControlKind {
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

/// The `Alignment` property of a `Label`, `TextBox`, or `ComboBox` control determines
/// the alignment of the text within the control. The `Alignment` property is used
/// to specify how the text is aligned within the control, such as left-aligned,
/// right-aligned, or centered.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum Alignment {
    /// The text is aligned to the left side of the control.
    ///
    /// This is the default setting.
    #[default]
    LeftJustify = 0,
    /// The text is aligned to the right side of the control.
    RightJustify = 1,
    /// The text is centered within the control.
    Center = 2,
}

/// The `BackStyle` determines whether the background of a control is opaque or transparent.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum BackStyle {
    /// The background of the control is transparent.
    Transparent = 0,
    /// The background of the control is opaque.
    ///
    /// This is the default setting.
    #[default]
    Opaque = 1,
}

/// The `Appearance` determines whether or not a control is painted at run time
/// with 3D effects.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default)]
#[repr(i32)]
pub enum Appearance {
    /// The control is painted with a flat style.
    Flat = 0,
    /// The control is painted with a 3D style.
    ///
    /// This is the default setting.
    #[default]
    ThreeD = 1,
}

/// The `BorderStyle` determines the appearance of the border of a control.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum BorderStyle {
    /// The control has no border.
    None = 0,
    /// The control has a single-line border.
    ///
    /// This is the default setting.
    #[default]
    FixedSingle = 1,
}

/// Determines the style of drag and drop operations.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum DragMode {
    /// The control does not support drag and drop operations until
    /// the program manually initiates the drag operation.
    ///
    /// This is the default setting.
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
    ///
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
    /// A solid line.
    ///
    /// This is the default setting.
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
    /// Standard pointer. The image is determined by the hovered over object.
    ///
    /// This is the default setting.
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
    /// The programmer handles all OLE drag/drop events manually.
    ///
    /// This is the default setting.
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
    ///
    /// This is the default setting.
    #[default]
    None = 0,
    /// The programmer handles all OLE drop events manually.
    Manual = 1,
}

/// Determines if the control is clipped to the bounds of the parent control.
/// This is used with the `Form` and `MDIForm` controls.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum ClipControls {
    /// The controls are not clipped to the bounds of the parent control.
    Unbounded = 0,
    /// The controls are clipped to the bounds of the parent control.
    ///
    /// This is the default setting.
    #[default]
    Clipped = 1,
}

/// Determines if the control uses standard styling or if it uses graphical styling from it's
/// picture properties.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum Style {
    /// The control uses standard styling.
    ///
    /// This is the default setting.
    #[default]
    Standard = 0,
    /// The control uses graphical styling using its appropriate picture properties.
    Graphical = 1,
}

/// Determines the fill style of the control for drawing purposes.
/// This is used with the `Form` and `PictureBox` controls.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive)]
#[repr(i32)]
pub enum FillStyle {
    /// The background is filled with a solid color.
    Solid = 0,
    /// The background is not filled.
    ///
    /// This is the default setting.
    #[default]
    Transparent = 1,
    /// The background is filled with a horizontal line pattern.
    HorizontalLine = 2,
    /// The background is filled with a vertical line pattern.
    VerticalLine = 3,
    /// The background is filled with a diagonal line pattern.
    UpwardDiagonal = 4,
    /// The background is filled with a diagonal line pattern that goes from the bottom left to the top right.
    /// This is the same as `UpwardDiagonal` but rotated 90 degrees.
    DownwardDiagonal = 5,
    /// The background is filled with a cross-hatch pattern.
    Cross = 6,
    /// The background is filled with a diagonal cross-hatch pattern.
    /// This is the same as `Cross` but rotated 45 degrees.
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
