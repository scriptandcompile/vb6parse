//! VB6 Control definitions and properties.
//!
//! This module contains the definitions for various VB6 controls,
//! including their properties and enumerations used to represent
//! different settings for these controls.
//! Each control is represented as a struct with associated properties,
//! and enumerations are used to define specific options for properties
//! such as alignment, visibility, and behavior.
//! This module is essential for parsing and representing VB6 forms
//! and their controls in a structured manner.
//!
//! References to official Microsoft documentation are provided for
//! each property and enumeration to ensure accuracy and completeness.
//!
//! Modules for individual controls are also included, each defining
//! the properties specific to that control type.
//!

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

use std::fmt::{Display, Formatter};

use num_enum::TryFromPrimitive;
use serde::Serialize;

use crate::PropertyGroup;

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
    menus::{MenuControl, MenuProperties},
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    serde::Serialize,
    Default,
    TryFromPrimitive,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
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
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    serde::Serialize,
    Default,
    TryFromPrimitive,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
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
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    serde::Serialize,
    Default,
    TryFromPrimitive,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
///
/// A device context is a Windows data structure that defines a set of graphic
/// objects and their associated attributes, and it defines a mapping between
/// the logical coordinates and device coordinates for a particular device, such
/// as a display or printer.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245860(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum HasDeviceContext {
    /// The control does not have a device context.
    No = 0,
    /// The control has a device context.
    ///
    /// This is the default setting.
    #[default]
    Yes = -1,
}

/// Determines whether the color assigned in the `mask_color` property is used
/// as a mask.
/// That is, if it is used to create transparent regions.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445753(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum UseMaskColor {
    /// The control does not use the mask color.
    ///
    /// This is the default setting.
    #[default]
    DoNotUseMaskColor = 0,
    /// The color assigned to the `mask_color` property is used as a mask,
    /// creating a transparent region wherever that color is.
    UseMaskColor = -1,
}

/// Determines if the control causes validation.
/// In VB6, the `CausesValidation` property determines whether a control causes validation
/// to occur when the user attempts to move focus from the control.
///
/// If `CausesValidation` is set to `true`, validation occurs when the user attempts to move
/// focus from the control to another control.
///
/// If `CausesValidation` is set to `false`, validation does not occur when the user attempts
/// to move focus from the control to another control.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245065(v=vs.60))
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    serde::Serialize,
    Default,
    TryFromPrimitive,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum CausesValidation {
    /// The control does not cause validation.
    ///
    /// The control from which the focus has shifted does not fire its `Validate` event.
    No = 0,
    /// The control causes validation.
    /// The control from which the focus has shifted fires its `Validate` event.
    ///
    /// This is the default setting.
    #[default]
    Yes = -1,
}

/// The `Movability` property of a `Form` control determines whether the
/// form can be moved by the user. If the form is not moveable, the user cannot
/// move the form by dragging its title bar or by using the arrow keys.
/// If the form is moveable, the user can move the form by dragging its title
/// bar or by using the arrow keys.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa235194(v=vs.60))
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    Default,
    TryFromPrimitive,
    serde::Serialize,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum Movability {
    /// The form is not moveable.
    Fixed = 0,
    /// The form is moveable.
    ///
    /// This is the default setting.
    #[default]
    Moveable = -1,
}

/// Determines whether background text and graphics on a `Form` or a `PictureBox`
/// control are displayed in the spaces around characters.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267490(v=vs.60))
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    Default,
    TryFromPrimitive,
    serde::Serialize,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum FontTransparency {
    /// Masks existing background graphics and text around the characters of a
    /// font.
    Opaque = 0,
    /// Permits background graphics and text to show around the spaces of the
    /// characters in a font.
    ///
    /// This is the default setting.
    #[default]
    Transparent = -1,
}

/// Determines whether context-sensitive Help uses the What's This pop-up
/// (provided by Help in 32-bit Windows operating systems) or the main Help
/// window.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445772(v=vs.60))
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    Default,
    TryFromPrimitive,
    serde::Serialize,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum WhatsThisHelp {
    /// The application uses the F1 key to start Windows Help and load the topic
    /// identified by the `help_context_id` property.
    ///
    /// This is the default setting.
    #[default]
    F1Help = 0,
    /// The application uses one of the "What's This?" access techniques to
    /// start Windows Help and load a topic identified by the
    /// `help_context_id` property.
    WhatsThisHelp = -1,
}

/// Determines the type of link used for a DDE conversation and activates the
/// connection.
///
/// Forms allow a destination application to initiate a conversation with a
/// Visual Basic source form as specified by the destination applications
/// `application**|topic!**item` expression.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa235154(v=vs.60))
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    serde::Serialize,
    Default,
    TryFromPrimitive,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum FormLinkMode {
    /// No DDE interaction. No destination application can initiate a conversation
    /// with the source form as the topic, and no application can poke data to
    /// the form.
    ///
    /// This is the default setting.
    #[default]
    None = 0,
    /// Allows any `Label`, `PictureBox`, or `TextBox` control on a form to supply
    /// data to any destination application that establishes a DDE conversation
    /// with the form. If such a link exists, Visual Basic automatically
    /// notifies the destination whenever the contents of a control are changed.
    /// In addition, a destination application can poke data to any `Label`,
    /// `PictureBox`, or `TextBox` control on the form.
    Source = 1,
}

/// Controls the display state of a form from normal, minimized, or maximized.
/// This is used with the `Form` and `MDIForm` controls.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445778(v=vs.60))
#[derive(
    Debug,
    PartialEq,
    Eq,
    Clone,
    serde::Serialize,
    Default,
    TryFromPrimitive,
    Copy,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum WindowState {
    /// The form is in its normal state.
    ///
    /// This is the default setting.
    #[default]
    Normal = 0,
    /// The form is minimized (minimized to an icon0).
    Minimized = 1,
    /// The form is maximized (enlarged to maximum size).
    Maximized = 2,
}

/// The `StartUpPosition` property of a `Form` or `MDIForm` control determines
/// the initial position of the form when it first appears.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445708(v=vs.60))
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize, Default, Copy, Hash, PartialOrd, Ord)]
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
    /// Position in upper-left corner of screen.
    ///
    /// The `WindowsDefault` variant is saved as a 3 in the VB6 file.
    ///
    /// This is the default setting.
    WindowsDefault,
}

/// Represents either a reference to an external resource within a *.frx file or an embedded value.
///
/// This is used to represent properties that can either be stored directly within the VB6 form file
/// or as a reference to an external resource stored in the associated *.frx file.
///
/// The `Reference` variant contains the filename and offset within the *.frx file where the resource can be found.
/// The `Value` variant contains the actual value of type `T`.
///
/// This is useful for handling properties such as images, icons, or other binary data that may be
/// stored externally to keep the form file size manageable.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub enum ReferenceOrValue<T> {
    Reference { filename: String, offset: u32 },
    Value(T),
}

/// Represents a VB6 control.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct Control {
    /// The name of the control.
    pub name: String,
    /// The tag of the control.
    pub tag: String,
    /// The index of the control.
    pub index: i32,
    /// The kind of control.
    pub kind: ControlKind,
}

impl Display for Control {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "Control: {} ({})", self.name, self.kind)
    }
}

/// The `ControlKind` determines the specific kind of control that the `Control` represents.
///
/// Each variant contains the properties that are specific to that kind of control.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub enum ControlKind {
    /// A command button control.
    CommandButton {
        /// The properties of the command button control.
        properties: CommandButtonProperties,
    },
    /// A data control.
    Data {
        /// The properties of the data control.
        properties: DataProperties,
    },
    /// A text box control.
    TextBox {
        /// The properties of the text box control.
        properties: TextBoxProperties,
    },
    /// A check box control.
    CheckBox {
        /// The properties of the check box control.
        properties: CheckBoxProperties,
    },
    /// A line control.
    Line {
        /// The properties of the line control.
        properties: LineProperties,
    },
    /// A shape control.
    Shape {
        /// The properties of the shape control.
        properties: ShapeProperties,
    },
    /// A list box control.
    ListBox {
        /// The properties of the list box control.
        properties: ListBoxProperties,
    },
    /// A timer control.
    Timer {
        /// The properties of the timer control.
        properties: TimerProperties,
    },
    /// A label control.
    Label {
        /// The properties of the label control.
        properties: LabelProperties,
    },
    /// A frame control.
    Frame {
        /// The properties of the frame control.
        properties: FrameProperties,
        /// The child controls of the frame control.
        controls: Vec<Control>,
    },
    /// A picture box control.
    PictureBox {
        /// The properties of the picture box control.
        properties: PictureBoxProperties,
    },
    /// A file list box control.
    FileListBox {
        /// The properties of the file list box control.
        properties: FileListBoxProperties,
    },
    /// A drive list box control.
    DriveListBox {
        /// The properties of the drive list box control.
        properties: DriveListBoxProperties,
    },
    /// A directory list box control.
    DirListBox {
        /// The properties of the directory list box control.
        properties: DirListBoxProperties,
    },
    /// An OLE control.
    Ole {
        /// The properties of the OLE control.
        properties: OLEProperties,
    },
    /// An option button control.
    OptionButton {
        /// The properties of the option button control.
        properties: OptionButtonProperties,
    },
    /// An image control.
    Image {
        /// The properties of the image control.
        properties: ImageProperties,
    },
    /// A combo box control.
    ComboBox {
        /// The properties of the combo box control.
        properties: ComboBoxProperties,
    },
    /// A horizontal scroll bar control.
    HScrollBar {
        /// The properties of the horizontal scroll bar control.
        properties: ScrollBarProperties,
    },
    /// A vertical scroll bar control.
    VScrollBar {
        /// The properties of the vertical scroll bar control.
        properties: ScrollBarProperties,
    },
    /// A menu control.
    Menu {
        /// The properties of the menu control.
        properties: MenuProperties,
        /// The sub-menus of the menu control.
        sub_menus: Vec<MenuControl>,
    },
    /// A form control.
    Form {
        /// The properties of the form control.
        properties: FormProperties,
        /// The child controls of the form control.
        controls: Vec<Control>,
        /// The menus of the form control.
        menus: Vec<MenuControl>,
    },
    /// An MDI form control.
    MDIForm {
        /// The properties of the MDI form control.
        properties: MDIFormProperties,
        /// The child controls of the MDI form control.
        controls: Vec<Control>,
        /// The menus of the MDI form control.
        menus: Vec<MenuControl>,
    },
    /// A custom control.
    Custom {
        /// The properties of the custom control.
        properties: CustomControlProperties,
        /// The property groups of the custom control.
        property_groups: Vec<PropertyGroup>,
    },
}

impl Display for ControlKind {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            ControlKind::CommandButton { .. } => write!(f, "CommandButton"),
            ControlKind::Data { .. } => write!(f, "Data"),
            ControlKind::TextBox { .. } => write!(f, "TextBox"),
            ControlKind::CheckBox { .. } => write!(f, "CheckBox"),
            ControlKind::Line { .. } => write!(f, "Line"),
            ControlKind::Shape { .. } => write!(f, "Shape"),
            ControlKind::ListBox { .. } => write!(f, "ListBox"),
            ControlKind::Timer { .. } => write!(f, "Timer"),
            ControlKind::Label { .. } => write!(f, "Label"),
            ControlKind::Frame { .. } => write!(f, "Frame"),
            ControlKind::PictureBox { .. } => write!(f, "PictureBox"),
            ControlKind::FileListBox { .. } => write!(f, "FileListBox"),
            ControlKind::DriveListBox { .. } => write!(f, "DriveListBox"),
            ControlKind::DirListBox { .. } => write!(f, "DirListBox"),
            ControlKind::Ole { .. } => write!(f, "OLE"),
            ControlKind::OptionButton { .. } => write!(f, "OptionButton"),
            ControlKind::Image { .. } => write!(f, "Image"),
            ControlKind::ComboBox { .. } => write!(f, "ComboBox"),
            ControlKind::HScrollBar { .. } => write!(f, "HScrollBar"),
            ControlKind::VScrollBar { .. } => write!(f, "VScrollBar"),
            ControlKind::Menu { .. } => write!(f, "Menu"),
            ControlKind::Form { .. } => write!(f, "Form"),
            ControlKind::MDIForm { .. } => write!(f, "MDIForm"),
            ControlKind::Custom { .. } => write!(f, "Custom"),
        }
    }
}

/// Helper methods for `ControlKind`.
impl ControlKind {
    /// Returns `true` if the control kind is a `Menu`.
    #[must_use]
    pub fn is_menu(&self) -> bool {
        matches!(self, ControlKind::Menu { .. })
    }
}

/// Determines whether an object is displayed in any size anywhere on a form or
/// whether it's displayed at the top, bottom, left, or right of the form and is
/// automatically sized to fit the form's width.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267259(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum Align {
    /// The control is not docked to any side of the parent control.
    /// This setting is ignored if the object is on an `MDIForm`.
    ///
    /// This is the default setting in a non-MDI form.
    #[default]
    None = 0,
    /// The top of the control is at the top of the form, and its width is equal
    /// to the form's `ScaleWidth` property setting.
    ///
    /// This is the default setting in an MDI form.
    Top = 1,
    /// The bottom of the control is at the bottom of the form, and its width is
    /// equal to the form's `ScaleWidth` property setting.
    Bottom = 2,
    /// The left side of the control is at the left of the form, and its width
    /// is equal to the form's `ScaleWidth` property setting.
    Left = 3,
    /// The right side of the control is at the right of the form, and its width
    /// is equal to the form's `ScaleWidth` property setting.
    Right = 4,
}

/// Determines the alignment of a `CheckBox` or `OptionButton` control.
///
/// This enum is the 'Alignment' property in VB6 specifically for `CheckBox` and
/// `OptionButton` controls only.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267261(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum JustifyAlignment {
    /// The text is left-aligned. The control is right-aligned.
    ///
    /// This is the default setting.
    #[default]
    LeftJustify = 0,
    /// The text is right-aligned. The control is left-aligned.
    RightJustify = 1,
}

/// The `Alignment` property of a `Label` and `TextBox` control determines
/// the alignment of the text within the control. The `Alignment` property is used
/// to specify how the text is aligned within the control, such as left-aligned,
/// right-aligned, or centered.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa267261(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum Alignment {
    /// The text is left-aligned within the control.
    ///
    /// This is the default setting.
    #[default]
    LeftJustify = 0,
    /// The text is right-aligned within the control.
    RightJustify = 1,
    /// The text is centered within the control.
    Center = 2,
}

/// Indicates whether a `Label` control or the background of a `Shape` control
/// is transparent or opaque.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245038(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum BackStyle {
    /// The transparent background color and any graphics are visible behind the
    /// control.
    Transparent = 0,
    /// The control's `BackColor` property setting fills the control and
    /// obscures any color or graphics behind it.
    ///
    /// This is the default setting.
    #[default]
    Opaque = 1,
}

/// The `Appearance` determines whether or not a control is painted at run time
/// with 3D effects.
///
/// Note:
///
/// If set to `ThreeD` (1) at design time, the `Appearance` property draws
/// controls with three-dimensional effects. If the form's `BorderStyle`
/// property is set to `FixedDouble` (vbFixedDouble, or 3), the caption and
/// border of the form are also painted with three-dimensional effects.
///
/// Setting the `Appearance` property to `ThreeD` (1) also causes the form and its
/// controls to have their `BackColor` property set to the color selected for 3D
/// Objects in the `Appearance` tab of the operating system's Display Properties
/// dialog box.
///
/// Setting the `Appearance` property to `ThreeD` (1) for an `MDIForm` object
/// affects only the MDI parent form. To have three-dimensional effects on MDI
/// child forms, you must set each child form's `Appearance` property to
/// `ThreeD` (1).
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa244932(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default, Copy, Hash, PartialOrd, Ord,
)]
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
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa245047(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum BorderStyle {
    /// The control has no border.
    ///
    /// This is the default setting for `Image` and `Label` controls.
    None = 0,
    /// The control has a single-line border.
    ///
    /// This is the default setting for `PictureBox`, `TextBox`, `OLE` container
    /// controls.
    #[default]
    FixedSingle = 1,
}

/// Determines the style of drag and drop operations.
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
    /// Inversion of the combination of the colors in the pen and the background color but not in both (ie, NXOR).
    NotXorPen = 10,
    /// No operation is performed. The output remains unchanged. In effect, this turns drawing off (No Operation).
    Nop = 11,
    /// The combination of the display color and the inverse of the pen color.
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
    /// Arrow with an hourglass.
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
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

/// Determines the link mode of a control for DDE conversations.
/// This is used with the `Form` control.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa235154(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, TryFromPrimitive, Default, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum LinkMode {
    /// No DDE interaction. No destination application can initiate a conversation
    /// with the source control as the topic, and no application can poke data to
    /// the control.
    #[default]
    None = 0,
    /// Allows any `Label`, `PictureBox`, or `TextBox` control on a form to supply
    /// data to any destination application that establishes a DDE conversation
    /// with the control. If such a link exists, Visual Basic automatically
    /// notifies the destination whenever the contents of a control are changed.
    /// In addition, a destination application can poke data to any `Label`,
    /// `PictureBox`, or `TextBox` control on the form.
    Automatic = 1,
    /// Allows any `Label`, `PictureBox`, or `TextBox` control on a form to supply
    /// data to any destination application that establishes a DDE conversation
    /// with the control. However, Visual Basic does not automatically notify
    /// the destination whenever the contents of a control are changed. In
    /// addition, a destination application can poke data to any `Label`,
    /// `PictureBox`, or `TextBox` control on the form.
    Manual = 2,
    /// Allows any `Label`, `PictureBox`, or `TextBox` control on a form to supply
    /// data to any destination application that establishes a DDE conversation
    /// with the control. Visual Basic automatically notifies the destination
    /// whenever the contents of a control are changed. However, a destination
    /// application cannot poke data to any `Label`, `PictureBox`, or `TextBox`
    /// control on the form.
    Notify = 3,
}

/// Determines the multi-select behavior of a `ListBox` control.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa235198(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum MultiSelect {
    /// The user cannot select more than one item in the list box.
    #[default]
    None = 0,
    /// The user can select multiple items in the list box by holding down the
    /// `SHIFT` key while clicking items.
    Simple = 1,
    /// The user can select multiple items in the list box by holding down the
    /// `CTRL` key while clicking items.
    Extended = 2,
}

/// Determines the scale mode of the control for sizing and positioning.
/// This is used with the `Form` and `PictureBox` controls.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445668(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum ScaleMode {
    /// Indicates that one or more of the `ScaleHeight`, `ScaleWidth`, `ScaleLeft`, and `ScaleTop` properties are set to custom values.
    User = 0,
    /// The control uses twips as the unit of measurement. (1440 twips per logical inch; 567 twips per logical centimeter).
    #[default]
    Twip = 1,
    /// The control uses Points as the unit of measurement. (72 points per logical inch).
    Point = 2,
    /// The control uses Pixels as the unit of measurement. (The number of pixels per logical inch depends on the system's display settings).
    Pixel = 3,
    /// The control uses Characters as the unit of measurement. Character (horizontal = 120 twips per unit; vertical = 240 twips per unit).
    Character = 4,
    /// The control uses Inches as the unit of measurement.
    Inches = 5,
    /// The control uses Millimeters as the unit of measurement.
    Millimeter = 6,
    /// The control uses Centimeters as the unit of measurement.
    Centimeter = 7,
    /// The control uses `HiMetrics` as the unit of measurement.
    HiMetric = 8,
    /// The control uses the Units used by the control's container to determine the control's position.
    ContainerPosition = 9,
    /// The control uses the Units used by the control's container to determine the control's size.
    ContainerSize = 10,
}

/// Determines how the control sizes the picture within its bounds.
/// This is used with the `Image` and `PictureBox` controls.
///
/// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445695(v=vs.60))
#[derive(
    Debug, PartialEq, Eq, Clone, Serialize, Default, TryFromPrimitive, Copy, Hash, PartialOrd, Ord,
)]
#[repr(i32)]
pub enum SizeMode {
    /// The picture is displayed in its original size. If the picture is larger than
    /// the control, the picture is clipped to fit within the control's bounds.
    ///
    /// If the picture is smaller than the control, the picture is displayed in the
    /// top-left corner of the control, and the remaining area of the control is
    /// left blank.
    #[default]
    Clip = 0,
    /// The picture is stretched or shrunk to fit the control's bounds.
    Stretch = 1,
    /// The control is automatically resized to fit the picture.
    AutoSize = 2,
    /// The picture is stretched or shrunk to fit the control's bounds while maintaining its aspect ratio.
    Zoom = 3,
}
