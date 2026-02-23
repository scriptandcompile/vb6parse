//! This module contains the definitions for the Visual Basic 6 language.
//!
//! The two major components of this module are:
//! the `VB6Token` enum, which defines the different types of tokens
//! that can be found in VB6 code, and the `VB6Control` enum and the associated
//! types, which define the different kinds of controls that can be used in VB6
//! forms.

pub mod color;
mod controls;
mod tokens;

pub use color::{
    Color, VB_3D_DK_SHADOW, VB_3D_FACE, VB_3D_HIGHLIGHT, VB_3D_LIGHT, VB_3D_SHADOW,
    VB_ACTIVE_BORDER, VB_ACTIVE_TITLE_BAR, VB_APPLICATION_WORKSPACE, VB_BLACK, VB_BLUE,
    VB_BUTTON_FACE, VB_BUTTON_SHADOW, VB_BUTTON_TEXT, VB_CYAN, VB_DESKTOP, VB_GRAY_TEXT, VB_GREEN,
    VB_HIGHLIGHT, VB_HIGHLIGHT_TEXT, VB_INACTIVE_BORDER, VB_INACTIVE_CAPTION_TEXT,
    VB_INACTIVE_TITLE_BAR, VB_INFO_BACKGROUND, VB_INFO_TEXT, VB_MAGENTA, VB_MENU_BAR, VB_MENU_TEXT,
    VB_MSG_BOX, VB_MSG_BOX_TEXT, VB_RED, VB_SCROLL_BARS, VB_TITLE_BAR_TEXT, VB_WHITE,
    VB_WINDOW_BACKGROUND, VB_WINDOW_FRAME, VB_WINDOW_TEXT, VB_YELLOW,
};

pub use crate::files::common::PropertyGroup;

pub use controls::{
    checkbox::{CheckBoxProperties, CheckBoxValue},
    combobox::{ComboBoxProperties, ComboBoxStyle},
    commandbutton::CommandButtonProperties,
    custom::CustomControlProperties,
    data::{
        BOFAction, ConnectionType, DataProperties, DatabaseDriverType, DefaultCursorType,
        EOFAction, RecordSetType,
    },
    dirlistbox::DirListBoxProperties,
    drivelistbox::DriveListBoxProperties,
    filelistbox::{
        ArchiveAttribute, FileListBoxProperties, HiddenAttribute, NormalAttribute,
        ReadOnlyAttribute, SystemAttribute,
    },
    form::{
        ControlBox, FormBorderStyle, FormProperties, MaxButton, MinButton, PaletteMode,
        ShowInTaskbar, WhatsThisButton,
    },
    frame::FrameProperties,
    image::ImageProperties,
    label::{LabelProperties, WordWrap},
    line::LineProperties,
    listbox::{ListBoxProperties, ListBoxStyle},
    mdiform::MDIFormProperties,
    menus::{MenuControl, MenuProperties, NegotiatePosition, ShortCut},
    ole::{AutoActivate, DisplayType, OLEProperties, OLETypeAllowed, UpdateOptions},
    optionbutton::{OptionButtonProperties, OptionButtonValue},
    picturebox::PictureBoxProperties,
    scrollbars::ScrollBarProperties,
    shape::{Shape, ShapeProperties},
    textbox::{MultiLine, ScrollBars, TextBoxProperties},
    timer::TimerProperties,
    Activation, Align, Alignment, Appearance, AutoRedraw, AutoSize, BackStyle, BorderStyle,
    CausesValidation, ClipControls, Control, ControlKind, DragMode, DrawMode, DrawStyle, FillStyle,
    Font, FontTransparency, Form, FormLinkMode, FormRoot, HasDeviceContext, JustifyAlignment,
    LinkMode, MDIForm, MousePointer, Movability, MultiSelect, OLEDragMode, OLEDropMode, ScaleMode,
    SizeMode, StartUpPosition, Style, TabStop, TextDirection, UseMaskColor, Visibility,
    WhatsThisHelp, WindowState,
};

pub use tokens::Token;
