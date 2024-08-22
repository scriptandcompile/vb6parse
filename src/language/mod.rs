mod color;
mod controls;
mod tokens;

pub use color::{
    VB6Color, VB_3D_DK_SHADOW, VB_3D_FACE, VB_3D_HIGHLIGHT, VB_3D_LIGHT, VB_3D_SHADOW,
    VB_ACTIVE_BORDER, VB_ACTIVE_TITLE_BAR, VB_APPLICATION_WORKSPACE, VB_BLACK, VB_BLUE,
    VB_BUTTON_FACE, VB_BUTTON_SHADOW, VB_BUTTON_TEXT, VB_CYAN, VB_DESKTOP, VB_GRAY_TEXT, VB_GREEN,
    VB_HIGHLIGHT, VB_HIGHLIGHT_TEXT, VB_INACTIVE_BORDER, VB_INACTIVE_CAPTION_TEXT,
    VB_INACTIVE_TITLE_BAR, VB_INFO_BACKGROUND, VB_INFO_TEXT, VB_MAGENTA, VB_MENU_BAR, VB_MENU_TEXT,
    VB_MSG_BOX, VB_MSG_BOX_TEXT, VB_RED, VB_SCROLL_BARS, VB_TITLE_BAR_TEXT, VB_WHITE,
    VB_WINDOW_BACKGROUND, VB_WINDOW_FRAME, VB_WINDOW_TEXT, VB_YELLOW,
};

pub use controls::{
    checkbox::CheckBoxProperties,
    commandbutton::CommandButtonProperties,
    menus::{MenuProperties, VB6MenuControl},
    picturebox::PictureBoxProperties,
    ComboBoxProperties, FormProperties, FrameProperties, LabelProperties, LineProperties,
    ScrollBarProperties, TextBoxProperties, VB6Control, VB6ControlKind,
};

pub use tokens::VB6Token;
