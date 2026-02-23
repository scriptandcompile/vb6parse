//! Properties for a `FileListBox` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::FileListBox`](crate::language::controls::ControlKind::FileListBox).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use std::convert::TryFrom;
use std::fmt::Display;
use std::str::FromStr;

use crate::errors::{ErrorKind, FormError};
use crate::files::common::Properties;
use crate::language::color::Color;
use crate::language::controls::{
    Activation, Appearance, CausesValidation, DragMode, Font, MousePointer, MultiSelect,
    OLEDragMode, OLEDropMode, ReferenceOrValue, TabStop, Visibility,
};

use image::DynamicImage;
use num_enum::TryFromPrimitive;
use serde::Serialize;

/// The `ArchiveAttribute` enum represents the archive bit of a file.
/// It is used to indicate whether a file should be included or excluded from being
/// shown in the `FileListBox` control based on its archive status.
#[derive(
    Debug,
    PartialEq,
    Default,
    Clone,
    Copy,
    serde::Serialize,
    TryFromPrimitive,
    Eq,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum ArchiveAttribute {
    /// The file is excluded in the `FileListBox` if it has the archive attribute bit set.
    Exclude = 0,
    /// The file is included in the `FileListBox` if it has the archive attribute bit set.
    ///
    /// This is the default value.
    #[default]
    Include = -1,
}

impl Display for ArchiveAttribute {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            ArchiveAttribute::Exclude => "Exclude",
            ArchiveAttribute::Include => "Include",
        };
        write!(f, "{text}")
    }
}

impl TryFrom<&str> for ArchiveAttribute {
    type Error = ErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(ArchiveAttribute::Exclude),
            "-1" => Ok(ArchiveAttribute::Include),
            _ => Err(ErrorKind::Form(FormError::InvalidArchiveAttribute {
                value: value.to_string(),
            })),
        }
    }
}

impl TryFrom<bool> for ArchiveAttribute {
    type Error = ErrorKind;

    fn try_from(value: bool) -> Result<Self, Self::Error> {
        if value {
            Ok(ArchiveAttribute::Include)
        } else {
            Ok(ArchiveAttribute::Exclude)
        }
    }
}

impl FromStr for ArchiveAttribute {
    type Err = ErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        ArchiveAttribute::try_from(s)
    }
}

/// The `HiddenAttribute` enum represents the hidden bit of a file.
/// It is used to indicate whether a file should be included or excluded from being
/// shown in the `FileListBox` control based on its hidden status.
#[derive(
    Debug,
    PartialEq,
    Default,
    Clone,
    Copy,
    serde::Serialize,
    TryFromPrimitive,
    Eq,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum HiddenAttribute {
    /// The file is excluded in the `FileListBox` if it has the hidden attribute bit set.
    ///
    /// This is the default value.
    #[default]
    Exclude = 0,
    /// The file is included in the `FileListBox` if it has the hidden attribute bit set.
    Include = -1,
}

impl Display for HiddenAttribute {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            HiddenAttribute::Exclude => "Exclude",
            HiddenAttribute::Include => "Include",
        };
        write!(f, "{text}")
    }
}

impl TryFrom<&str> for HiddenAttribute {
    type Error = ErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(HiddenAttribute::Exclude),
            "-1" => Ok(HiddenAttribute::Include),
            _ => Err(ErrorKind::Form(FormError::InvalidHiddenAttribute {
                value: value.to_string(),
            })),
        }
    }
}

impl FromStr for HiddenAttribute {
    type Err = ErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        HiddenAttribute::try_from(s)
    }
}

/// The `ReadOnlyAttribute` enum represents the read-only bit of a file.
/// It is used to indicate whether a file should be included or excluded from being
/// shown in the `FileListBox` control based on its read-only status.
#[derive(
    Debug,
    PartialEq,
    Default,
    Clone,
    Copy,
    serde::Serialize,
    TryFromPrimitive,
    Eq,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum ReadOnlyAttribute {
    /// The file is excluded in the `FileListBox` if it has the read-only attribute bit set.
    Exclude = 0,
    /// The file is included in the `FileListBox` if it has the read-only attribute bit set.
    ///
    /// This is the default value.
    #[default]
    Include = -1,
}

impl TryFrom<&str> for ReadOnlyAttribute {
    type Error = ErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(ReadOnlyAttribute::Exclude),
            "-1" => Ok(ReadOnlyAttribute::Include),
            _ => Err(ErrorKind::Form(FormError::InvalidReadOnlyAttribute {
                value: value.to_string(),
            })),
        }
    }
}

impl TryFrom<bool> for ReadOnlyAttribute {
    type Error = ErrorKind;

    fn try_from(value: bool) -> Result<Self, Self::Error> {
        if value {
            Ok(ReadOnlyAttribute::Include)
        } else {
            Ok(ReadOnlyAttribute::Exclude)
        }
    }
}

impl FromStr for ReadOnlyAttribute {
    type Err = ErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        ReadOnlyAttribute::try_from(s)
    }
}

impl Display for ReadOnlyAttribute {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            ReadOnlyAttribute::Exclude => "Exclude",
            ReadOnlyAttribute::Include => "Include",
        };
        write!(f, "{text}")
    }
}

/// The `SystemAttribute` enum represents the system bit of a file.
/// It is used to indicate whether a file should be included or excluded from being
/// shown in the `FileListBox` control based on its system status.
#[derive(
    Debug,
    PartialEq,
    Default,
    Clone,
    Copy,
    serde::Serialize,
    TryFromPrimitive,
    Eq,
    Hash,
    PartialOrd,
    Ord,
)]
#[repr(i32)]
pub enum SystemAttribute {
    /// The file is excluded in the `FileListBox` if it has the system attribute bit set.
    ///
    /// This is the default value.
    #[default]
    Exclude = 0,
    /// The file is included in the `FileListBox` if it has the system attribute bit set.
    Include = -1,
}

impl TryFrom<&str> for SystemAttribute {
    type Error = ErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(SystemAttribute::Exclude),
            "-1" => Ok(SystemAttribute::Include),
            _ => Err(ErrorKind::Form(FormError::InvalidSystemAttribute {
                value: value.to_string(),
            })),
        }
    }
}

impl TryFrom<bool> for SystemAttribute {
    type Error = ErrorKind;

    fn try_from(value: bool) -> Result<Self, Self::Error> {
        if value {
            Ok(SystemAttribute::Include)
        } else {
            Ok(SystemAttribute::Exclude)
        }
    }
}

impl FromStr for SystemAttribute {
    type Err = ErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        SystemAttribute::try_from(s)
    }
}

impl Display for SystemAttribute {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            SystemAttribute::Exclude => "Exclude",
            SystemAttribute::Include => "Include",
        };
        write!(f, "{text}")
    }
}

/// The `NormalAttribute` determines if the `FileListBox` control will include
/// files that are not hidden, read-only, archive, or system files.
#[derive(
    Debug,
    PartialEq,
    Default,
    Clone,
    Copy,
    serde::Serialize,
    TryFromPrimitive,
    Hash,
    PartialOrd,
    Eq,
    Ord,
)]
#[repr(i32)]
pub enum NormalAttribute {
    /// The file is excluded in the `FileListBox` if it has none of the hidden, read-only, archive, or system attributes set.
    Exclude = 0,
    /// The file is included in the `FileListBox` if it has none of the hidden, read-only, archive, or system attributes set.
    ///
    /// This is the default value.
    #[default]
    Include = -1,
}

impl TryFrom<&str> for NormalAttribute {
    type Error = ErrorKind;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        match value {
            "0" => Ok(NormalAttribute::Exclude),
            "-1" => Ok(NormalAttribute::Include),
            _ => Err(ErrorKind::Form(FormError::InvalidNormalAttribute {
                value: value.to_string(),
            })),
        }
    }
}

impl TryFrom<bool> for NormalAttribute {
    type Error = ErrorKind;

    fn try_from(value: bool) -> Result<Self, Self::Error> {
        if value {
            Ok(NormalAttribute::Include)
        } else {
            Ok(NormalAttribute::Exclude)
        }
    }
}

impl FromStr for NormalAttribute {
    type Err = ErrorKind;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        NormalAttribute::try_from(s)
    }
}

impl Display for NormalAttribute {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let text = match self {
            NormalAttribute::Exclude => "Exclude",
            NormalAttribute::Include => "Include",
        };
        write!(f, "{text}")
    }
}

/// Properties for a `FileListBox` control.
///
/// This is used as an enum variant of
/// [`ControlKind::FileListBox`](crate::language::controls::ControlKind::FileListBox).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone)]
pub struct FileListBoxProperties {
    /// The appearance of the `FileListBox`.
    pub appearance: Appearance,
    /// The archive attribute of the `FileListBox`.
    pub archive: ArchiveAttribute,
    /// The background color of the `FileListBox`.
    pub back_color: Color,
    /// Whether the `FileListBox` causes validation.
    pub causes_validation: CausesValidation,
    /// The drag icon of the `FileListBox`.
    pub drag_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// The drag mode of the `FileListBox`.
    pub drag_mode: DragMode,
    /// Whether the `FileListBox` is enabled.
    pub enabled: Activation,
    /// The font style for the form.
    pub font: Option<Font>,
    /// The foreground color of the `FileListBox`.
    pub fore_color: Color,
    /// The height of the `FileListBox`.
    pub height: i32,
    /// The help context ID of the `FileListBox`.
    pub help_context_id: i32,
    /// The hidden attribute of the `FileListBox`.
    pub hidden: HiddenAttribute,
    /// The left position of the `FileListBox`.
    pub left: i32,
    /// The mouse icon of the `FileListBox`.
    pub mouse_icon: Option<ReferenceOrValue<DynamicImage>>,
    /// The mouse pointer of the `FileListBox`.
    pub mouse_pointer: MousePointer,
    /// The multi-select mode of the `FileListBox`.
    pub multi_select: MultiSelect,
    /// The normal attribute of the `FileListBox`.
    pub normal: NormalAttribute,
    /// The OLE drag mode of the `FileListBox`.
    pub ole_drag_mode: OLEDragMode,
    /// The OLE drop mode of the `FileListBox`.
    pub ole_drop_mode: OLEDropMode,
    /// The file pattern filter of the `FileListBox`.
    pub pattern: String,
    /// The read-only attribute of the `FileListBox`.
    pub read_only: ReadOnlyAttribute,
    /// The system attribute of the `FileListBox`.
    pub system: SystemAttribute,
    /// The tab index of the `FileListBox`.
    pub tab_index: i32,
    /// The tab stop of the `FileListBox`.
    pub tab_stop: TabStop,
    /// The tool tip text of the `FileListBox`.
    pub tool_tip_text: String,
    /// The top position of the `FileListBox`.
    pub top: i32,
    /// The visibility of the `FileListBox`.
    pub visible: Visibility,
    /// The "What's This?" help ID of the `FileListBox`.
    pub whats_this_help_id: i32,
    /// The width of the `FileListBox`.
    pub width: i32,
}

impl Default for FileListBoxProperties {
    fn default() -> Self {
        FileListBoxProperties {
            appearance: Appearance::ThreeD,
            archive: ArchiveAttribute::Include,
            back_color: Color::System { index: 5 },
            causes_validation: CausesValidation::Yes,
            drag_icon: None,
            drag_mode: DragMode::Manual,
            enabled: Activation::Enabled,
            font: Some(Font::default()),
            fore_color: Color::System { index: 8 },
            height: 1260,
            help_context_id: 0,
            hidden: HiddenAttribute::Exclude,
            left: 710,
            mouse_icon: None,
            mouse_pointer: MousePointer::Default,
            multi_select: MultiSelect::None,
            normal: NormalAttribute::Include,
            ole_drag_mode: OLEDragMode::Manual,
            ole_drop_mode: OLEDropMode::default(),
            pattern: "*.*".into(), // Default pattern for file selection
            read_only: ReadOnlyAttribute::Include,
            system: SystemAttribute::Include,
            tab_index: 0,
            tab_stop: TabStop::Included,
            tool_tip_text: String::new(),
            top: 480,
            visible: Visibility::Visible,
            whats_this_help_id: 0,
            width: 735,
        }
    }
}

impl Serialize for FileListBoxProperties {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut s = serializer.serialize_struct("FileListBoxProperties", 27)?;
        s.serialize_field("appearance", &self.appearance)?;
        s.serialize_field("archive", &self.archive)?;
        s.serialize_field("back_color", &self.back_color)?;
        s.serialize_field("causes_validation", &self.causes_validation)?;

        let option_text = self.drag_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("drag_icon", &option_text)?;
        s.serialize_field("drag_mode", &self.drag_mode)?;
        s.serialize_field("enabled", &self.enabled)?;
        s.serialize_field("fore_color", &self.fore_color)?;
        s.serialize_field("height", &self.height)?;
        s.serialize_field("help_context_id", &self.help_context_id)?;
        s.serialize_field("hidden", &self.hidden)?;
        s.serialize_field("left", &self.left)?;

        let option_text = self.mouse_icon.as_ref().map(|_| "Some(DynamicImage)");

        s.serialize_field("mouse_icon", &option_text)?;
        s.serialize_field("mouse_pointer", &self.mouse_pointer)?;
        s.serialize_field("multi_select", &self.multi_select)?;
        s.serialize_field("normal", &self.normal)?;
        s.serialize_field("ole_drag_mode", &self.ole_drag_mode)?;
        s.serialize_field("ole_drop_mode", &self.ole_drop_mode)?;
        s.serialize_field("pattern", &self.pattern)?;
        s.serialize_field("read_only", &self.read_only)?;
        s.serialize_field("system", &self.system)?;
        s.serialize_field("tab_index", &self.tab_index)?;
        s.serialize_field("tab_stop", &self.tab_stop)?;
        s.serialize_field("tool_tip_text", &self.tool_tip_text)?;
        s.serialize_field("top", &self.top)?;
        s.serialize_field("visible", &self.visible)?;
        s.serialize_field("whats_this_help_id", &self.whats_this_help_id)?;
        s.serialize_field("width", &self.width)?;

        s.end()
    }
}

impl From<Properties> for FileListBoxProperties {
    fn from(prop: Properties) -> Self {
        let mut file_list_box_prop = FileListBoxProperties::default();

        file_list_box_prop.appearance =
            prop.get_property("Appearance", file_list_box_prop.appearance);
        file_list_box_prop.archive = prop.get_property("Archive", file_list_box_prop.archive);
        file_list_box_prop.back_color = prop.get_color("BackColor", file_list_box_prop.back_color);
        file_list_box_prop.causes_validation =
            prop.get_property("CausesValidation", file_list_box_prop.causes_validation);
        file_list_box_prop.drag_mode = prop.get_property("DragMode", file_list_box_prop.drag_mode);
        file_list_box_prop.enabled = prop.get_property("Enabled", file_list_box_prop.enabled);
        file_list_box_prop.fore_color = prop.get_color("ForeColor", file_list_box_prop.fore_color);
        file_list_box_prop.height = prop.get_i32("Height", file_list_box_prop.height);
        file_list_box_prop.help_context_id =
            prop.get_i32("HelpContextID", file_list_box_prop.help_context_id);
        file_list_box_prop.hidden = prop.get_property("Hidden", file_list_box_prop.hidden);
        file_list_box_prop.left = prop.get_i32("Left", file_list_box_prop.left);
        file_list_box_prop.mouse_pointer =
            prop.get_property("MousePointer", file_list_box_prop.mouse_pointer);
        file_list_box_prop.multi_select =
            prop.get_property("MultiSelect", file_list_box_prop.multi_select);
        file_list_box_prop.normal = prop.get_property("Normal", file_list_box_prop.normal);
        file_list_box_prop.ole_drag_mode =
            prop.get_property("OLEDragMode", file_list_box_prop.ole_drag_mode);
        file_list_box_prop.ole_drop_mode =
            prop.get_property("OLEDropMode", file_list_box_prop.ole_drop_mode);
        file_list_box_prop.pattern = match prop.get("Pattern") {
            Some(pattern) => pattern.into(),
            None => file_list_box_prop.pattern,
        };
        file_list_box_prop.read_only = prop.get_property("ReadOnly", file_list_box_prop.read_only);
        file_list_box_prop.system = prop.get_property("System", file_list_box_prop.system);
        file_list_box_prop.tab_index = prop.get_i32("TabIndex", file_list_box_prop.tab_index);
        file_list_box_prop.tab_stop = prop.get_property("TabStop", file_list_box_prop.tab_stop);
        file_list_box_prop.tool_tip_text = match prop.get("ToolTipText") {
            Some(tool_tip_text) => tool_tip_text.into(),
            None => file_list_box_prop.tool_tip_text,
        };
        file_list_box_prop.top = prop.get_i32("Top", file_list_box_prop.top);
        file_list_box_prop.visible = prop.get_property("Visible", file_list_box_prop.visible);
        file_list_box_prop.whats_this_help_id =
            prop.get_i32("WhatsThisHelpID", file_list_box_prop.whats_this_help_id);
        file_list_box_prop.width = prop.get_i32("Width", file_list_box_prop.width);

        file_list_box_prop
    }
}
