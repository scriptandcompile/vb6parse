use std::convert::TryFrom;
use std::vec::Vec;
use std::{collections::HashMap, fmt::Debug};

use bstr::{BStr, ByteSlice};
use either::Either;
use num_enum::TryFromPrimitive;
use serde::Serialize;
use uuid::Uuid;
use winnow::error::ParserError;
use winnow::{
    ascii::{line_ending, space0, space1},
    combinator::{alt, opt},
    error::ErrMode,
    token::{literal, take_till, take_until},
    Parser,
};

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    language::{
        CheckBoxProperties, ComboBoxProperties, CommandButtonProperties, DataProperties,
        DirListBoxProperties, DriveListBoxProperties, FileListBoxProperties, FormProperties,
        FrameProperties, ImageProperties, LabelProperties, LineProperties, ListBoxProperties,
        MDIFormProperties, MenuProperties, OLEProperties, OptionButtonProperties,
        PictureBoxProperties, ScrollBarProperties, ShapeProperties, TextBoxProperties,
        TimerProperties, VB6Color, VB6Control, VB6ControlKind, VB6MenuControl, VB6Token,
    },
    parsers::{
        header::{
            attributes_parse, key_resource_offset_line_parse, version_parse, HeaderKind,
            VB6FileAttributes, VB6FileFormatVersion,
        },
        VB6ObjectReference, VB6Stream,
    },
    vb6::{keyword_parse, line_comment_parse, vb6_parse, VB6Result},
};

use super::{header::object_parse, vb6::string_parse};

/// Represents a VB6 Form file.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct VB6FormFile<'a> {
    pub form: VB6Control<'a>,
    pub objects: Vec<VB6ObjectReference<'a>>,
    pub format_version: VB6FileFormatVersion,
    pub attributes: VB6FileAttributes<'a>,
    pub tokens: Vec<VB6Token<'a>>,
}

#[derive(Debug, PartialEq, Eq, Clone, Copy, Serialize)]
struct VB6FullyQualifiedName<'a> {
    pub namespace: &'a BStr,
    pub kind: &'a BStr,
    pub name: &'a BStr,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6PropertyGroup<'a> {
    pub name: &'a BStr,
    pub guid: Option<Uuid>,
    pub properties: HashMap<&'a BStr, Either<&'a BStr, VB6PropertyGroup<'a>>>,
}

impl Serialize for VB6PropertyGroup<'_> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("VB6PropertyGroup", 3)?;

        state.serialize_field("name", self.name)?;

        if let Some(guid) = &self.guid {
            state.serialize_field("guid", &guid.to_string())?;
        } else {
            state.serialize_field("guid", &"None")?;
        }

        state.serialize_field("properties", &self.properties)?;

        state.end()
    }
}

impl<'a> VB6FormFile<'a> {
    /// Parses a VB6 form file from a byte slice.
    ///
    /// # Arguments
    ///
    /// * `input` The byte slice to parse.
    ///
    /// # Returns
    ///
    /// A result containing the parsed VB6 form file or an error.
    ///
    /// # Errors
    ///
    /// An error will be returned if the input is not a valid VB6 form file.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::parsers::VB6FormFile;
    ///
    /// let input = b"VERSION 5.00\r
    /// Begin VB.Form frmExampleForm\r
    ///    BackColor       =   &H80000005&\r
    ///    Caption         =   \"example form\"\r
    ///    ClientHeight    =   6210\r
    ///    ClientLeft      =   60\r
    ///    ClientTop       =   645\r
    ///    ClientWidth     =   9900\r
    ///    BeginProperty Font\r
    ///       Name            =   \"Arial\"\r
    ///       Size            =   8.25\r
    ///       Charset         =   0\r
    ///       Weight          =   400\r
    ///       Underline       =   0   'False\r
    ///       Italic          =   0   'False\r
    ///       Strikethrough   =   0   'False\r
    ///    EndProperty\r
    ///    LinkTopic       =   \"Form1\"\r
    ///    ScaleHeight     =   414\r
    ///    ScaleMode       =   3  'Pixel\r
    ///    ScaleWidth      =   660\r
    ///    StartUpPosition =   2  'CenterScreen\r
    ///    Begin VB.Menu mnuFile\r
    ///       Caption         =   \"&File\"\r
    ///       Begin VB.Menu mnuOpenImage\r
    ///          Caption         =   \"&Open image\"\r
    ///       End\r
    ///    End\r
    /// End\r
    /// Attribute VB_Name = \"frmExampleForm\"\r
    /// ";
    ///
    /// let result = VB6FormFile::parse("form_parse.frm".to_owned(), &mut input.as_ref());
    ///
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse(file_name: String, input: &'a [u8]) -> Result<Self, VB6Error> {
        let mut input = VB6Stream::new(file_name, input);

        let format_version = match version_parse(HeaderKind::Form).parse_next(&mut input) {
            Ok(version) => version,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        let objects = match form_object_parse.parse_next(&mut input) {
            Ok(objects) => objects,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        match (space0, keyword_parse("BEGIN"), space1).parse_next(&mut input) {
            Ok(_) => (),
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        let form = match block_parse.parse_next(&mut input) {
            Ok(form) => form,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        let attributes = match attributes_parse.parse_next(&mut input) {
            Ok(attributes) => attributes,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        let tokens = match vb6_parse.parse_next(&mut input) {
            Ok(tokens) => tokens,
            Err(err) => return Err(input.error(err.into_inner().unwrap())),
        };

        Ok(VB6FormFile {
            form,
            objects,
            format_version,
            attributes,
            tokens,
        })
    }
}

fn form_object_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<Vec<VB6ObjectReference<'a>>> {
    let mut objects = vec![];

    loop {
        space0.parse_next(input)?;

        if literal::<_, _, VB6ErrorKind>("Object")
            .parse_next(input)
            .is_err()
        {
            break;
        }

        let object = object_parse.parse_next(input)?;

        objects.push(object);
    }

    Ok(objects)
}

fn block_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6Control<'a>> {
    let fully_qualified_name = property_parse.parse_next(input)?;

    let mut controls = vec![];
    let mut menus = vec![];
    let mut property_groups = vec![];
    let mut properties = HashMap::new();

    while !input.is_empty() {
        if (space0, keyword_parse("END"), space0, line_ending)
            .parse_next(input)
            .is_ok()
        {
            match build_control(
                fully_qualified_name,
                controls,
                menus,
                properties,
                property_groups,
            ) {
                Ok(control) => return Ok(control),
                Err(err) => return Err(ErrMode::Cut(err)),
            };
        }

        if (space0, keyword_parse("BEGIN"), space1)
            .parse_next(input)
            .is_ok()
        {
            let control = block_parse.parse_next(input)?;
            if control.kind.is_menu() {
                menus.push(control);
            } else {
                controls.push(control);
            }

            continue;
        }

        if let Ok(property_group) = property_group_parse.parse_next(input) {
            property_groups.push(property_group);
            continue;
        }

        if let Ok((name, _resource_file, _offset)) =
            key_resource_offset_line_parse.parse_next(input)
        {
            // TODO: At the moment we just eat the resource file look up.

            properties.insert(name, BStr::new(""));
            continue;
        }

        space0.parse_next(input)?;

        let name = take_till(1.., (b' ', b'\t', b'=')).parse_next(input)?;

        space0.parse_next(input)?;

        "=".parse_next(input)?;

        space0.parse_next(input)?;

        let value =
            alt((string_parse, take_till(1.., (' ', '\t', '\'', '\r', '\n')))).parse_next(input)?;

        properties.insert(name, value);

        space0.parse_next(input)?;

        opt(line_comment_parse).parse_next(input)?;

        line_ending.parse_next(input)?;
    }

    Err(ParserError::assert(input, "Unknown control kind"))
}

fn property_group_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6PropertyGroup<'a>> {
    (space0, keyword_parse("BeginProperty"), space1).parse_next(input)?;

    let Ok(property_group_name): VB6Result<&BStr> = take_till(1.., |c| {
        c == b'{' || c == b'\r' || c == b'\t' || c == b' ' || c == b'\n'
    })
    .parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoPropertyName));
    };

    space0.parse_next(input)?;

    // Check if we have a GUID here.
    let guid = if literal::<_, _, VB6ErrorKind>('{').parse_next(input).is_ok() {
        let uuid_segment = take_until(1.., '}').parse_next(input)?;
        '}'.parse_next(input)?;

        space0.parse_next(input)?;

        let Ok(uuid) = Uuid::parse_str(uuid_segment.to_str().unwrap()) else {
            return Err(ErrMode::Cut(VB6ErrorKind::UnableToParseUuid));
        };

        Some(uuid)
    } else {
        None
    };

    alt((line_comment_parse, line_ending)).parse_next(input)?;

    let mut property_group = VB6PropertyGroup {
        guid,
        name: property_group_name,
        properties: HashMap::new(),
    };

    while !input.is_empty() {
        if (space0, keyword_parse("EndProperty"), space0)
            .parse_next(input)
            .is_ok()
        {
            if line_ending::<_, VB6Error>.parse_next(input).is_err() {
                return Err(ErrMode::Cut(VB6ErrorKind::NoLineEndingAfterEndProperty));
            }

            break;
        }

        // looks like we have a nested property group.
        if let Ok(nested_property_group) = property_group_parse.parse_next(input) {
            property_group.properties.insert(
                nested_property_group.name,
                Either::Right(nested_property_group),
            );

            continue;
        }

        if let Ok((name, _resource_file, _offset)) =
            key_resource_offset_line_parse.parse_next(input)
        {
            // TODO: At the moment we just eat the resource file look up.
            property_group
                .properties
                .insert(name, Either::Left(BStr::new("")));

            continue;
        }

        space0.parse_next(input)?;

        let name = take_until(1.., ("\t", " ", "=")).parse_next(input)?;

        (space0, "=", space0).parse_next(input)?;

        let value =
            alt((string_parse, take_till(1.., (' ', '\t', '\'', '\r', '\n')))).parse_next(input)?;

        property_group.properties.insert(name, Either::Left(value));

        (space0, opt(line_comment_parse), line_ending).parse_next(input)?;
    }

    Ok(property_group)
}

fn property_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<VB6FullyQualifiedName<'a>> {
    let Ok(namespace) = take_until::<_, _, VB6Error>(0.., ".").parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoNamespaceAfterBegin));
    };

    if literal::<&str, _, VB6Error>(".").parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoDotAfterNamespace));
    };

    let Ok(kind) = take_until::<_, _, VB6Error>(0.., (" ", "\t")).parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoUserControlNameAfterDot));
    };

    if space1::<_, VB6Error>.parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoSpaceAfterControlKind));
    }

    let Ok(name) =
        take_till::<_, _, VB6Error>(0.., (b" ", b"\t", b"\r", b"\r\n", b"\n")).parse_next(input)
    else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoControlNameAfterControlKind));
    };

    // If there are spaces after the control name, eat those up since we don't care about them.
    space0.parse_next(input)?;
    // eat the line ending and move on.
    line_ending.parse_next(input)?;

    Ok(VB6FullyQualifiedName {
        namespace,
        kind,
        name,
    })
}

fn build_control<'a>(
    fully_qualified_name: VB6FullyQualifiedName<'a>,
    controls: Vec<VB6Control<'a>>,
    menus: Vec<VB6Control<'a>>,
    properties: HashMap<&'a BStr, &'a BStr>,
    property_groups: Vec<VB6PropertyGroup<'a>>,
) -> Result<VB6Control<'a>, VB6ErrorKind> {
    let tag_key = BStr::new("Tag");
    let tag = if properties.contains_key(tag_key) {
        properties[tag_key]
    } else {
        BStr::new("")
    };

    if fully_qualified_name.namespace != "VB" {
        let custom_control = VB6Control {
            name: fully_qualified_name.name,
            tag,
            index: 0,
            kind: VB6ControlKind::Custom {
                properties,
                property_groups,
            },
        };

        return Ok(custom_control);
    }

    let kind = match fully_qualified_name.kind.as_bytes() {
        b"Form" => {
            let form_properties = FormProperties::construct_control(properties, property_groups)?;

            let mut converted_menus = vec![];

            for menu in menus {
                if let VB6ControlKind::Menu {
                    properties: menu_properties,
                    sub_menus,
                } = menu.kind
                {
                    converted_menus.push(VB6MenuControl {
                        name: menu.name,
                        tag: menu.tag,
                        index: menu.index,
                        properties: menu_properties,
                        sub_menus,
                    });
                }
            }

            VB6ControlKind::Form {
                controls,
                properties: form_properties,
                menus: converted_menus,
            }
        }
        b"MDIForm" => {
            let mdi_form_properties =
                MDIFormProperties::construct_control(properties, property_groups)?;

            let mut converted_menus = vec![];

            for menu in menus {
                if let VB6ControlKind::Menu {
                    properties: menu_properties,
                    sub_menus,
                } = menu.kind
                {
                    converted_menus.push(VB6MenuControl {
                        name: menu.name,
                        tag: menu.tag,
                        index: menu.index,
                        properties: menu_properties,
                        sub_menus,
                    });
                }
            }

            VB6ControlKind::MDIForm {
                controls,
                properties: mdi_form_properties,
                menus: converted_menus,
            }
        }
        b"Menu" => {
            let menu_properties = MenuProperties::build_control(&properties)?;
            let mut converted_menus = vec![];

            for menu in menus {
                if let VB6ControlKind::Menu {
                    properties: menu_properties,
                    sub_menus,
                } = menu.kind
                {
                    converted_menus.push(VB6MenuControl {
                        name: menu.name,
                        tag: menu.tag,
                        index: menu.index,
                        properties: menu_properties,
                        sub_menus,
                    });
                }
            }

            VB6ControlKind::Menu {
                properties: menu_properties,
                sub_menus: converted_menus,
            }
        }
        b"Frame" => {
            let frame_properties =
                FrameProperties::construct_control(&properties, &property_groups)?;

            VB6ControlKind::Frame {
                controls,
                properties: frame_properties,
            }
        }
        b"CheckBox" => {
            let chechbox_properties = CheckBoxProperties::construct_control(&properties)?;
            VB6ControlKind::CheckBox {
                properties: chechbox_properties,
            }
        }
        b"ComboBox" => {
            let combobox_properties = ComboBoxProperties::construct_control(&properties)?;
            VB6ControlKind::ComboBox {
                properties: combobox_properties,
            }
        }
        b"CommandButton" => {
            let command_button_properties =
                CommandButtonProperties::construct_control(&properties)?;
            VB6ControlKind::CommandButton {
                properties: command_button_properties,
            }
        }
        b"Data" => {
            let data_properties = DataProperties::construct_control(&properties)?;

            VB6ControlKind::Data {
                properties: data_properties,
            }
        }
        b"DirListBox" => {
            let dir_list_box_properties = DirListBoxProperties::construct_control(&properties)?;

            VB6ControlKind::DirListBox {
                properties: dir_list_box_properties,
            }
        }
        b"DriveListBox" => {
            let drive_list_box_properties = DriveListBoxProperties::construct_control(&properties)?;

            VB6ControlKind::DriveListBox {
                properties: drive_list_box_properties,
            }
        }
        b"FileListBox" => {
            let file_list_box_properties = FileListBoxProperties::construct_control(&properties)?;

            VB6ControlKind::FileListBox {
                properties: file_list_box_properties,
            }
        }
        b"Image" => {
            let image_properties = ImageProperties::construct_control(&properties)?;

            VB6ControlKind::Image {
                properties: image_properties,
            }
        }
        b"Label" => {
            let label_properties = LabelProperties::construct_control(&properties)?;

            VB6ControlKind::Label {
                properties: label_properties,
            }
        }
        b"Line" => {
            let line_properties = LineProperties::construct_control(&properties)?;

            VB6ControlKind::Line {
                properties: line_properties,
            }
        }
        b"ListBox" => {
            let list_box_properties = ListBoxProperties::construct_control(&properties)?;

            VB6ControlKind::ListBox {
                properties: list_box_properties,
            }
        }
        b"OLE" => {
            let ole_properties = OLEProperties::construct_control(&properties)?;

            VB6ControlKind::Ole {
                properties: ole_properties,
            }
        }
        b"OptionButton" => {
            let option_button_properties = OptionButtonProperties::construct_control(&properties)?;

            VB6ControlKind::OptionButton {
                properties: option_button_properties,
            }
        }
        b"PictureBox" => {
            let picture_box_properties = PictureBoxProperties::construct_control(&properties)?;

            VB6ControlKind::PictureBox {
                properties: picture_box_properties,
            }
        }
        b"HScrollBar" => {
            let scroll_bar_properties = ScrollBarProperties::construct_control(&properties)?;

            VB6ControlKind::HScrollBar {
                properties: scroll_bar_properties,
            }
        }
        b"VScrollBar" => {
            let scroll_bar_properties = ScrollBarProperties::construct_control(&properties)?;

            VB6ControlKind::VScrollBar {
                properties: scroll_bar_properties,
            }
        }
        b"Shape" => {
            let shape_properties = ShapeProperties::construct_control(&properties)?;

            VB6ControlKind::Shape {
                properties: shape_properties,
            }
        }
        b"TextBox" => {
            let textbox_properties = TextBoxProperties::construct_control(&properties)?;

            VB6ControlKind::TextBox {
                properties: textbox_properties,
            }
        }
        b"Timer" => {
            let timer_properties = TimerProperties::construct_control(&properties)?;

            VB6ControlKind::Timer {
                properties: timer_properties,
            }
        }
        _ => {
            return Err(VB6ErrorKind::UnknownControlKind);
        }
    };

    let parent_control = VB6Control {
        name: fully_qualified_name.name,
        tag,
        index: 0,
        kind,
    };

    Ok(parent_control)
}

#[must_use]
pub fn build_property<T, B: AsRef<[u8]>, S: std::hash::BuildHasher>(
    properties: &HashMap<&BStr, &BStr, S>,
    property_key: &B,
) -> T
where
    T: Default + TryFromPrimitive + TryFrom<i32>,
{
    let key = property_key.as_ref().as_bstr();
    if !properties.contains_key(key) {
        return T::default();
    }

    let property_ascii = properties[key].to_str().unwrap();

    match property_ascii.parse::<i32>() {
        Ok(value) => T::try_from(value).unwrap_or_default(),
        Err(_) => T::default(),
    }
}

#[must_use]
pub fn build_option_property<'a, B: AsRef<[u8]>, S: std::hash::BuildHasher, T>(
    properties: &HashMap<&'a BStr, &'a BStr, S>,
    property_key: &B,
) -> Option<T>
where
    T: TryFrom<&'a str>,
{
    let key = property_key.as_ref().as_bstr();
    if !properties.contains_key(key) {
        return None;
    }

    let property_ascii = properties[key].to_str().unwrap();

    match T::try_from(property_ascii) {
        Ok(value) => Some(value),
        Err(_) => None,
    }
}

#[must_use]
pub fn build_i32_property<B: AsRef<[u8]>, S: std::hash::BuildHasher>(
    properties: &HashMap<&BStr, &BStr, S>,
    property_key: &B,
    default: i32,
) -> i32 {
    let key = property_key.as_ref().as_bstr();
    if !properties.contains_key(key) {
        return 0;
    }

    let property_ascii = properties[key].to_str().unwrap();

    match property_ascii.parse::<i32>() {
        Ok(value) => value,
        Err(_) => default,
    }
}

#[must_use]
pub fn build_color_property<B: AsRef<[u8]>, S: std::hash::BuildHasher>(
    properties: &HashMap<&BStr, &BStr, S>,
    property_key: &B,
    default: VB6Color,
) -> VB6Color {
    let key = property_key.as_ref().as_bstr();
    if !properties.contains_key(key) {
        return default;
    }

    let property_ascii = properties[key].to_str().unwrap();

    match VB6Color::from_hex(property_ascii) {
        Ok(color) => color,
        Err(_) => default,
    }
}

#[must_use]
pub fn build_bool_property<B: AsRef<[u8]>, S: std::hash::BuildHasher>(
    properties: &HashMap<&BStr, &BStr, S>,
    property_key: &B,
    default: bool,
) -> bool {
    let key = property_key.as_ref().as_bstr();
    if !properties.contains_key(key) {
        return default;
    }

    let property_ascii = properties[key].to_str().unwrap();

    match property_ascii.as_bytes() {
        b"0" => false,
        b"1" | b"-1" => true,
        _ => default,
    }
}

#[cfg(test)]
mod tests {

    use super::*;

    #[test]
    fn begin_property_valid() {
        let source = b"BeginProperty Font\r
                        Name = \"Arial\"\r
                        Size = 8.25\r
                        Charset = 0\r
                        Weight = 400\r
                        Underline = 0 'False\r
                        Italic = 0 'False\r
                        Strikethrough = 0 'False\r
                    EndProperty\r\n";

        let mut input = VB6Stream::new("", source);
        let result = property_group_parse.parse_next(&mut input);

        assert!(result.is_ok());

        let result = result.unwrap();
        assert_eq!(result.name, "Font");
        assert_eq!(result.properties.len(), 7);
    }

    #[test]
    fn mdi_main_frm() {
        use crate::parsers::form::VB6FormFile;

        let input = include_bytes!("../../tests/data/omelette-vb6/Forms/mdiMain.frm");

        let _result = VB6FormFile::parse("mdiMain.frm".to_owned(), input).unwrap();
    }

    #[test]
    fn nested_property_group() {
        use crate::language::VB6ControlKind;
        use crate::parsers::form::VB6FormFile;

        let input = b"VERSION 5.00\r
Object = \"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0\"; \"mscomctl.ocx\"\r
Begin VB.Form Form_Main \r
   BackColor       =   &H00000000&\r
   BorderStyle     =   1  'Fixed Single\r
   Caption         =   \"Audiostation\"\r
   ClientHeight    =   10005\r
   ClientLeft      =   4695\r
   ClientTop       =   1275\r
   ClientWidth     =   12960\r
   BeginProperty Font \r
      Name            =   \"Verdana\"\r
      Size            =   8.25\r
      Charset         =   0\r
      Weight          =   400\r
      Underline       =   0   'False\r
      Italic          =   0   'False\r
      Strikethrough   =   0   'False\r
   EndProperty\r
   Icon            =   \"Form_Main.frx\":0000\r
   LinkTopic       =   \"Form1\"\r
   MaxButton       =   0   'False\r
   OLEDropMode     =   1  'Manual\r
   ScaleHeight     =   10005\r
   ScaleWidth      =   12960\r
   StartUpPosition =   2  'CenterScreen\r
   Begin MSComctlLib.ImageList Imagelist_CDDisplay \r
      Left            =   12000\r
      Top             =   120\r
      _ExtentX        =   1005\r
      _ExtentY        =   1005\r
      BackColor       =   -2147483643\r
      ImageWidth      =   53\r
      ImageHeight     =   42\r
      MaskColor       =   12632256\r
      _Version        =   393216\r
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} \r
         NumListImages   =   5\r
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":17789\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":1921B\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":1ACAD\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":1C73F\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            Picture         =   \"Form_Main.frx\":1E1D1\r
            Key             =   \"\"\r
         EndProperty\r
      EndProperty\r
   End\r
End\r
Attribute VB_Name = \"Form_Main\"\r
";

        let result = VB6FormFile::parse("form_parse.frm".to_owned(), input.as_bytes()).unwrap();

        assert_eq!(result.objects.len(), 1);
        assert_eq!(result.format_version.major, 5);
        assert_eq!(result.format_version.minor, 0);
        assert_eq!(result.form.name, "Form_Main");
        assert_eq!(
            matches!(result.form.kind, VB6ControlKind::Form { .. }),
            true
        );

        if let VB6ControlKind::Form {
            controls,
            properties,
            menus,
        } = &result.form.kind
        {
            assert_eq!(controls.len(), 1);
            assert_eq!(menus.len(), 0);
            assert_eq!(properties.caption, "Audiostation");
            assert_eq!(controls[0].name, "Imagelist_CDDisplay");
            assert!(matches!(controls[0].kind, VB6ControlKind::Custom { .. }));

            if let VB6ControlKind::Custom {
                properties,
                property_groups,
            } = &controls[0].kind
            {
                assert_eq!(properties.len(), 9);
                assert_eq!(property_groups.len(), 1);

                if let Some(group) = property_groups.get(0) {
                    assert_eq!(group.name, "Images");
                    assert_eq!(group.properties.len(), 6);

                    if let Some(Either::Right(image1)) =
                        group.properties.get(BStr::new("ListImage1"))
                    {
                        assert_eq!(image1.name, BStr::new("ListImage1"));
                        assert_eq!(image1.properties.len(), 2);
                    } else {
                        panic!("Expected nested ListImage1");
                    }

                    if let Some(Either::Right(image2)) =
                        group.properties.get(BStr::new("ListImage2"))
                    {
                        assert_eq!(image2.name, BStr::new("ListImage2"));
                        assert_eq!(image2.properties.len(), 2);
                    } else {
                        panic!("Expected nested ListImage2");
                    }

                    if let Some(Either::Right(image3)) =
                        group.properties.get(BStr::new("ListImage3"))
                    {
                        assert_eq!(image3.name, BStr::new("ListImage3"));
                        assert_eq!(image3.properties.len(), 2);
                    } else {
                        panic!("Expected nested ListImage3");
                    }
                } else {
                    panic!("Expected property group");
                }
            } else {
                panic!("Expected custom control");
            }
        } else {
            panic!("Expected form kind");
        }
    }

    #[test]
    fn parse_english_code_non_english_text() {
        use crate::errors::VB6ErrorKind;
        use crate::parsers::form::VB6FormFile;

        let input = "VERSION 5.00\r
Object = \"{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0\"; \"msscript.ocx\"\r
Object = \"{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0\"; \"COMDLG32.OCX\"\r
Object = \"{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0\"; \"Imagex.ocx\"\r
Begin VB.Form FormMainMode \r
   BorderStyle     =   1  '��u�T�w\r
   Caption         =   \"UnlightVBE-QS Origin\"\r
   ClientHeight    =   11100\r
   ClientLeft      =   45\r
   ClientTop       =   375\r
   ClientWidth     =   20400\r
   BeginProperty Font \r
      Name            =   \"�L�n������\"\r
      Size            =   12\r
      Charset         =   136\r
      Weight          =   400\r
      Underline       =   0   'False\r
      Italic          =   0   'False\r
      Strikethrough   =   0   'False\r
   EndProperty\r
   Icon            =   \"FormMainMode.frx\":0000\r
   LinkTopic       =   \"Form3\"\r
   MaxButton       =   0   'False\r
   ScaleHeight     =   11100\r
   ScaleWidth      =   20400\r
   StartUpPosition =   2  '�ù�����\r
   Tag             =   \"UnlightVBE-QS Origin\"\r
   Begin VB.CommandButton �v�l�]�w \r
         Caption         =   \"�v�l�]�w\"\r
         BeginProperty Font \r
            Name            =   \"�L�n������\"\r
            Size            =   8.25\r
            Charset         =   136\r
            Weight          =   400\r
            Underline       =   0   'False\r
            Italic          =   0   'False\r
            Strikethrough   =   0   'False\r
         EndProperty\r
         Height          =   405\r
         Left            =   8760\r
         TabIndex        =   124\r
         Top             =   9360\r
         Width           =   975\r
   End\r
End\r
Attribute VB_Name = \"FormMainMode\"\r
";

        let result = VB6FormFile::parse("form_parse.frm".to_owned(), input.as_bytes());

        assert_eq!(
            result.err().unwrap().kind,
            VB6ErrorKind::LikelyNonEnglishCharacterSet
        );
    }

    #[test]
    fn parse_indented_menu_valid() {
        use crate::language::VB_WINDOW_BACKGROUND;

        let input = b"VERSION 5.00\r
    Begin VB.Form frmExampleForm\r
        BackColor       =   &H80000005&\r
        Caption         =   \"example form\"\r
        ClientHeight    =   6210\r
        ClientLeft      =   60\r
        ClientTop       =   645\r
        ClientWidth     =   9900\r
        BeginProperty Font\r
            Name            =   \"Arial\"\r
            Size            =   8.25\r
            Charset         =   0\r
            Weight          =   400\r
            Underline       =   0   'False\r
            Italic          =   0   'False\r
            Strikethrough   =   0   'False\r
        EndProperty\r
        LinkTopic       =   \"Form1\"\r
        ScaleHeight     =   414\r
        ScaleMode       =   3  'Pixel\r
        ScaleWidth      =   660\r
        StartUpPosition =   2  'CenterScreen\r
        Begin VB.Menu mnuFile\r
            Caption         =   \"&File\"\r
            Begin VB.Menu mnuOpenImage\r
                Caption         =   \"&Open image\"\r
           End\r
        End\r
    End\r
    Attribute VB_Name = \"frmExampleForm\"\r
    ";

        let result = VB6FormFile::parse("form_parse.frm".to_owned(), &mut input.as_ref());

        assert!(result.is_ok());

        let result = result.unwrap();

        assert_eq!(result.format_version.major, 5);
        assert_eq!(result.format_version.minor, 0);

        assert_eq!(result.form.name, "frmExampleForm");
        assert_eq!(result.form.tag, "");
        assert_eq!(result.form.index, 0);

        if let VB6ControlKind::Form {
            controls,
            properties,
            menus,
        } = &result.form.kind
        {
            assert_eq!(controls.len(), 0);
            assert_eq!(menus.len(), 1);
            assert_eq!(properties.caption, "example form");
            assert_eq!(properties.back_color, VB_WINDOW_BACKGROUND);
            assert_eq!(
                menus,
                &vec![VB6MenuControl {
                    name: BStr::new("mnuFile"),
                    tag: BStr::new(""),
                    index: 0,
                    properties: MenuProperties {
                        caption: BStr::new("&File"),
                        ..Default::default()
                    },
                    sub_menus: vec![VB6MenuControl {
                        name: BStr::new("mnuOpenImage"),
                        tag: BStr::new(""),
                        index: 0,
                        properties: MenuProperties {
                            caption: BStr::new("&Open image"),
                            ..Default::default()
                        },
                        sub_menus: vec![],
                    }]
                }]
            );
        } else {
            panic!("Expected form kind");
        }
    }
}
