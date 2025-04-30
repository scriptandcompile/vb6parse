use std::fmt::Debug;
use std::vec::Vec;
use std::{collections::HashMap, path::Path};

use bstr::{BStr, BString, ByteSlice};
use either::Either;
use serde::Serialize;
use uuid::Uuid;
use winnow::{
    ascii::{line_ending, space0, space1},
    combinator::{alt, opt},
    error::{ErrMode, ParserError},
    token::{literal, take_till, take_until},
    Parser,
};

use crate::{
    errors::{VB6Error, VB6ErrorKind},
    language::{VB6Control, VB6ControlKind, VB6MenuControl, VB6Token},
    parsers::{
        header::{
            attributes_parse, key_resource_offset_line_parse, object_parse, version_parse,
            HeaderKind, VB6FileAttributes, VB6FileFormatVersion,
        },
        Properties, VB6ObjectReference, VB6Stream,
    },
    vb6::{keyword_parse, line_comment_parse, string_parse, vb6_parse, VB6Result},
};

/// Represents a VB6 Form file.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct VB6FormFile<'a> {
    pub form: VB6Control,
    pub objects: Vec<VB6ObjectReference<'a>>,
    pub format_version: VB6FileFormatVersion,
    pub attributes: VB6FileAttributes<'a>,
    pub tokens: Vec<VB6Token<'a>>,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
struct VB6FullyQualifiedName {
    pub namespace: BString,
    pub kind: BString,
    pub name: BString,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6PropertyGroup {
    pub name: BString,
    pub guid: Option<Uuid>,
    pub properties: HashMap<BString, Either<BString, VB6PropertyGroup>>,
}

impl Serialize for VB6PropertyGroup {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("VB6PropertyGroup", 3)?;

        state.serialize_field("name", &self.name)?;

        if let Some(guid) = &self.guid {
            state.serialize_field("guid", &guid.to_string())?;
        } else {
            state.serialize_field("guid", &"None")?;
        }

        state.serialize_field("properties", &self.properties)?;

        state.end()
    }
}

pub fn resource_file_resolver(file_path: &str, offset: usize) -> Result<Vec<u8>, std::io::Error> {
    // VB6 FRX files are resource files that contain binary data for controls, forms, and other UI elements.
    // They are typically used in conjunction with VB6 FRM files.
    // The overall format of a VB6 FRX file is not well documented, but it generally consists of a
    // header per record, followed by the binary data for the control or form. There
    // is no overall header for the FRX file itself, and the records are not necessarily in any
    // particular order.
    //
    // The records can be of variable length, so we cannot just read a fixed number of bytes
    // from the file. Instead, we need to parse the file record by record, looking for the specific
    // records that we are interested in from the frm offset.

    // load the bytes from the frx file.
    let buffer = match std::fs::read(file_path) {
        Ok(bytes) => bytes,
        Err(err) => {
            return Err(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                format!("Failed to read resource file {file_path}: {err}"),
            ));
        }
    };

    // Check if the offset is within the bounds of the file.
    if offset >= buffer.len() {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            format!("Offset is out of bounds for resource file {file_path}: {offset}"),
        ));
    }

    let binary_blob_signature = buffer[offset + 4..offset + 8].to_vec();
    if buffer.len() >= 12 && binary_blob_signature.as_slice() == b"lt\0\0" {
        // this is almost certainly a 12 byte header (0-12) where the first 4 bytes
        // is the offset of the record from the end of the signature (exclusive).
        // the next 4 bytes is the magic signature b"lt\0\0".
        // The next four bytes after the 12 byte record heading should be
        // the size of the record from the start of the record buffer.
        // which should be 8 less than the record size from the start of the header.

        let buffer_size_1 =
            u32::from_le_bytes(buffer[offset..(offset + 4)].try_into().unwrap()) as usize;
        // the next 4 bytes after the 12 byte record heading should be
        // the size of the record from the start of the record buffer.
        // which should be 8 less than the record size from the start of the header.
        let buffer_size_2 =
            u32::from_le_bytes(buffer[(offset + 8)..(offset + 12)].try_into().unwrap()) as usize;

        if buffer_size_1 == 8 && buffer_size_2 == 0 {
            // This is a special case where the record is empty.
            // We can just return an empty vector.
            // This usually is the case when someone adds an icon to the form
            // then later removes it.

            return Ok(vec![]);
        }

        // we subtract 8 since the offset is zero index based.
        if buffer_size_2 != buffer_size_1 - 8 {
            return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    format!(
                        "Record size from start of record buffer does not match record size from header: {} != {}. This likely indicates a corrupted resource file.",
                        // We subtract 8 since the record header is 8 bytes (4 for initial size, 4 for magic signature).
                        // the next four bytes is the confirmation buffer size (which brings the total header size to 12).
                        buffer_size_2, buffer_size_1 - 8
                    ),
                ));
        }

        let header_size = 12;
        // The record start is the header size element offset + the header size element length.
        let record_start = offset + header_size;
        let record_end = record_start + buffer_size_2;

        if record_end > buffer.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("Record end is out of bounds for resource file {file_path}: {record_end}"),
            ));
        }

        // Read the record data.
        let record_data = &buffer[record_start..record_end];

        // Return the record data as a vector of bytes.
        return Ok(record_data.to_vec());
    }

    if buffer[offset] == 0xFF {
        // If the first byte of the record is 0xFF, then the record is a 16-bit record.

        // it's a bit excessive to lay out the record size offset/length/end, but it makes it easier to read.
        let header_size_element_offset = offset + 1;
        let header_size_element_length = 2usize;
        let header_size_element_end = header_size_element_offset + header_size_element_length;

        if header_size_element_end > buffer.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!(
                    "Header size element end is out of bounds for resource file {file_path}: {header_size_element_end}"
                ),
            ));
        }

        let header_size_element_bytes = buffer[header_size_element_offset..header_size_element_end]
            .try_into()
            .unwrap();

        let mut record_size = u16::from_le_bytes(header_size_element_bytes) as usize;

        // Unfortunately, vb6 has this goofy way of handling small resource files where if you
        // only have a single short record, it will not include the last byte in the record.
        // This almost only ever happens with string resources.
        // It will indicate that the record is, say, 56 bytes long, when in fact, it's only
        // 55. This usually means that instead of ending on a \r\n, it will end on a \r.
        // This is a bit of a hack, but we need to check if the record size is greater than the
        // buffer length, and if so, we need to subtract 1 from the record size.
        //
        // This is an off by one error in the IDE and almost always results in a string resource
        // that is missing the last byte.
        if header_size_element_offset + record_size > buffer.len() {
            record_size -= 1;
        }

        let record_offset = header_size_element_offset + header_size_element_length;

        let (record_start, record_end) = (record_offset, record_offset + record_size);

        if record_start > buffer.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!(
                    "Record start is out of bounds for resource file {file_path}: {record_start}"
                ),
            ));
        }

        if record_end > buffer.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("Record end is out of bounds for resource file {file_path}: {record_end}"),
            ));
        }

        let record_data = &buffer[record_start..record_end];

        return Ok(record_data.to_vec());
    }

    // List items are a bit special since we need to know how many items there are in the list.
    // This means we can't just remove the header and return the rest of the buffer like we do with
    // the other records.
    let list_signature = buffer[offset + 2..offset + 4].to_vec();
    if buffer.len() >= 12 && (list_signature == [0x03, 0x00] || list_signature == [0x07, 0x00]) {
        // looks like we have a list items record.

        // index 0, 1 = number of list items.
        // index 2, 3 = [0x03, 0x00] || [0x07, 0x00] = list magic indicator.
        //
        // repeats for each list item:
        //      16 bit size of the next list item.
        //      list item without null terminator.

        let list_item_count =
            u16::from_le_bytes(buffer[offset..offset + 2].try_into().unwrap()) as usize;

        // we are going to read the header and the list items into a single vector.
        let header_size = 4;
        let mut record_offset = offset + header_size;
        let list_item_header_size = 2;
        for _ in 0..list_item_count {
            if record_offset > buffer.len() {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    format!(
                        "Record offset of list is out of bounds for resource file {file_path}: {record_offset}"
                    ),
                ));
            }

            let record_end = record_offset + list_item_header_size;

            if record_end > buffer.len() {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    format!(
                        "Record end of list is out of bounds for resource file {file_path}: {record_end}"
                    ),
                ));
            }

            let list_item_size =
                u16::from_le_bytes(buffer[record_offset..record_end].try_into().unwrap()) as usize;

            // If we were trying to pull out a list from this, this is where we would do it.
            //
            // let record_item_start = record_offset + list_item_header_size;
            // let list_item = &buffer[record_item_start..record_item_start + list_item_size];

            record_offset += list_item_header_size + list_item_size;
        }

        if record_offset > buffer.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!(
                    "Record end of list is out of bounds for resource file {file_path}: {record_offset}"
                ),
            ));
        }

        return Ok(buffer[offset..record_offset].to_vec());
    }

    // If the first byte of the record is not 0xFF, then the record is likely a 4 byte header record.
    // check if we have any null bytes in the first 4 bytes of the record.
    // this probably indicates that the record is a 4 byte header record.
    if buffer.len() >= 12 && buffer[(offset)..(offset + 4)].contains(&0u8) {
        // this looks like a 4 byte header (0-4) where the 4 bytes
        // is the size of the record from the start of the record + header.

        // often, this is what is used when we have a larger chunk of text data.

        let header_size = 4;
        let record_size =
            u32::from_le_bytes(buffer[(offset)..(offset + 4)].try_into().unwrap()) as usize;

        let record_start = offset + header_size;
        let record_end = record_start + record_size;

        if record_end > buffer.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("Record end is out of bounds for resource file {file_path}: {record_end}"),
            ));
        }

        let record_data = &buffer[record_start..record_end];

        return Ok(record_data.to_vec());
    }

    // If the first byte of the record is not 0xFF, then the record is likely an 8-bit record.
    // It's a bit excessive to lay out the record size offset/ length/end, but it makes it easier to read.
    let header_size = 1; // 1 byte header size element.
    let record_size = buffer[offset] as usize;
    let record_start = offset + header_size;

    if record_start > buffer.len() {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            format!("Record start is out of bounds for resource file {file_path}: {record_start}"),
        ));
    }

    let record_end = match buffer.len() {
        // If the record size is greater than the buffer length, then we need to subtract 1 from the record size.
        // This is a bit of a hack, but we need to check if the record size is greater than the buffer length,
        // and if so, we need to subtract 1 from the record size.
        // This is an off by one error in the IDE and almost always results in a string resource
        // that is missing the last byte.
        _ if record_size >= buffer.len() => record_start + record_size - 1,
        _ => record_start + record_size,
    };

    if record_end > buffer.len() {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            format!("Record end is out of bounds for resource file {file_path}: {record_end}"),
        ));
    }

    // Read the record data.
    let record_data = &buffer[record_start..record_end];

    // Return the record data as a vector of bytes.
    Ok(record_data.to_vec())
}

#[must_use]
pub fn list_resolver(buffer: &[u8]) -> Vec<BString> {
    let mut list_items = vec![];

    if buffer.len() < 2 {
        return list_items;
    }

    let item_count_buffer: [u8; 2] = match buffer[0..2].try_into() {
        Ok(bytes) => bytes,
        Err(_) => return list_items,
    };

    let list_item_count = u16::from_le_bytes(item_count_buffer) as usize;

    // we are going to read the header and the list items into a single vector.
    let header_size = 4;
    let mut record_offset = header_size;
    let list_item_header_size = 2;
    for _ in 0..list_item_count {
        if record_offset > buffer.len() {
            return list_items;
        }

        let record_end = record_offset + list_item_header_size;

        if record_end > buffer.len() {
            return list_items;
        }

        let Ok(list_item_buffer) = buffer[record_offset..record_end].try_into() else {
            return list_items;
        };

        let list_item_size = u16::from_le_bytes(list_item_buffer) as usize;

        let record_item_start = record_offset + list_item_header_size;
        let list_item = &buffer[record_item_start..record_item_start + list_item_size];

        list_items.push(list_item.as_bstr().to_owned());

        record_offset += list_item_header_size + list_item_size;
    }

    list_items
}

impl<'a> VB6FormFile<'a> {
    /// Parses a VB6 form file from a byte slice using the selected `resource_file_resolver`.
    ///
    /// # Arguments
    ///
    /// * `file_name` The name of the file being parsed.
    /// * `input` The byte slice to parse.
    /// * `resource_resolver` A function that resolves resource files based on their name and offset.
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
    /// use vb6parse::parsers::resource_file_resolver;
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
    /// let result = VB6FormFile::parse_with_resolver("form_parse.frm", &mut input.as_ref(), resource_file_resolver);
    ///
    ///
    /// assert!(result.is_ok());
    /// ```
    pub fn parse_with_resolver(
        form_path: &str,
        input: &'a [u8],
        resource_resolver: impl Fn(&str, usize) -> Result<Vec<u8>, std::io::Error>,
    ) -> Result<Self, VB6Error> {
        let mut input = VB6Stream::new(form_path, input);

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
        }

        let form = match control_parse(&resource_resolver).parse_next(&mut input) {
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

    pub fn parse(form_path: &str, input: &'a [u8]) -> Result<Self, VB6Error> {
        VB6FormFile::parse_with_resolver(form_path, input, resource_file_resolver)
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

struct ControlBlock<'a> {
    pub fully_qualified_name: VB6FullyQualifiedName,
    pub sub_controls: Vec<VB6Control>,
    pub sub_menus: Vec<VB6Control>,
    pub property_groups: Vec<VB6PropertyGroup>,
    pub properties: Properties<'a>,
}

fn control_parse<'a>(
    resource_resolver: impl Fn(&str, usize) -> Result<Vec<u8>, std::io::Error>,
) -> impl FnMut(&mut VB6Stream<'a>) -> VB6Result<VB6Control> {
    move |input: &mut VB6Stream<'a>| -> VB6Result<VB6Control> {
        let file_path = input.file_path.clone();
        let parent_folder = Path::new(&file_path).parent().unwrap_or(Path::new(""));

        let mut fully_qualified_name = property_parse.parse_next(input)?;

        let mut current_control_block = ControlBlock {
            fully_qualified_name,
            sub_controls: vec![],
            sub_menus: vec![],
            property_groups: vec![],
            properties: Properties::new(),
        };

        let mut control_block_stack: Vec<ControlBlock<'_>> = vec![];

        while !input.is_empty() {
            // Check if we are at the end of the control block.
            if (space0, keyword_parse("END"), space0, line_ending)
                .parse_next(input)
                .is_ok()
            {
                // We are at the end of the control block, so we need to build the control.
                match build_control(current_control_block) {
                    Ok(control) => {
                        // We have a valid control, so we need to check if we have a parent control block.
                        match control_block_stack.pop() {
                            // If we have a parent control block, we need to add the control to it.
                            Some(mut parent_control_block) => {
                                if control.kind.is_menu() {
                                    parent_control_block.sub_menus.push(control);
                                } else {
                                    parent_control_block.sub_controls.push(control);
                                }

                                // the current control block is now the parent control block.
                                current_control_block = parent_control_block;
                            }
                            // If we don't have a parent control block, we are done parsing the control and its sub-controls.
                            None => return Ok(control), // We are done parsing the control.
                        }
                    }
                    // If we had and error while trying to build the control, we return the error.
                    Err(err) => return Err(ErrMode::Cut(err)),
                }

                continue;
            }

            // Check if we have a nested control.
            if (space0, keyword_parse("BEGIN"), space1)
                .parse_next(input)
                .is_ok()
            {
                // push the current control block onto the stack and create a new one for the nested control.
                control_block_stack.push(current_control_block);

                // Parse the next control's fully qualified name.
                fully_qualified_name = property_parse.parse_next(input)?;

                // Create a new control block for the nested control.
                current_control_block = ControlBlock {
                    fully_qualified_name,
                    sub_controls: vec![],
                    sub_menus: vec![],
                    property_groups: vec![],
                    properties: Properties::new(),
                };

                continue;
            }

            // Property groups start with "BeginProperty" and end with "EndProperty".
            if let Ok(property_group) = property_group_parse.parse_next(input) {
                // We have a property group.

                current_control_block.property_groups.push(property_group);
                continue;
            }

            // We have a resource file reference property.
            if let Ok((name, resource_file, offset)) =
                key_resource_offset_line_parse.parse_next(input)
            {
                let resource_path = Path::join(parent_folder, resource_file.to_string())
                    .to_string_lossy()
                    .to_string();

                let resource = match resource_resolver(&resource_path, offset as usize) {
                    Ok(res) => res,
                    Err(err) => {
                        return Err(ErrMode::Cut(VB6ErrorKind::ResourceFile(err)));
                    }
                };

                current_control_block
                    .properties
                    .insert_resource(name, resource);

                continue;
            }

            // We have a property key-value pair.
            let (name, value) = property_key_value_parse.parse_next(input)?;

            current_control_block.properties.insert(name, value);
        }

        Err(ParserError::assert(input, "Unknown control kind"))
    }
}

fn property_key_value_parse<'a>(input: &mut VB6Stream<'a>) -> VB6Result<(&'a BStr, &'a [u8])> {
    space0.parse_next(input)?;

    let name = take_till(1.., (b' ', b'\t', b'=')).parse_next(input)?;

    space0.parse_next(input)?;

    "=".parse_next(input)?;

    space0.parse_next(input)?;

    let value =
        alt((string_parse, take_till(1.., (' ', '\t', '\'', '\r', '\n')))).parse_next(input)?;

    space0.parse_next(input)?;

    opt(line_comment_parse).parse_next(input)?;

    line_ending.parse_next(input)?;

    Ok((name, value.as_bytes()))
}

fn property_group_parse(input: &mut VB6Stream<'_>) -> VB6Result<VB6PropertyGroup> {
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
        name: property_group_name.into(),
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
                nested_property_group.name.clone(),
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
                .insert(name.into(), Either::Left("".into()));

            continue;
        }

        space0.parse_next(input)?;

        let name = take_till(1.., (b'\t', b' ', b'=')).parse_next(input)?;

        (space0, "=", space0).parse_next(input)?;

        let value =
            alt((string_parse, take_till(1.., (' ', '\t', '\'', '\r', '\n')))).parse_next(input)?;

        property_group
            .properties
            .insert(name.into(), Either::Left(value.into()));

        (space0, opt(line_comment_parse), line_ending).parse_next(input)?;
    }

    Ok(property_group)
}

fn property_parse(input: &mut VB6Stream<'_>) -> VB6Result<VB6FullyQualifiedName> {
    let Ok(namespace) = take_until::<_, _, VB6Error>(0.., ".").parse_next(input) else {
        return Err(ErrMode::Cut(VB6ErrorKind::NoNamespaceAfterBegin));
    };

    if literal::<&str, _, VB6Error>(".").parse_next(input).is_err() {
        return Err(ErrMode::Cut(VB6ErrorKind::NoDotAfterNamespace));
    }

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
        namespace: namespace.into(),
        kind: kind.into(),
        name: name.into(),
    })
}

fn build_control(control_block: ControlBlock<'_>) -> Result<VB6Control, VB6ErrorKind> {
    let tag = match control_block.properties.get(b"Tag".into()) {
        Some(text) => text.into(),
        None => b"".into(),
    };

    if control_block.fully_qualified_name.namespace != "VB" {
        let custom_control = VB6Control {
            name: control_block.fully_qualified_name.name,
            tag,
            index: 0,
            kind: VB6ControlKind::Custom {
                properties: control_block.properties.into(),
                property_groups: control_block.property_groups,
            },
        };

        return Ok(custom_control);
    }

    let kind = match control_block.fully_qualified_name.kind.as_bytes() {
        b"Form" => {
            let mut converted_menus = vec![];

            for menu in control_block.sub_menus {
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
                controls: control_block.sub_controls,
                properties: control_block.properties.into(),
                menus: converted_menus,
            }
        }
        b"MDIForm" => {
            let mut converted_menus = vec![];

            for menu in control_block.sub_menus {
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
                controls: control_block.sub_controls,
                properties: control_block.properties.into(),
                menus: converted_menus,
            }
        }
        b"Menu" => {
            let menu_properties = control_block.properties.into();
            let mut converted_menus = vec![];

            for menu in control_block.sub_menus {
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
        b"Frame" => VB6ControlKind::Frame {
            controls: control_block.sub_controls,
            properties: control_block.properties.into(),
        },
        b"CheckBox" => VB6ControlKind::CheckBox {
            properties: control_block.properties.into(),
        },
        b"ComboBox" => VB6ControlKind::ComboBox {
            properties: control_block.properties.into(),
        },
        b"CommandButton" => VB6ControlKind::CommandButton {
            properties: control_block.properties.into(),
        },
        b"Data" => VB6ControlKind::Data {
            properties: control_block.properties.into(),
        },
        b"DirListBox" => VB6ControlKind::DirListBox {
            properties: control_block.properties.into(),
        },
        b"DriveListBox" => VB6ControlKind::DriveListBox {
            properties: control_block.properties.into(),
        },
        b"FileListBox" => VB6ControlKind::FileListBox {
            properties: control_block.properties.into(),
        },
        b"Image" => VB6ControlKind::Image {
            properties: control_block.properties.into(),
        },
        b"Label" => VB6ControlKind::Label {
            properties: control_block.properties.into(),
        },
        b"Line" => VB6ControlKind::Line {
            properties: control_block.properties.into(),
        },
        b"ListBox" => VB6ControlKind::ListBox {
            properties: control_block.properties.into(),
        },
        b"OLE" => VB6ControlKind::Ole {
            properties: control_block.properties.into(),
        },
        b"OptionButton" => VB6ControlKind::OptionButton {
            properties: control_block.properties.into(),
        },
        b"PictureBox" => VB6ControlKind::PictureBox {
            properties: control_block.properties.into(),
        },
        b"HScrollBar" => VB6ControlKind::HScrollBar {
            properties: control_block.properties.into(),
        },
        b"VScrollBar" => VB6ControlKind::VScrollBar {
            properties: control_block.properties.into(),
        },
        b"Shape" => VB6ControlKind::Shape {
            properties: control_block.properties.into(),
        },
        b"TextBox" => VB6ControlKind::TextBox {
            properties: control_block.properties.into(),
        },
        b"Timer" => VB6ControlKind::Timer {
            properties: control_block.properties.into(),
        },
        _ => {
            return Err(VB6ErrorKind::UnknownControlKind);
        }
    };

    let parent_control = VB6Control {
        name: control_block.fully_qualified_name.name,
        tag: tag.clone(),
        index: 0,
        kind,
    };

    Ok(parent_control)
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
    fn nested_property_group() {
        //use crate::language::VB6ControlKind;
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
            _Version        =   9\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            _Version        =   1\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            _Version        =   1\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            _Version        =   5\r
            Key             =   \"\"\r
         EndProperty\r
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} \r
            _Version        =   1\r
            Key             =   \"\"\r
         EndProperty\r
      EndProperty\r
   End\r
End\r
Attribute VB_Name = \"Form_Main\"\r
";

        let result = VB6FormFile::parse("form_parse.frm", input.as_bytes());

        assert!(result.is_ok());

        let result = result.unwrap();

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

        let result = VB6FormFile::parse("form_parse.frm", input.as_bytes());

        assert!(matches!(
            result.err().unwrap().kind,
            VB6ErrorKind::LikelyNonEnglishCharacterSet
        ));
    }

    #[test]
    fn parse_indented_menu_valid() {
        use crate::language::MenuProperties;
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

        let result = VB6FormFile::parse("form_parse.frm", &mut input.as_ref());

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
                    name: "mnuFile".into(),
                    tag: "".into(),
                    index: 0,
                    properties: MenuProperties {
                        caption: "&File".into(),
                        ..Default::default()
                    },
                    sub_menus: vec![VB6MenuControl {
                        name: "mnuOpenImage".into(),
                        tag: "".into(),
                        index: 0,
                        properties: MenuProperties {
                            caption: "&Open image".into(),
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
