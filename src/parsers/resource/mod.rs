/// Resolves a resource file from the given file path and offset.
///
/// # Arguments
///
/// * `file_path` - The path to the resource file.
/// * `offset` - The offset of the resource in the file.
///
/// # Returns
///
/// A result containing the resource data as a vector of bytes or an error.
///
/// # Errors
///
/// An error will be returned if the resource file cannot be read or if the offset is out of bounds.
///
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

        let Ok(size_buffer) = buffer[offset..(offset + 4)].try_into() else {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("Failed to read size buffer for resource file {file_path}"),
            ));
        };
        let buffer_size_1 = u32::from_le_bytes(size_buffer) as usize;

        // the next 4 bytes after the 12 byte record heading should be
        // the size of the record from the start of the record buffer.
        // which should be 8 less than the record size from the start of the header.

        let Ok(secondary_buffer_size) = buffer[(offset + 8)..(offset + 12)].try_into() else {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("Failed to read secondary buffer size for resource file {file_path}"),
            ));
        };
        let buffer_size_2 = u32::from_le_bytes(secondary_buffer_size) as usize;

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

        let Ok(header_size_element_bytes) =
            buffer[header_size_element_offset..header_size_element_end].try_into()
        else {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("Failed to read header size element bytes for resource file {file_path}"),
            ));
        };

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

        let Ok(list_item_buffer) = buffer[offset..(offset + 2)].try_into() else {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("Failed to read list item buffer for resource file {file_path}"),
            ));
        };

        let list_item_count = u16::from_le_bytes(list_item_buffer) as usize;

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

            let Ok(list_item_size_buffer) = buffer[record_offset..record_end].try_into() else {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    format!("Failed to read list item size buffer for resource file {file_path}"),
                ));
            };

            let list_item_size = u16::from_le_bytes(list_item_size_buffer) as usize;

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
        let Ok(record_size_buffer) = buffer[offset..(offset + 4)].try_into() else {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("Failed to read record size buffer for resource file {file_path}"),
            ));
        };
        let record_size = u32::from_le_bytes(record_size_buffer) as usize;
        let header_size = 4;
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
pub fn list_resolver(buffer: &[u8]) -> Vec<String> {
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

        list_items.push(String::from_utf8_lossy(list_item).to_string());

        record_offset += list_item_header_size + list_item_size;
    }

    list_items
}
