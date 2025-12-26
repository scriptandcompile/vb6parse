//! Module for resolving resource files in VB6 projects.
//!
//! VB6 resource files (FRX files) contain binary data for controls, forms, and other UI elements.
//! This module provides functions to read and extract resource data from these files.
//!

use serde::Serialize;
use std::collections::HashMap;
use std::fmt::Display;
use std::path::Path;

use crate::errors::{ErrorDetails, ResourceErrorKind};
use crate::ParseResult;

/// Represents a parsed VB6 Form Resource file (.frx).
///
/// FRX files contain binary data for controls, forms, and other UI elements.
/// They are structured as a sequence of variable-length records without an
/// overall file header. Records are referenced from the associated .frm file
/// by byte offset.
///
/// # Example
///
/// ```no_run
/// use vb6parse::parsers::FormResourceFile;
///
/// let bytes = std::fs::read("path/to/form.frx")?;
/// let result = FormResourceFile::parse("form.frx", bytes);
/// let resource_file = result.unwrap_or_fail();
///
/// // Or use from_file for convenience
/// let result = FormResourceFile::from_file("path/to/form.frx")?;
/// let resource_file = result.unwrap_or_fail();
///
/// // Access a binary blob (e.g., icon) at offset 0x00
/// if let Some(data) = resource_file.get_binary_blob(0x00) {
///     println!("Icon size: {} bytes", data.len());
/// }
///
/// // Access list items (e.g., combo box contents) at offset 0x100
/// if let Some(items) = resource_file.get_list_items(0x100) {
///     for item in items {
///         println!("List item: {}", item);
///     }
/// }
/// # Ok::<(), std::io::Error>(())
/// ```
#[derive(Debug, Clone, Serialize)]
pub struct FormResourceFile {
    /// Complete buffer of the resource file contents
    ///
    /// Stored for reference and to avoid re-reading the file.
    /// Individual entries reference slices of this buffer.
    #[serde(skip)]
    buffer: Vec<u8>,

    /// File name of the resource file
    file_name: Box<str>,

    /// Parsed resource entries indexed by their byte offset in the file
    ///
    /// Keys are the byte offsets where each resource entry begins.
    /// This allows O(1) lookup when the FRM file references a resource.
    entries: HashMap<usize, ResourceEntry>,
}

impl Display for FormResourceFile {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "FormResourceFile {{ file_name: {:?}, size: {} bytes, entries: {} }}",
            self.file_name,
            self.buffer.len(),
            self.entries.len()
        )
    }
}

/// Represents a single resource entry in a VB6 FRX file.
///
/// Each entry type corresponds to a different binary format used by VB6
/// for storing various kinds of data.
#[derive(Debug, Clone, Serialize, PartialEq)]
pub enum ResourceEntry {
    /// Binary blob with 12-byte header (signature: "lt\0\0")
    ///
    /// Used for: Images, icons, cursor files, OLE objects
    ///
    /// Header format:
    /// - Bytes 0-3: Size from end of signature (u32 LE)
    /// - Bytes 4-7: Magic signature "lt\0\0"
    /// - Bytes 8-11: Size from start of data (u32 LE, should be bytes 0-3 minus 8)
    /// - Bytes 12+: Binary data
    Record12ByteHeader {
        /// The binary data (excluding header)
        data: Vec<u8>,
    },

    /// 16-bit record with 0xFF header
    ///
    /// Used for: Small to medium text/string data
    ///
    /// Header format:
    /// - Byte 0: 0xFF (marker)
    /// - Bytes 1-2: Size of data (u16 LE)
    /// - Bytes 3+: Data
    ///
    /// Note: VB6 IDE has an off-by-one bug where some records
    /// are marked as N bytes but are actually N-1 bytes.
    Record3ByteHeader {
        /// The record data (excluding header)
        data: Vec<u8>,
    },

    /// List items record with magic signature [0x03, 0x00] or [0x07, 0x00]
    ///
    /// Used for: ComboBox/ListBox list items
    ///
    /// Format:
    /// - Bytes 0-1: Number of items (u16 LE)
    /// - Bytes 2-3: Magic signature [0x03, 0x00] or [0x07, 0x00]
    /// - For each item:
    ///   - 2 bytes: Item length (u16 LE)
    ///   - N bytes: Item data (no null terminator)
    ListItems {
        /// Parsed list of strings
        items: Vec<String>,
    },

    /// 4-byte header record
    ///
    /// Used for: Large text/binary data blocks
    ///
    /// Header format:
    /// - Bytes 0-3: Size of data including header (u32 LE)
    /// - Bytes 4+: Raw binary data
    ///
    /// Note: This data may be text (Windows-1252), images (PNG/BMP), or other binary data.
    /// Use helper methods to extract as needed format.
    Record4ByteHeader {
        /// The raw binary data (excluding header)
        data: Vec<u8>,
    },

    /// 8-bit record with single-byte header
    ///
    /// Used for: Small data (< 256 bytes)
    ///
    /// Header format:
    /// - Byte 0: Size of data (u8)
    /// - Bytes 1+: Data
    Record1ByteHeader {
        /// The record data (excluding header)
        data: Vec<u8>,
    },

    /// Empty record (special case)
    ///
    /// Occurs when someone adds an icon/image to a form then removes it.
    /// The IDE leaves behind an empty 12-byte header with:
    /// - Bytes 0-3: 0x00000008
    /// - Bytes 4-7: "lt\0\0"
    /// - Bytes 8-11: 0x00000000
    Empty {
        /// Offset where this empty record was found
        offset: usize,
    },
}

impl ResourceEntry {
    /// Attempts to decode the entry data as Windows-1252 text.
    ///
    /// Works for `Record12ByteHeader` and `Record4ByteHeader` entries.
    ///
    /// # Returns
    ///
    /// `Some(String)` if the data can be decoded as valid Windows-1252,
    /// `None` otherwise.
    #[must_use]
    pub fn as_text(&self) -> Option<String> {
        let bytes = match self {
            ResourceEntry::Record12ByteHeader { data }
            | ResourceEntry::Record4ByteHeader { data } => data.as_slice(),
            _ => return None,
        };

        encoding_rs::WINDOWS_1252
            .decode_without_bom_handling_and_without_replacement(bytes)
            .map(|s| s.to_string())
    }

    /// Returns the raw bytes of the entry data.
    ///
    /// Works for `Record12ByteHeader`, `Record4ByteHeader`, `Record3ByteHeader`, and `Record1ByteHeader` entries.
    ///
    /// # Returns
    ///
    /// `Some(&[u8])` containing the raw data, `None` for `ListItems` or `Empty`.
    #[must_use]
    pub fn as_bytes(&self) -> Option<&[u8]> {
        match self {
            ResourceEntry::Record12ByteHeader { data }
            | ResourceEntry::Record4ByteHeader { data }
            | ResourceEntry::Record3ByteHeader { data }
            | ResourceEntry::Record1ByteHeader { data } => Some(data.as_slice()),
            ResourceEntry::ListItems { .. } | ResourceEntry::Empty { .. } => None,
        }
    }
}

/// Metadata about a resource entry location and type.
///
/// Used internally during parsing to track entry boundaries.
#[derive(Debug, Clone)]
struct ResourceEntryMetadata {
    /// Byte offset where the entry starts
    offset: usize,
    /// Total size including header
    total_size: usize,
    /// Type of entry
    entry_type: ResourceEntryType,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ResourceEntryType {
    Record12ByteHeader,
    Record3ByteHeader,
    ListItems,
    Record4ByteHeader,
    Record1ByteHeader,
    Empty,
}
impl FormResourceFile {
    /// Parses a VB6 Form Resource file from an owned byte vector.
    ///
    /// This method takes ownership of the buffer, avoiding an extra copy.
    ///
    /// # Arguments
    ///
    /// * `buffer` - Byte vector containing the .frx file contents (consumed)
    ///
    /// # Returns
    ///
    /// A `ParseResult` containing the parsed resource file and any non-fatal errors.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use vb6parse::parsers::FormResourceFile;
    ///
    /// let bytes = std::fs::read("tests/data/form.frx")?;
    /// let result = FormResourceFile::parse("form.frx", bytes);
    ///
    /// if result.has_failures() {
    ///     for failure in result.failures() {
    ///         failure.print();
    ///     }
    /// }
    ///
    /// let resource_file = result.unwrap_or_fail();
    /// println!("Parsed {} entries", resource_file.entry_count());
    /// # Ok::<(), std::io::Error>(())
    /// ```
    #[must_use]
    pub fn parse(
        file_name: &str,
        buffer: Vec<u8>,
    ) -> ParseResult<'static, Self, ResourceErrorKind> {
        let mut failures = Vec::new();
        let mut entries = HashMap::new();
        let file_name_box = file_name.to_string().into_boxed_str();

        // 1. Scan through and identify all resource entries
        let entry_offsets = Self::scan_entries(&buffer, &mut failures, &file_name_box);

        // 2. Parse each entry into ResourceEntry enum
        for metadata in entry_offsets {
            match Self::parse_entry(&buffer, &metadata) {
                Ok(entry) => {
                    entries.insert(metadata.offset, entry);
                }
                Err(err) => {
                    // Create ErrorDetails manually - resource files have no source_content
                    // so we use an empty string
                    failures.push(ErrorDetails {
                        source_name: file_name_box.clone(),
                        source_content: "",
                        // The offset should fit in u32 as FRX files are legacy 32-bit format
                        // and limited to 4GB in size.
                        error_offset: u32::try_from(metadata.offset).unwrap_or(0),
                        line_start: 0,
                        line_end: 0,
                        kind: err,
                    });
                }
            }
        }

        // 3. Return FormResourceFile with HashMap of entries
        ParseResult::new(
            Some(FormResourceFile {
                file_name: file_name_box,
                buffer, // Move the Vec, no clone needed
                entries,
            }),
            failures,
        )
    }

    /// Loads and parses a VB6 Form Resource file from a file path.
    ///
    /// This is a convenience method that reads the file and calls `parse()`.
    /// Use `parse()` directly if you already have the file contents in memory.
    ///
    /// # Arguments
    ///
    /// * `file_path` - Path to the .frx file to load and parse
    ///
    /// # Returns
    ///
    /// A `Result` containing the `ParseResult` or an I/O error.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use vb6parse::parsers::FormResourceFile;
    ///
    /// let result = FormResourceFile::from_file("tests/data/form.frx")?;
    /// let resource_file = result.unwrap_or_fail();
    /// # Ok::<(), std::io::Error>(())
    /// ```
    pub fn from_file<P: AsRef<Path>>(
        file_path: P,
    ) -> std::io::Result<ParseResult<'static, Self, ResourceErrorKind>> {
        let path = file_path.as_ref();
        let bytes = std::fs::read(path)?;
        let file_name = path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("unknown.frx");
        Ok(Self::parse(file_name, bytes))
    }

    /// Scans through the buffer to identify all resource entry locations.
    ///
    /// # Arguments
    ///
    /// * `buffer` - The complete .frx file buffer
    /// * `failures` - Collection for non-fatal parse errors
    /// * `file_name_box` - The name of the resource file for error reporting
    ///
    /// # Returns
    ///
    /// A vector of metadata for each discovered entry
    fn scan_entries(
        buffer: &[u8],
        failures: &mut Vec<ErrorDetails<'static, ResourceErrorKind>>,
        file_name: &str,
    ) -> Vec<ResourceEntryMetadata> {
        let mut entries = Vec::new();
        let mut offset = 0;

        while offset < buffer.len() {
            match Self::identify_entry(buffer, offset) {
                Ok(metadata) => {
                    offset = metadata.offset + metadata.total_size;
                    entries.push(metadata);
                }
                Err(err) => {
                    // Create ErrorDetails manually for scan errors
                    failures.push(ErrorDetails {
                        source_name: file_name.into(),
                        source_content: "",
                        // The offset should fit in u32 as FRX files are legacy 32-bit format
                        // and limited to 4GB in size.
                        error_offset: u32::try_from(offset).unwrap_or(0),
                        line_start: 0,
                        line_end: 0,
                        kind: err,
                    });
                    // Skip to next byte and try again
                    offset += 1;
                }
            }
        }

        entries
    }

    /// Identifies the type and size of an entry at the given offset.
    ///
    /// Detects the entry type based on header signatures and calculates
    /// the total size including header. Does not parse the entry data.
    fn identify_entry(
        buffer: &[u8],
        offset: usize,
    ) -> Result<ResourceEntryMetadata, ResourceErrorKind> {
        if offset >= buffer.len() {
            return Err(ResourceErrorKind::OffsetOutOfBounds {
                offset,
                file_length: buffer.len(),
            });
        }

        // Check for Record12ByteHeader (12-byte header with "lt\0\0" signature)
        if offset + 12 <= buffer.len() && &buffer[offset + 4..offset + 8] == b"lt\0\0" {
            let size_buffer_1 = buffer[offset..offset + 4]
                .try_into()
                .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
            let buffer_size_1 = u32::from_le_bytes(size_buffer_1) as usize;

            let size_buffer_2 = buffer[offset + 8..offset + 12]
                .try_into()
                .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
            let buffer_size_2 = u32::from_le_bytes(size_buffer_2) as usize;

            // Check for empty record special case
            if buffer_size_1 == 8 && buffer_size_2 == 0 {
                return Ok(ResourceEntryMetadata {
                    offset,
                    total_size: 12,
                    entry_type: ResourceEntryType::Empty,
                });
            }

            // Regular binary blob
            let total_size = 12 + buffer_size_2;
            return Ok(ResourceEntryMetadata {
                offset,
                total_size,
                entry_type: ResourceEntryType::Record12ByteHeader,
            });
        }

        // Check for Record3ByteHeader (0xFF header)
        if buffer[offset] == 0xFF && offset + 3 <= buffer.len() {
            let size_buffer = buffer[offset + 1..offset + 3]
                .try_into()
                .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
            let mut record_size = u16::from_le_bytes(size_buffer) as usize;

            // Handle VB6 IDE off-by-one bug
            if offset + 3 + record_size > buffer.len() {
                record_size -= 1;
            }

            let total_size = 3 + record_size;
            return Ok(ResourceEntryMetadata {
                offset,
                total_size,
                entry_type: ResourceEntryType::Record3ByteHeader,
            });
        }

        // Check for ListItems (signature [0x03, 0x00] or [0x07, 0x00] at offset+2)
        if offset + 4 <= buffer.len() {
            let signature = &buffer[offset + 2..offset + 4];
            if signature == [0x03, 0x00] || signature == [0x07, 0x00] {
                // Calculate total size by scanning list items
                let count_buffer = buffer[offset..offset + 2]
                    .try_into()
                    .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
                let item_count = u16::from_le_bytes(count_buffer) as usize;

                let mut current_offset = offset + 4;
                for _ in 0..item_count {
                    if current_offset + 2 > buffer.len() {
                        return Err(ResourceErrorKind::CorruptedListItems {
                            offset,
                            details: "Item header out of bounds".to_string(),
                        });
                    }

                    let item_size_buffer = buffer[current_offset..current_offset + 2]
                        .try_into()
                        .map_err(|_| ResourceErrorKind::BufferConversionError {
                            offset: current_offset,
                        })?;
                    let item_size = u16::from_le_bytes(item_size_buffer) as usize;

                    current_offset += 2 + item_size;

                    if current_offset > buffer.len() {
                        return Err(ResourceErrorKind::CorruptedListItems {
                            offset,
                            details: "Item data out of bounds".to_string(),
                        });
                    }
                }

                let total_size = current_offset - offset;
                return Ok(ResourceEntryMetadata {
                    offset,
                    total_size,
                    entry_type: ResourceEntryType::ListItems,
                });
            }
        }

        // Check for Record4ByteHeader (4-byte header with null bytes)
        if offset + 4 <= buffer.len() && buffer[offset..offset + 4].contains(&0u8) {
            let size_buffer = buffer[offset..offset + 4]
                .try_into()
                .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
            let record_size = u32::from_le_bytes(size_buffer) as usize;

            let total_size = 4 + record_size;
            return Ok(ResourceEntryMetadata {
                offset,
                total_size,
                entry_type: ResourceEntryType::Record4ByteHeader,
            });
        }

        // Default: Record1ByteHeader (single-byte header)
        let mut record_size = buffer[offset] as usize;

        // Handle VB6 IDE off-by-one bug
        if offset + 1 + record_size > buffer.len() {
            record_size = record_size.saturating_sub(1);
        }

        let total_size = 1 + record_size;
        Ok(ResourceEntryMetadata {
            offset,
            total_size,
            entry_type: ResourceEntryType::Record1ByteHeader,
        })
    }

    /// Parses a single entry based on its metadata.
    fn parse_entry(
        buffer: &[u8],
        metadata: &ResourceEntryMetadata,
    ) -> Result<ResourceEntry, ResourceErrorKind> {
        match metadata.entry_type {
            ResourceEntryType::Record12ByteHeader => Self::parse_binary_blob(buffer, metadata),
            ResourceEntryType::Record3ByteHeader => Self::parse_16bit_record(buffer, metadata),
            ResourceEntryType::ListItems => Self::parse_list_items(buffer, metadata),
            ResourceEntryType::Record4ByteHeader => Self::parse_text_data(buffer, metadata),
            ResourceEntryType::Record1ByteHeader => Self::parse_8bit_record(buffer, metadata),
            ResourceEntryType::Empty => Ok(ResourceEntry::Empty {
                offset: metadata.offset,
            }),
        }
    }

    /// Parses a binary blob entry (12-byte header with "lt\0\0" signature).
    fn parse_binary_blob(
        buffer: &[u8],
        metadata: &ResourceEntryMetadata,
    ) -> Result<ResourceEntry, ResourceErrorKind> {
        let offset = metadata.offset;

        // Verify we have enough bytes for the header
        if offset + 12 > buffer.len() {
            return Err(ResourceErrorKind::HeaderReadError {
                offset,
                reason: "Not enough bytes for 12-byte header".to_string(),
            });
        }

        // Verify signature
        let signature = &buffer[offset + 4..offset + 8];
        if signature != b"lt\0\0" {
            return Err(ResourceErrorKind::InvalidData {
                offset,
                details: format!("Invalid signature: {signature:?}"),
            });
        }

        // Extract sizes
        let size_buffer_1 = buffer[offset..offset + 4]
            .try_into()
            .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
        let buffer_size_1 = u32::from_le_bytes(size_buffer_1) as usize;

        let size_buffer_2 = buffer[offset + 8..offset + 12]
            .try_into()
            .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
        let buffer_size_2 = u32::from_le_bytes(size_buffer_2) as usize;

        // Check for empty record special case
        if buffer_size_1 == 8 && buffer_size_2 == 0 {
            return Ok(ResourceEntry::Empty { offset });
        }

        // Verify size consistency
        if buffer_size_2 != buffer_size_1 - 8 {
            return Err(ResourceErrorKind::SizeMismatch {
                offset,
                expected: buffer_size_1 - 8,
                actual: buffer_size_2,
            });
        }

        let data_start = offset + 12;
        let data_end = data_start + buffer_size_2;

        if data_end > buffer.len() {
            return Err(ResourceErrorKind::OffsetOutOfBounds {
                offset: data_end,
                file_length: buffer.len(),
            });
        }

        Ok(ResourceEntry::Record12ByteHeader {
            data: buffer[data_start..data_end].to_vec(),
        })
    }

    /// Parses a 16-bit record entry (0xFF header).
    fn parse_16bit_record(
        buffer: &[u8],
        metadata: &ResourceEntryMetadata,
    ) -> Result<ResourceEntry, ResourceErrorKind> {
        let offset = metadata.offset;

        if offset + 3 > buffer.len() {
            return Err(ResourceErrorKind::HeaderReadError {
                offset,
                reason: "Not enough bytes for 16-bit header".to_string(),
            });
        }

        // Verify 0xFF marker
        if buffer[offset] != 0xFF {
            return Err(ResourceErrorKind::InvalidData {
                offset,
                details: format!("Expected 0xFF marker, got 0x{:02X}", buffer[offset]),
            });
        }

        let size_buffer = buffer[offset + 1..offset + 3]
            .try_into()
            .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
        let mut record_size = u16::from_le_bytes(size_buffer) as usize;

        // Handle VB6 IDE off-by-one bug
        if offset + 3 + record_size > buffer.len() {
            record_size -= 1;
        }

        let data_start = offset + 3;
        let data_end = data_start + record_size;

        if data_end > buffer.len() {
            return Err(ResourceErrorKind::OffsetOutOfBounds {
                offset: data_end,
                file_length: buffer.len(),
            });
        }

        Ok(ResourceEntry::Record3ByteHeader {
            data: buffer[data_start..data_end].to_vec(),
        })
    }

    /// Parses list items entry (signature [0x03, 0x00] or [0x07, 0x00]).
    fn parse_list_items(
        buffer: &[u8],
        metadata: &ResourceEntryMetadata,
    ) -> Result<ResourceEntry, ResourceErrorKind> {
        let offset = metadata.offset;

        if offset + 4 > buffer.len() {
            return Err(ResourceErrorKind::HeaderReadError {
                offset,
                reason: "Not enough bytes for list header".to_string(),
            });
        }

        let count_buffer = buffer[offset..offset + 2]
            .try_into()
            .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
        let item_count = u16::from_le_bytes(count_buffer) as usize;

        // Verify signature
        let signature = &buffer[offset + 2..offset + 4];
        if signature != [0x03, 0x00] && signature != [0x07, 0x00] {
            return Err(ResourceErrorKind::InvalidData {
                offset,
                details: format!("Invalid list signature: {signature:?}"),
            });
        }

        let mut items = Vec::with_capacity(item_count);
        let mut current_offset = offset + 4;

        for item_idx in 0..item_count {
            if current_offset + 2 > buffer.len() {
                return Err(ResourceErrorKind::CorruptedListItems {
                    offset,
                    details: format!("Item {item_idx} header out of bounds"),
                });
            }

            let item_size_buffer = buffer[current_offset..current_offset + 2]
                .try_into()
                .map_err(|_| ResourceErrorKind::BufferConversionError {
                    offset: current_offset,
                })?;
            let item_size = u16::from_le_bytes(item_size_buffer) as usize;

            let item_start = current_offset + 2;
            let item_end = item_start + item_size;

            if item_end > buffer.len() {
                return Err(ResourceErrorKind::CorruptedListItems {
                    offset,
                    details: format!("Item {item_idx} data out of bounds"),
                });
            }

            let item_bytes = &buffer[item_start..item_end];
            let item_string = String::from_utf8_lossy(item_bytes).to_string();
            items.push(item_string);

            current_offset = item_end;
        }

        Ok(ResourceEntry::ListItems { items })
    }

    /// Parses 4-byte header record (large text data).
    fn parse_text_data(
        buffer: &[u8],
        metadata: &ResourceEntryMetadata,
    ) -> Result<ResourceEntry, ResourceErrorKind> {
        let offset = metadata.offset;

        if offset + 4 > buffer.len() {
            return Err(ResourceErrorKind::HeaderReadError {
                offset,
                reason: "Not enough bytes for 4-byte header".to_string(),
            });
        }

        let size_buffer = buffer[offset..offset + 4]
            .try_into()
            .map_err(|_| ResourceErrorKind::BufferConversionError { offset })?;
        let record_size = u32::from_le_bytes(size_buffer) as usize;

        let data_start = offset + 4;
        let data_end = data_start + record_size;

        if data_end > buffer.len() {
            return Err(ResourceErrorKind::OffsetOutOfBounds {
                offset: data_end,
                file_length: buffer.len(),
            });
        }

        // Store raw bytes - caller can decode as needed
        let data = buffer[data_start..data_end].to_vec();

        Ok(ResourceEntry::Record4ByteHeader { data })
    }

    /// Parses 8-bit record entry (single-byte header).
    fn parse_8bit_record(
        buffer: &[u8],
        metadata: &ResourceEntryMetadata,
    ) -> Result<ResourceEntry, ResourceErrorKind> {
        let offset = metadata.offset;

        if offset >= buffer.len() {
            return Err(ResourceErrorKind::HeaderReadError {
                offset,
                reason: "Offset at end of file".to_string(),
            });
        }

        let mut record_size = buffer[offset] as usize;

        // Handle VB6 IDE off-by-one bug
        if offset + 1 + record_size > buffer.len() {
            record_size -= 1;
        }

        let data_start = offset + 1;
        let data_end = data_start + record_size;

        if data_end > buffer.len() {
            return Err(ResourceErrorKind::OffsetOutOfBounds {
                offset: data_end,
                file_length: buffer.len(),
            });
        }

        Ok(ResourceEntry::Record1ByteHeader {
            data: buffer[data_start..data_end].to_vec(),
        })
    }

    /// Gets a reference to a resource entry at the specified offset.
    ///
    /// # Arguments
    ///
    /// * `offset` - Byte offset where the resource entry begins
    ///
    /// # Returns
    ///
    /// `Some(&ResourceEntry)` if an entry exists at that offset, `None` otherwise.
    #[must_use]
    pub fn get_entry(&self, offset: usize) -> Option<&ResourceEntry> {
        self.entries.get(&offset)
    }

    /// Gets binary data from a resource entry, if it contains binary data.
    ///
    /// Works with `Record12ByteHeader`, `Record3ByteHeader`, `Record4ByteHeader`, and `Record1ByteHeader` entries.
    ///
    /// # Arguments
    ///
    /// * `offset` - Byte offset where the resource entry begins
    ///
    /// # Returns
    ///
    /// `Some(&[u8])` containing the data, or `None` if no entry exists at that
    /// offset or the entry is not a binary data type.
    #[must_use]
    pub fn get_binary_blob(&self, offset: usize) -> Option<&[u8]> {
        self.entries.get(&offset).and_then(|entry| match entry {
            ResourceEntry::Record12ByteHeader { data }
            | ResourceEntry::Record3ByteHeader { data }
            | ResourceEntry::Record4ByteHeader { data }
            | ResourceEntry::Record1ByteHeader { data } => Some(data.as_slice()),
            _ => None,
        })
    }

    /// Gets list items from a resource entry, if it contains a list.
    ///
    /// # Arguments
    ///
    /// * `offset` - Byte offset where the list resource entry begins
    ///
    /// # Returns
    ///
    /// `Some(&[String])` containing the list items, or `None` if no entry exists
    /// at that offset or the entry is not a `ListItems` type.
    #[must_use]
    pub fn get_list_items(&self, offset: usize) -> Option<&[String]> {
        self.entries.get(&offset).and_then(|entry| match entry {
            ResourceEntry::ListItems { items } => Some(items.as_slice()),
            _ => None,
        })
    }

    /// Gets raw data from a `Record4ByteHeader` resource entry.
    ///
    /// # Arguments
    ///
    /// * `offset` - Byte offset where the text resource entry begins
    ///
    /// # Returns
    ///
    /// `Some(&[u8])` containing the raw data, or `None` if no entry exists
    /// at that offset or the entry is not a `Record4ByteHeader` type.
    #[must_use]
    pub fn get_text_data(&self, offset: usize) -> Option<&[u8]> {
        self.entries.get(&offset).and_then(|entry| match entry {
            ResourceEntry::Record4ByteHeader { data } => Some(data.as_slice()),
            _ => None,
        })
    }

    /// Returns an iterator over all entries in the resource file.
    ///
    /// Entries are yielded in arbitrary order (`HashMap` iteration order).
    pub fn iter_entries(&self) -> impl Iterator<Item = (usize, &ResourceEntry)> {
        self.entries.iter().map(|(&offset, entry)| (offset, entry))
    }

    /// Returns the total number of resource entries in the file.
    #[must_use]
    pub fn entry_count(&self) -> usize {
        self.entries.len()
    }

    /// Returns the total size of the resource file in bytes.
    #[must_use]
    pub fn file_size(&self) -> usize {
        self.buffer.len()
    }
}
