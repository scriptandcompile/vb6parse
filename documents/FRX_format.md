# Visual Basic 6 Form Resource File Format (.frx)

## Overview

VB6 Form Resource files (`.frx`) are binary files that store property data for forms and controls that cannot be represented as plain text in the `.frm` file. These files contain a sequence of variable-length records without an overall file header. Each record is referenced from the `.frm` file using a byte offset notation like `$"FormName.frx":0000`.

FRX files use **Windows-1252** encoding for text data and store values in **little-endian** byte order.

## File Structure

```
┌────────────────────┐
│  Entry 1           │  ← Offset 0x0000
│  (variable size)   │
├────────────────────┤
│  Entry 2           │  ← Offset depends on Entry 1 size
│  (variable size)   │
├────────────────────┤
│  Entry 3           │  ← Offset depends on previous entries
│  (variable size)   │
├────────────────────┤
│  ...               │
└────────────────────┘
```

There is **no file header** - the file immediately begins with the first resource entry. Each entry's location is calculated by summing the sizes of all previous entries.

## Entry Types and Headers

FRX files contain five distinct entry types, each with a unique header format:

### 1. Record12ByteHeader (Binary Blobs)

**Magic Signature:** `"lt\0\0"` at bytes 4-7

**Header Structure (12 bytes):**
```
Offset  Size  Description
------  ----  -----------
0x00    4     Size from end of signature to end of data (u32 LE)
0x04    4     Magic signature: 0x6C 0x74 0x00 0x00 ("lt\0\0")
0x08    4     Size of data section only (u32 LE) = [offset 0x00] - 8
0x0C    N     Binary data payload
```

**Used For:**
- Icons (`.ico`)
- Cursor files (`.cur`)
- Bitmap images (`.bmp`, `.dib`)
- PNG images (embedded as raw PNG data)
- OLE objects
- Picture properties (`Icon`, `Picture`, `MouseIcon`)

**Size Validation:**
The two size fields must satisfy: `size_at_0x08 = size_at_0x00 - 8`

**Example (Icon at offset 0x0000):**
```
Offset  Hex Data                         Description
------  -------------------------------- -----------
0x0000  3E 04 00 00                      Size: 0x043E (1086 bytes from end of sig)
0x0004  6C 74 00 00                      Magic: "lt\0\0"
0x0008  36 04 00 00                      Data size: 0x0436 (1078 bytes)
0x000C  00 00 01 00 02 00 10 10...       Icon data (ICONDIR structure)
```

**Special Case - Empty Record:**
When a user adds then removes an icon/image, VB6 IDE leaves behind an empty record:
```
0x00000008  (Size field)
0x6C 0x74 0x00 0x00  (Magic: "lt\0\0")
0x00000000  (Data size: 0)
```

### 2. Record4ByteHeader (Large Text/Binary Data)

**Identifier:** First 4 bytes contain at least one `0x00` byte

**Header Structure (4 bytes):**
```
Offset  Size  Description
------  ----  -----------
0x00    4     Size of data section (u32 LE)
0x04    N     Raw data payload
```

**Used For:**
- Long text strings (Caption, ToolTipText with >255 characters)
- Multiline text (Text property of TextBox)
- Form descriptions
- Binary data embedded in properties

**Data Encoding:**
- Text data: Windows-1252 encoded strings
- Binary data: Raw bytes (PNG, BMP, or other formats)

**Example (Long caption at offset 0x0000):**
```
Offset  Hex Data                         Description
------  -------------------------------- -----------
0x0000  A2 00 00 00                      Size: 0xA2 (162 bytes)
0x0004  41 6C 73 6F 20 74 68 65 72...    "Also there are other ways..."
```

**Note:** The size field includes everything after the header, not the header itself.

### 3. Record3ByteHeader (Medium Text Data)

**Magic Marker:** `0xFF` at byte 0

**Header Structure (3 bytes):**
```
Offset  Size  Description
------  ----  -----------
0x00    1     Magic marker: 0xFF
0x01    2     Size of data section (u16 LE)
0x03    N     Data payload (typically text)
```

**Used For:**
- Medium-length strings (up to 65535 bytes)
- Text properties
- String data

**VB6 IDE Off-by-One Bug:**
The VB6 IDE sometimes writes N in the size field when the actual data is N-1 bytes. Parsers must check if reading N bytes would exceed the file length and adjust by subtracting 1.

**Example:**
```
Offset  Hex Data                         Description
------  -------------------------------- -----------
0x0000  FF                               Marker: 0xFF
0x0001  1A 00                            Size: 0x001A (26 bytes declared)
0x0003  54 68 69 73 20 69 73...          "This is text data..."
                                         (May actually be 25 bytes due to bug)
```

### 4. ListItems (ComboBox/ListBox Contents)

**Magic Signature:** `0x03 0x00` or `0x07 0x00` at bytes 2-3

**Header Structure (4+ bytes):**
```
Offset  Size  Description
------  ----  -----------
0x00    2     Number of items (u16 LE)
0x02    2     Magic signature: 0x03 0x00 or 0x07 0x00
0x04    N     Item entries (see below)
```

**Item Entry Format:**
Each item is stored sequentially:
```
Offset  Size  Description
------  ----  -----------
0x00    2     Length of item string (u16 LE)
0x02    N     Item string data (no null terminator)
```

**Used For:**
- `List` property of `ComboBox` controls
- `List` property of `ListBox` controls
- Predefined list items in the form designer

**Example (12 items at offset 0x0054):**
```
Offset  Hex Data                         Description
------  -------------------------------- -----------
0x0054  0C 00                            Item count: 12
0x0056  03 00                            Magic: 0x03 0x00
0x0058  01 00                            Item 0 length: 1
0x005A  30                               Item 0: "0"
0x005B  01 00                            Item 1 length: 1
0x005D  30                               Item 1: "0"
...     ...                              (10 more items)
```

**Signature Variants:**
- `0x03 0x00`: Standard list items
- `0x07 0x00`: Alternative format (observed in some projects)

Both formats use identical structure after the signature.

### 5. Record1ByteHeader (Small Data)

**Identifier:** Default/fallback type (no specific magic)

**Header Structure (1 byte):**
```
Offset  Size  Description
------  ----  -----------
0x00    1     Size of data section (u8)
0x01    N     Data payload (max 255 bytes)
```

**Used For:**
- Very short strings (< 256 bytes)
- Small binary chunks
- Fallback for unrecognized patterns

**VB6 IDE Off-by-One Bug:**
Like Record3ByteHeader, parsers should check if N bytes would exceed file bounds and reduce by 1 if necessary.

**Example:**
```
Offset  Hex Data                         Description
------  -------------------------------- -----------
0x0000  0E                               Size: 14 bytes
0x0001  53 68 6F 72 74 20 74 65 78 74    "Short text"
```

## Entry Detection Algorithm

The parser uses a waterfall detection approach at each offset:

```
1. Check for Record12ByteHeader:
   - Offset + 4-7 == "lt\0\0" ?
   - If yes: Parse 12-byte header
   - Special case: size1==8 && size2==0 → Empty record

2. Check for Record3ByteHeader:
   - buffer[offset] == 0xFF ?
   - If yes: Parse 3-byte header

3. Check for ListItems:
   - buffer[offset+2..offset+4] == [0x03, 0x00] or [0x07, 0x00] ?
   - If yes: Parse list structure

4. Check for Record4ByteHeader:
   - buffer[offset..offset+4] contains any 0x00 byte ?
   - If yes: Parse 4-byte header

5. Default: Record1ByteHeader
   - Parse single-byte header
```

This order is critical because later checks can produce false positives on earlier formats.

## Cross-Referencing with .frm Files

FRM files reference FRX entries using the syntax:

```vb
PropertyName = $"FormName.frx":OFFSET
```

Where `OFFSET` is a **hexadecimal** byte offset (without `0x` prefix) indicating where the resource entry begins in the FRX file.

**Examples from .frm files:**

```vb
' Icon property - references binary blob at offset 0x0000
Icon = "DebugMain.frx":0000

' Long caption - references large text at offset 0x00A6
Caption = $"Form4.frx":00A6

' ListBox items - references list structure at offset 0x0054
ItemData = "SQLGenerator.frx":0054
List = "SQLGenerator.frx":007C

' Multiline text - references text data at offset 0x0000
Text = "SQLGenerator.frx":0000
```

### Properties That Use FRX References

Common properties that store data in FRX files:

| Property        | Control Types            | FRX Entry Type      |
|----------------|--------------------------|---------------------|
| `Icon`         | Form, MDIForm            | Record12ByteHeader  |
| `Picture`      | Image, PictureBox, Form  | Record12ByteHeader  |
| `MouseIcon`    | All controls             | Record12ByteHeader  |
| `List`         | ListBox, ComboBox        | ListItems           |
| `ItemData`     | ListBox, ComboBox        | ListItems           |
| `Caption`      | Label, Button, etc.      | Record4ByteHeader   |
| `Text`         | TextBox                  | Record4ByteHeader   |
| `ToolTipText`  | All controls             | Record4ByteHeader   |
| `Tag`          | All controls             | Record4ByteHeader   |

## Parsing Considerations

### 1. Windows-1252 Encoding

All text data in FRX files uses Windows-1252 encoding, **not UTF-8**. Parsers must:
- Decode text entries using Windows-1252 codec
- Handle extended characters (0x80-0xFF range)
- Use replacement characters for invalid sequences

### 2. VB6 IDE Off-by-One Bug

The VB6 IDE has a known bug where it declares size N but writes N-1 bytes for:
- Record3ByteHeader entries
- Record1ByteHeader entries

**Detection:** If `offset + header_size + declared_size > file_length`, subtract 1 from declared_size.

### 3. Little-Endian Byte Order

All multi-byte integers are stored in **little-endian** format:
```
0x0010 0x0000  → 0x0010 (16 decimal)
0x3E 0x04 0x00 0x00 → 0x043E (1086 decimal)
```

### 4. No File-Level Metadata

FRX files contain:
- ❌ No file signature/header
- ❌ No version information  
- ❌ No entry count or index
- ❌ No checksums or validation

The only way to parse an FRX file is to sequentially scan from offset 0, identifying each entry's type and size, then advancing to the next entry.

### 5. Binary Data Recognition

Record12ByteHeader and Record4ByteHeader can contain various binary formats:

**Icons/Cursors:**
- Start with ICONDIR structure: `0x00 0x00 0x01 0x00...`
- May contain Windows ICO format data

**PNG Images:**
- Start with PNG signature: `0x89 0x50 0x4E 0x47 0x0D 0x0A 0x1A 0x0A`
- Embedded directly as PNG file data

**BMP/DIB Images:**
- May start with BITMAPINFOHEADER: `0x28 0x00 0x00 0x00...`
- Check for color table and pixel data

**Text vs Binary Heuristic:**
If attempting to decode as Windows-1252 succeeds and produces valid characters, treat as text. Otherwise, treat as binary data.

## Example: Complete Entry Parse

Given this FRX file hex dump:

```
00000000  A2 00 00 00 41 6C 73 6F  20 74 68 65 72 65 20 61  |....Also there a|
00000010  72 65 20 6F 74 68 65 72  20 77 61 79 73 2F 63 6F  |re other ways/co|
...
000000A0  34 2C 20 35 29 2E                                 |4, 5).|
000000A6  F4 00 00 00 46 6F 72 6D  34 20 69 73 20 61 20 6E  |....Form4 is a n|
```

**Entry 1 at offset 0x0000:**
- Header: `A2 00 00 00` (4 bytes)
- Type: Record4ByteHeader (has 0x00 bytes)
- Size: 0xA2 = 162 bytes
- Data: 162 bytes of text starting at 0x0004
- Next entry offset: 0x0004 + 162 = 0x00A6

**Entry 2 at offset 0x00A6:**
- Header: `F4 00 00 00` (4 bytes)  
- Type: Record4ByteHeader
- Size: 0xF4 = 244 bytes
- Data: 244 bytes of text starting at 0x00AA
- Next entry offset: 0x00AA + 244 = 0x019E

## Implementation Notes

### Robust Parsing Strategy

1. **Start at offset 0**
2. **Identify entry type** using detection algorithm
3. **Read header** to determine data size
4. **Extract data payload**
5. **Store entry** with its offset as key
6. **Advance offset** by header_size + data_size
7. **Repeat until EOF**

### Error Handling

Common errors to handle gracefully:

- **Offset out of bounds:** Entry extends past file end
- **Size mismatch:** Record12ByteHeader size fields don't match
- **Corrupted list:** ListItems structure truncated
- **Buffer conversion:** Not enough bytes for header
- **Invalid signature:** Record12ByteHeader lacks "lt\0\0"

Best practice: Continue parsing remaining entries even if one fails, accumulating non-fatal errors for reporting.

### Memory Efficiency

For large FRX files (>1MB):
- Use `HashMap<usize, ResourceEntry>` for O(1) offset lookups
- Store entries indexed by their starting offset
- Keep original buffer for reference slicing
- Avoid duplicating large binary blobs

## Historical Context

The FRX format was designed for Visual Basic 6 (released 1998) and reflects limitations of that era:

- **No compression:** Binary data stored raw
- **No Unicode:** Windows-1252 encoding only
- **IDE bugs:** Off-by-one size errors
- **Brittle format:** No version or magic signature
- **Sequential access:** Must parse from start to find entries

Modern parsers should handle all these quirks while providing robust error recovery and efficient random access to entries by offset.

## References

- VB6 source code patterns observed across 378 FRX files in test data
- Current vb6parse implementation: `src/parsers/resource/mod.rs`
- Form file format: Properties reference FRX via `$"file.frx":OFFSET` syntax

## Related Files

- `.frm` - Form definition file (text) containing references to FRX entries
- `.frx` - Form resource file (binary) containing the actual data
- `.cls` - Class files (may have FRX for persistence)
- `.ctl` - UserControl files (may have FRX)
- `.ctx` - UserControl resource file (same format as FRX)
