//! A simple example to demonstrate how to debug VB6 Form Resource Files (.frx).
//!
//! This example reads a VB6 Form Resource File, parses its content,
//! and then prints out the details of each resource entry for debugging purposes.
//!
//! To run this example, ensure that the test data submodules are initialized
//! to provide access to the sample .frx file.
//!

use vb6parse::FormResourceFile;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let file_path = if args.len() > 1 {
        &args[1]
    } else {
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx"
    };

    let result = FormResourceFile::from_file(file_path).expect("Failed to read file");

    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    println!("Total entries: {}", entries.len());
    for (i, (offset, entry)) in entries.iter().enumerate() {
        match entry {
            vb6parse::files::resource::ResourceEntry::Record12ByteHeader { data } => {
                println!(
                    "Entry {}: Offset 0x{:X} - Record12ByteHeader ({} bytes)",
                    i,
                    offset,
                    data.len()
                );
            }
            vb6parse::files::resource::ResourceEntry::Record4ByteHeader { data } => {
                println!(
                    "Entry {}: Offset 0x{:X} - Record4ByteHeader ({} bytes)",
                    i,
                    offset,
                    data.len()
                );
                // Try to decode as text
                if let Some(text) = entry.as_text() {
                    println!("As text ({} chars):\n{}", text.len(), text);
                } else {
                    println!("(Binary data, not valid Windows-1252 text)");
                }
                println!("---");
            }
            vb6parse::files::resource::ResourceEntry::ListItems { items } => {
                println!(
                    "Entry {}: Offset 0x{:X} - ListItems ({} items)",
                    i,
                    offset,
                    items.len()
                );
                for (j, item) in items.iter().enumerate() {
                    println!("  Item {j}: {item:?}");
                }
            }
            vb6parse::files::resource::ResourceEntry::Record3ByteHeader { data } => {
                println!(
                    "Entry {i}: Offset 0x{offset:X} - Record3ByteHeader ({} bytes)",
                    data.len()
                );
            }
            vb6parse::files::resource::ResourceEntry::Record1ByteHeader { data } => {
                println!(
                    "Entry {i}: Offset 0x{offset:X} - Record1ByteHeader ({} bytes)",
                    data.len()
                );
            }
            vb6parse::files::resource::ResourceEntry::Empty { .. } => {
                println!("Entry {i}: Offset 0x{offset:X} - Empty (removed image placeholder)");
            }
        }
    }
}
