#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::FormResourceFile;

fuzz_target!(|data: &[u8]| {
    // Parse the resource file - takes ownership of Vec
    let result = FormResourceFile::parse("fuzz.frx", data.to_vec());

    // Unpack and check failures
    let (resource_opt, failures) = result.unpack();

    // Exercise all failure fields
    for failure in failures {
        let _ = &failure.kind;
        let _ = failure.error_offset;
        let _ = failure.line_start;
        let _ = failure.line_end;
    }

    // If we got a resource file, validate its structure
    if let Some(resource) = resource_opt {
        // Exercise entry count and file size
        let _ = resource.entry_count();
        let _ = resource.file_size();

        // Iterate through all entries
        for (offset, entry) in resource.iter_entries() {
            // Try to get each entry type
            if let Some(blob) = resource.get_binary_blob(offset) {
                let _ = blob.len();
            }

            if let Some(items) = resource.get_list_items(offset) {
                for item in items {
                    let _ = item.len();
                }
            }

            if let Some(text_data) = resource.get_text_data(offset) {
                let _ = text_data.len();
            }

            // Exercise ResourceEntry methods
            let _ = entry.as_text();
            let _ = entry.as_bytes();

            // Check entry type variants
            match entry {
                vb6parse::files::resource::ResourceEntry::Record12ByteHeader { data } => {
                    let _ = data.len();
                }
                vb6parse::files::resource::ResourceEntry::Record3ByteHeader { data } => {
                    let _ = data.len();
                }
                vb6parse::files::resource::ResourceEntry::ListItems { items } => {
                    for item in items {
                        let _ = item.as_bytes();
                    }
                }
                vb6parse::files::resource::ResourceEntry::Record4ByteHeader { data } => {
                    let _ = data.len();
                }
                vb6parse::files::resource::ResourceEntry::Record1ByteHeader { data } => {
                    let _ = data.len();
                }
                vb6parse::files::resource::ResourceEntry::Empty { offset } => {
                    let _ = offset;
                }
            }
        }
    }
});
