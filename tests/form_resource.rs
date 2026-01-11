use vb6parse::files::resource::{FormResourceFile, ResourceEntry};

#[test]
fn audiostation_about_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 4);

    // Offset 0x00 - Empty icon placeholder
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));

    // Offset 0x0C - Picture Record12ByteHeader (775 bytes)
    assert_eq!(entries[1].0, 0x0C);
    if let ResourceEntry::Record12ByteHeader { data } = entries[1].1 {
        assert_eq!(data.len(), 775);
    } else {
        panic!("Expected Record12ByteHeader at 0x0C");
    }

    // Offset 0x31F - Record4ByteHeader (141 bytes)
    assert_eq!(entries[2].0, 0x31F);
    if let ResourceEntry::Record4ByteHeader { data } = entries[2].1 {
        assert_eq!(data.len(), 141);
        let text = entries[2].1.as_text().expect("Should decode as text");
        assert_eq!(text, "The program is distributed in the hope that it will be useful, but without any warranty. it is provided \"as is\" without warranty of any kind.");
    } else {
        panic!("Expected Record4ByteHeader at 0x31F");
    }

    // Offset 0x3B0 - Record4ByteHeader (299 bytes)
    assert_eq!(entries[3].0, 0x3B0);
    if let ResourceEntry::Record4ByteHeader { data } = entries[3].1 {
        assert_eq!(data.len(), 299);
        let text = entries[3].1.as_text().expect("Should decode as text");
        assert_eq!(text, "Audiostation is a typical old media player. Media players like audiostation where made for Windows 98 and changed the way of playing music on your computer. But now you can bring back those times with our Audiostation software. Just download and install and enjoy the old look and feel of Windows 98");
    } else {
        panic!("Expected Record4ByteHeader at 0x3B0");
    }
}

#[test]
fn audiostation_busy_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Busy.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));
}

#[test]
fn audiostation_init_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 5);

    // Offset 0x00 - Empty
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));

    // Offset 0x0C - Record4ByteHeader (large image, 113643 bytes, contains PNG signature)
    assert_eq!(entries[1].0, 0x0C);
    if let ResourceEntry::Record4ByteHeader { data } = entries[1].1 {
        assert_eq!(data.len(), 113_643);
        // PNG signature is embedded in the data (not at the start due to VB6 wrapper)
        let png_sig = b"\x89PNG";
        assert!(
            data.windows(png_sig.len()).any(|w| w == png_sig),
            "PNG signature not found in data"
        );
    } else {
        panic!("Expected Record4ByteHeader at 0x0C");
    }

    // Offset 0x1BBFB - Record4ByteHeader (20 bytes, binary data)
    assert_eq!(entries[2].0, 0x1BBFB);
    if let ResourceEntry::Record4ByteHeader { data } = entries[2].1 {
        assert_eq!(data.len(), 20);
    } else {
        panic!("Expected Record4ByteHeader at 0x1BBFB");
    }

    // Offset 0x1BC13 - Record4ByteHeader (another large image, 34141 bytes, contains PNG signature)
    assert_eq!(entries[3].0, 0x1BC13);
    if let ResourceEntry::Record4ByteHeader { data } = entries[3].1 {
        assert_eq!(data.len(), 34_141);
        // Check for PNG signature bytes
        let png_sig = b"\x89PNG";
        assert!(
            data.windows(png_sig.len()).any(|w| w == png_sig),
            "PNG signature not found in data"
        );
    } else {
        panic!("Expected Record4ByteHeader at 0x1BC13");
    }

    // Offset 0x24174 - Record4ByteHeader (20 bytes, binary data)
    assert_eq!(entries[4].0, 0x24174);
    if let ResourceEntry::Record4ByteHeader { data } = entries[4].1 {
        assert_eq!(data.len(), 20);
    } else {
        panic!("Expected Record4ByteHeader at 0x24174");
    }
}

#[test]
fn audiostation_main_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    // Form_Main has 84 total entries
    assert_eq!(entries.len(), 84);

    // Offset 0x00 - Record12ByteHeader (2175 bytes - main icon)
    assert_eq!(entries[0].0, 0x00);
    if let ResourceEntry::Record12ByteHeader { data } = entries[0].1 {
        assert_eq!(data.len(), 2175);
    } else {
        panic!("Expected Record12ByteHeader at 0x00");
    }

    // Offset 0x88B - Record12ByteHeader (93938 bytes - large image)
    assert_eq!(entries[1].0, 0x88B);
    if let ResourceEntry::Record12ByteHeader { data } = entries[1].1 {
        assert_eq!(data.len(), 93_938);
    } else {
        panic!("Expected Record12ByteHeader at 0x88B");
    }

    // Count entry types
    let binary_blob_count = entries
        .iter()
        .filter(|(_, e)| matches!(e, ResourceEntry::Record12ByteHeader { .. }))
        .count();
    let text_data_count = entries
        .iter()
        .filter(|(_, e)| matches!(e, ResourceEntry::Record4ByteHeader { .. }))
        .count();

    assert_eq!(binary_blob_count, 52);
    assert_eq!(text_data_count, 32);

    // Verify a few Record4ByteHeader entries (BMP images) have correct length and signature
    assert_eq!(entries[2].0, 0x17789); // Entry 2
    if let ResourceEntry::Record4ByteHeader { data } = entries[2].1 {
        assert_eq!(data.len(), 6798);
        // BMP signature is after VB6 wrappers (at offset 24)
        assert_eq!(&data[24..26], b"BM");
    }

    assert_eq!(entries[3].0, 0x1921B); // Entry 3
    if let ResourceEntry::Record4ByteHeader { data } = entries[3].1 {
        assert_eq!(data.len(), 6798);
        assert_eq!(&data[24..26], b"BM");
    }

    assert_eq!(entries[4].0, 0x1ACAD); // Entry 4
    if let ResourceEntry::Record4ByteHeader { data } = entries[4].1 {
        assert_eq!(data.len(), 6798);
        assert_eq!(&data[24..26], b"BM");
    }
}

#[test]
fn audiostation_normalize_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Normalize.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));
}

#[test]
fn audiostation_open_dialog_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_OpenDialog.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 4);

    // Offset 0x00 - Empty
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));

    // Entry 1: Offset 0x0C - Record4ByteHeader (24 bytes, cursor/icon data)
    assert_eq!(entries[1].0, 0x0C);
    if let ResourceEntry::Record4ByteHeader { data } = entries[1].1 {
        assert_eq!(data.len(), 24);
    } else {
        panic!("Expected Record4ByteHeader at 0x0C");
    }

    // Entry 2: Offset 0x28 - Record4ByteHeader (343 bytes, cursor/icon data)
    assert_eq!(entries[2].0, 0x28);
    if let ResourceEntry::Record4ByteHeader { data } = entries[2].1 {
        assert_eq!(data.len(), 343);
    } else {
        panic!("Expected Record4ByteHeader at 0x28");
    }

    // Entry 3: Offset 0x183 - Record4ByteHeader (1430 bytes, large cursor/icon)
    assert_eq!(entries[3].0, 0x183);
    if let ResourceEntry::Record4ByteHeader { data } = entries[3].1 {
        assert_eq!(data.len(), 1430);
    } else {
        panic!("Expected Record4ByteHeader at 0x183");
    }
}

#[test]
fn audiostation_playlist_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Playlist.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));
}

#[test]
fn audiostation_plugins_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Plugins.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));
}

#[test]
fn audiostation_settings_recorder_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Settings_Recorder.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));
}

#[test]
fn audiostation_settings_record_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Settings_Record.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 5);

    // Offset 0x00 - Empty
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));

    // Offset 0x0C - ListItems (3 items, all "0")
    assert_eq!(entries[1].0, 0x0C);
    if let ResourceEntry::ListItems { items } = entries[1].1 {
        assert_eq!(items.len(), 3);
        assert_eq!(items[0], "0");
        assert_eq!(items[1], "0");
        assert_eq!(items[2], "0");
    } else {
        panic!("Expected ListItems at 0x0C");
    }

    // Offset 0x19 - ListItems (3 language items)
    assert_eq!(entries[2].0, 0x19);
    if let ResourceEntry::ListItems { items } = entries[2].1 {
        assert_eq!(items.len(), 3);
        assert_eq!(items[0], "English");
        assert_eq!(items[1], "Dutch");
        assert_eq!(items[2], "German");
    } else {
        panic!("Expected ListItems at 0x19");
    }

    // Offset 0x35 - ListItems (3 items, all "0")
    assert_eq!(entries[3].0, 0x35);
    if let ResourceEntry::ListItems { items } = entries[3].1 {
        assert_eq!(items.len(), 3);
        assert_eq!(items[0], "0");
        assert_eq!(items[1], "0");
        assert_eq!(items[2], "0");
    } else {
        panic!("Expected ListItems at 0x35");
    }

    // Offset 0x42 - ListItems (3 language items)
    assert_eq!(entries[4].0, 0x42);
    if let ResourceEntry::ListItems { items } = entries[4].1 {
        assert_eq!(items.len(), 3);
        assert_eq!(items[0], "English");
        assert_eq!(items[1], "Dutch");
        assert_eq!(items[2], "German");
    } else {
        panic!("Expected ListItems at 0x42");
    }
}

#[test]
fn audiostation_streams_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Streams.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));
}

#[test]
fn audiostation_track_properties_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Track_Properties.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 2);

    // Offset 0x00 - Empty
    assert_eq!(entries[0].0, 0x00);
    assert!(matches!(entries[0].1, ResourceEntry::Empty { .. }));

    // Offset 0x0C - Record1ByteHeader (20 bytes)
    assert_eq!(entries[1].0, 0x0C);
    if let ResourceEntry::Record1ByteHeader { data } = entries[1].1 {
        assert_eq!(data.len(), 20);
    } else {
        panic!("Expected Record1ByteHeader at 0x0C");
    }
}
