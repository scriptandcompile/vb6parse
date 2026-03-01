use assert_matches::assert_matches;
use vb6parse::files::resource::{FormResourceFile, ResourceEntry};

#[test]
fn brightness_effect_part_1_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Record12ByteHeader { .. });
}

#[test]
fn brightness_effect_part_2_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Record12ByteHeader { .. });
}

#[test]
fn brightness_effect_part_3_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Record12ByteHeader { .. });
}

#[test]
fn brightness_effect_part_4_frx() {
    let result = FormResourceFile::from_file(
        "./tests/data/vb6-code/Brightness-effect/Part 4 - Even faster DIBs/Brightness.frx",
    )
    .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Record12ByteHeader { .. });
}

#[test]
fn fire_effect_frx() {
    let result = FormResourceFile::from_file("./tests/data/vb6-code/Fire-effect/frmFire.frx")
        .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Record12ByteHeader { .. });
}

#[test]
fn game_physics_basic_frx() {
    let result =
        FormResourceFile::from_file("./tests/data/vb6-code/Game-physics-basic/FormPhysics.frx")
            .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 10);
    // All entries are Record12ByteHeader
    for (i, (_, entry)) in entries.iter().enumerate() {
        assert_matches!(
            entry,
            ResourceEntry::Record12ByteHeader { .. },
            "Entry {} should be Record12ByteHeader",
            i
        );
    }
}

#[test]
fn gradient_2d_frx() {
    let result = FormResourceFile::from_file("./tests/data/vb6-code/Gradient-2D/Gradient.frx")
        .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Record4ByteHeader { .. });
}

#[test]
fn hidden_markov_model_frx() {
    let result =
        FormResourceFile::from_file("./tests/data/vb6-code/Hidden-Markov-model/frmHMM.frx")
            .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Record1ByteHeader { .. });
}

#[test]
fn map_editor_2d_frx() {
    let result = FormResourceFile::from_file("./tests/data/vb6-code/Map-editor-2D/Main Editor.frx")
        .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Record12ByteHeader { .. });
}

#[test]
fn transparency_2d_frx() {
    let result =
        FormResourceFile::from_file("./tests/data/vb6-code/Transparency-2D/frmTransparency.frx")
            .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert_eq!(entries.len(), 3);

    // Offset 0x00 - Empty
    assert_eq!(entries[0].0, 0x00);
    assert_matches!(entries[0].1, ResourceEntry::Empty { .. });

    // Offset 0x0C - Record12ByteHeader
    assert_eq!(entries[1].0, 0x0C);
    assert_matches!(entries[1].1, ResourceEntry::Record12ByteHeader { .. });

    // Offset 0xB2C5 - Record12ByteHeader
    assert_eq!(entries[2].0, 0xB2C5);
    assert_matches!(entries[2].1, ResourceEntry::Record12ByteHeader { .. });
}
