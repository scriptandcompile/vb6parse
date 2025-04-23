#[test]
fn audiostation_about_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_About.frx");

    let about_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx".to_owned(),
        0x00,
    ) {
        Ok(about_icon) => about_icon,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let about_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx".to_owned(),
        0x0C,
    ) {
        Ok(about_picture) => about_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let label6_caption = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx".to_owned(),
        0x031F,
    ) {
        Ok(label6_caption) => label6_caption,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let label4_caption = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx".to_owned(),
        0x03B0,
    ) {
        Ok(label4_caption) => label4_caption,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    assert!(matches!(about_icon.len(), 0));
    assert!(matches!(about_picture.len(), 775));
    assert!(matches!(label6_caption.len(), 141));
    assert!(matches!(label4_caption.len(), 299));

    assert_eq!(about_icon, resource_file_bytes[0..(0 + 0)].to_vec());
    assert_eq!(about_picture, resource_file_bytes[24..(24 + 775)].to_vec());
    assert_eq!(
        label6_caption,
        resource_file_bytes[803..(803 + 141)].to_vec()
    );
    assert_eq!(
        label4_caption,
        resource_file_bytes[948..(948 + 299)].to_vec()
    );
}
