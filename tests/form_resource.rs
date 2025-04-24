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

    assert_eq!(about_icon.len(), 0);
    assert_eq!(about_picture.len(), 775);
    assert_eq!(label6_caption.len(), 141);
    assert_eq!(label4_caption.len(), 299);

    assert_eq!(about_icon, resource_file_bytes[12..(12 + 0)].to_vec());
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

#[test]
fn audiostation_busy_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Busy.frx");

    let busy_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Busy.frx".to_owned(),
        0x00,
    ) {
        Ok(busy_icon) => busy_icon,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    assert_eq!(busy_icon.len(), 0);

    assert_eq!(busy_icon, resource_file_bytes[12..(12 + 0)].to_vec());
}

#[test]
fn audiostation_init_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Init.frx");

    let init_icon_offset = 0x00;
    let init_icon_header_size = 12;
    let init_icon_buffer_size = 0;

    let alpha_img_ctl2_image_offset = 0x0C;
    let alpha_img_ctl2_image_header_size = 4;
    let alpha_img_ctl2_image_buffer_size = 113643;

    let alpha_img_ctl2_effects_offset = 0x1BBFB;
    let alpha_img_ctl2_effects_header_size = 4;
    let alpha_img_ctl2_effects_buffer_size = 20;

    let alpha_img_ctl1_image_offset = 0x1BC13;
    let alpha_img_ctl1_image_header_size = 4;
    let alpha_img_ctl1_image_buffer_size = 34141;

    let alpha_img_ctl1_effects_offset = 0x24174;
    let alpha_img_ctl1_effects_header_size = 4;
    let alpha_img_ctl1_effects_buffer_size = 20;

    let init_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx".to_owned(),
        init_icon_offset,
    ) {
        Ok(init_icon) => init_icon,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let alpha_img_ctl2_image = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx".to_owned(),
        alpha_img_ctl2_image_offset,
    ) {
        Ok(alpha_img_ctl2_image) => alpha_img_ctl2_image,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let alpha_img_ctl2_effects = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx".to_owned(),
        alpha_img_ctl2_effects_offset,
    ) {
        Ok(alpha_img_ctl2_effects) => alpha_img_ctl2_effects,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let alpha_img_ctl1_image = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx".to_owned(),
        alpha_img_ctl1_image_offset,
    ) {
        Ok(alpha_img_ctl1_image) => alpha_img_ctl1_image,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let alpha_img_ctl1_effects = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx".to_owned(),
        alpha_img_ctl1_effects_offset,
    ) {
        Ok(alpha_img_ctl1_effects) => alpha_img_ctl1_effects,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    assert_eq!(init_icon.len(), init_icon_buffer_size);
    assert_eq!(alpha_img_ctl2_image.len(), alpha_img_ctl2_image_buffer_size);
    assert_eq!(
        alpha_img_ctl2_effects.len(),
        alpha_img_ctl2_effects_buffer_size
    );
    assert_eq!(alpha_img_ctl1_image.len(), alpha_img_ctl1_image_buffer_size);
    assert_eq!(
        alpha_img_ctl1_effects.len(),
        alpha_img_ctl1_effects_buffer_size
    );

    let init_icon_buffer_start = init_icon_offset + init_icon_header_size;
    let init_icon_buffer_end = init_icon_buffer_start + init_icon.len();

    assert_eq!(
        init_icon,
        resource_file_bytes[init_icon_buffer_start..init_icon_buffer_end].to_vec()
    );

    let alpha_img_ctl2_image_buffer_start =
        alpha_img_ctl2_image_offset + alpha_img_ctl2_image_header_size;
    let alpha_img_ctl2_image_buffer_end =
        alpha_img_ctl2_image_buffer_start + alpha_img_ctl2_image.len();

    assert_eq!(
        alpha_img_ctl2_image,
        resource_file_bytes[alpha_img_ctl2_image_buffer_start..alpha_img_ctl2_image_buffer_end]
            .to_vec()
    );

    let alpha_img_ctl2_effects_buffer_start =
        alpha_img_ctl2_effects_offset + alpha_img_ctl2_effects_header_size;
    let alpha_img_ctl2_effects_buffer_end =
        alpha_img_ctl2_effects_buffer_start + alpha_img_ctl2_effects.len();

    assert_eq!(
        alpha_img_ctl2_effects,
        resource_file_bytes[alpha_img_ctl2_effects_buffer_start..alpha_img_ctl2_effects_buffer_end]
            .to_vec()
    );

    let alpha_img_ctl1_image_buffer_start =
        alpha_img_ctl1_image_offset + alpha_img_ctl1_image_header_size;
    let alpha_img_ctl1_image_buffer_end =
        alpha_img_ctl1_image_buffer_start + alpha_img_ctl1_image.len();

    assert_eq!(
        alpha_img_ctl1_image,
        resource_file_bytes[alpha_img_ctl1_image_buffer_start..alpha_img_ctl1_image_buffer_end]
            .to_vec()
    );

    let alpha_img_ctl1_effects_buffer_start =
        alpha_img_ctl1_effects_offset + alpha_img_ctl1_effects_header_size;
    let alpha_img_ctl1_effects_buffer_end =
        alpha_img_ctl1_effects_buffer_start + alpha_img_ctl1_effects.len();

    assert_eq!(
        alpha_img_ctl1_effects,
        resource_file_bytes[alpha_img_ctl1_effects_buffer_start..alpha_img_ctl1_effects_buffer_end]
            .to_vec()
    );
}
