#[test]
fn audiostation_about_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_About.frx");

    let about_icon_offset = 0x00;
    let about_icon_header_size = 12;
    let about_icon_buffer_size = 0;
    let about_icon_buffer_start = about_icon_offset + about_icon_header_size;
    let about_icon_buffer_end = about_icon_buffer_start + about_icon_buffer_size;
    let about_icon_buffer =
        resource_file_bytes[about_icon_buffer_start..about_icon_buffer_end].to_vec();

    let about_picture_offset = 0x0C;
    let about_picture_header_size = 12;
    let about_picture_buffer_size = 775;
    let about_picture_buffer_start = about_picture_offset + about_picture_header_size;
    let about_picture_buffer_end = about_picture_buffer_start + about_picture_buffer_size;
    let about_picture_buffer =
        resource_file_bytes[about_picture_buffer_start..about_picture_buffer_end].to_vec();

    let label6_caption_offset = 0x031F;
    let label6_caption_header_size = 4;
    let label6_caption_buffer_size = 141;
    let label6_caption_buffer_start = label6_caption_offset + label6_caption_header_size;
    let label6_caption_buffer_end = label6_caption_buffer_start + label6_caption_buffer_size;
    let label6_caption_buffer =
        resource_file_bytes[label6_caption_buffer_start..label6_caption_buffer_end].to_vec();

    let label4_caption_offset = 0x03B0;
    let label4_caption_header_size = 4;
    let label4_caption_buffer_size = 299;
    let label4_caption_buffer_start = label4_caption_offset + label4_caption_header_size;
    let label4_caption_buffer_end = label4_caption_buffer_start + label4_caption_buffer_size;
    let label4_caption_buffer =
        resource_file_bytes[label4_caption_buffer_start..label4_caption_buffer_end].to_vec();

    let about_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx".to_owned(),
        about_icon_offset,
    ) {
        Ok(about_icon) => about_icon,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let about_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx".to_owned(),
        about_picture_offset,
    ) {
        Ok(about_picture) => about_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let label6_caption = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx".to_owned(),
        label6_caption_offset,
    ) {
        Ok(label6_caption) => label6_caption,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let label4_caption = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx".to_owned(),
        label4_caption_offset,
    ) {
        Ok(label4_caption) => label4_caption,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    assert_eq!(about_icon.len(), about_icon_buffer_size);
    assert_eq!(about_picture.len(), about_picture_buffer_size);
    assert_eq!(label6_caption.len(), label6_caption_buffer_size);
    assert_eq!(label4_caption.len(), label4_caption_buffer_size);

    assert_eq!(about_icon, about_icon_buffer);
    assert_eq!(about_picture, about_picture_buffer);
    assert_eq!(label6_caption, label6_caption_buffer);
    assert_eq!(label4_caption, label4_caption_buffer);
}

#[test]
fn audiostation_busy_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Busy.frx");

    let busy_icon_offset = 0x00;
    let busy_icon_header_size = 12;
    let busy_icon_buffer_size = 0;
    let busy_icon_buffer_start = busy_icon_offset + busy_icon_header_size;
    let busy_icon_buffer_end = busy_icon_buffer_start + busy_icon_buffer_size;
    let busy_icon_buffer =
        resource_file_bytes[busy_icon_buffer_start..busy_icon_buffer_end].to_vec();

    let busy_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Busy.frx".to_owned(),
        busy_icon_offset,
    ) {
        Ok(busy_icon) => busy_icon,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    assert_eq!(busy_icon.len(), busy_icon_buffer_size);

    assert_eq!(busy_icon, busy_icon_buffer);
}

#[test]
fn audiostation_init_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Init.frx");

    let init_icon_offset = 0x00;
    let init_icon_header_size = 12;
    let init_icon_buffer_size = 0;
    let init_icon_buffer_start = init_icon_offset + init_icon_header_size;
    let init_icon_buffer_end = init_icon_buffer_start + init_icon_buffer_size;
    let init_icon_buffer =
        resource_file_bytes[init_icon_buffer_start..init_icon_buffer_end].to_vec();

    let alpha_img_ctl2_image_offset = 0x0C;
    let alpha_img_ctl2_image_header_size = 4;
    let alpha_img_ctl2_image_buffer_size = 113643;
    let alpha_img_ctl2_image_buffer_start =
        alpha_img_ctl2_image_offset + alpha_img_ctl2_image_header_size;
    let alpha_img_ctl2_image_buffer_end =
        alpha_img_ctl2_image_buffer_start + alpha_img_ctl2_image_buffer_size;
    let alpha_img_ctl2_image_buffer = resource_file_bytes
        [alpha_img_ctl2_image_buffer_start..alpha_img_ctl2_image_buffer_end]
        .to_vec();

    let alpha_img_ctl2_effects_offset = 0x1BBFB;
    let alpha_img_ctl2_effects_header_size = 4;
    let alpha_img_ctl2_effects_buffer_size = 20;
    let alpha_img_ctl2_effects_buffer_start =
        alpha_img_ctl2_effects_offset + alpha_img_ctl2_effects_header_size;
    let alpha_img_ctl2_effects_buffer_end =
        alpha_img_ctl2_effects_buffer_start + alpha_img_ctl2_effects_buffer_size;
    let alpha_img_ctl2_effects_buffer = resource_file_bytes
        [alpha_img_ctl2_effects_buffer_start..alpha_img_ctl2_effects_buffer_end]
        .to_vec();

    let alpha_img_ctl1_image_offset = 0x1BC13;
    let alpha_img_ctl1_image_header_size = 4;
    let alpha_img_ctl1_image_buffer_size = 34141;
    let alpha_img_ctl1_image_buffer_start =
        alpha_img_ctl1_image_offset + alpha_img_ctl1_image_header_size;
    let alpha_img_ctl1_image_buffer_end =
        alpha_img_ctl1_image_buffer_start + alpha_img_ctl1_image_buffer_size;
    let alpha_img_ctl1_image_buffer = resource_file_bytes
        [alpha_img_ctl1_image_buffer_start..alpha_img_ctl1_image_buffer_end]
        .to_vec();

    let alpha_img_ctl1_effects_offset = 0x24174;
    let alpha_img_ctl1_effects_header_size = 4;
    let alpha_img_ctl1_effects_buffer_size = 20;
    let alpha_img_ctl1_effects_buffer_start =
        alpha_img_ctl1_effects_offset + alpha_img_ctl1_effects_header_size;
    let alpha_img_ctl1_effects_buffer_end =
        alpha_img_ctl1_effects_buffer_start + alpha_img_ctl1_effects_buffer_size;
    let alpha_img_ctl1_effects_buffer = resource_file_bytes
        [alpha_img_ctl1_effects_buffer_start..alpha_img_ctl1_effects_buffer_end]
        .to_vec();

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

    assert_eq!(init_icon, init_icon_buffer);
    assert_eq!(alpha_img_ctl2_image, alpha_img_ctl2_image_buffer);
    assert_eq!(alpha_img_ctl2_effects, alpha_img_ctl2_effects_buffer);
    assert_eq!(alpha_img_ctl1_image, alpha_img_ctl1_image_buffer);
    assert_eq!(alpha_img_ctl1_effects, alpha_img_ctl1_effects_buffer);
}

#[test]
fn audiostation_main_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Main.frx");

    let main_icon_offset = 0x00;
    let main_icon_header_size = 12;
    let main_icon_buffer_size = 2175;
    let main_icon_buffer_start = main_icon_offset + main_icon_header_size;
    let main_icon_buffer_end = main_icon_buffer_start + main_icon_buffer_size;
    let main_icon_buffer =
        resource_file_bytes[main_icon_buffer_start..main_icon_buffer_end].to_vec();

    let element1_picture_offset = 0x088B;
    let element1_picture_header_size = 12;
    let element1_picture_buffer_size = 93938;
    let element1_picture_buffer_start = element1_picture_offset + element1_picture_header_size;
    let element1_picture_buffer_end = element1_picture_buffer_start + element1_picture_buffer_size;
    let element1_picture_buffer =
        resource_file_bytes[element1_picture_buffer_start..element1_picture_buffer_end].to_vec();

    let cd_display_list_image1_picture_offset = 0x17789;
    let cd_display_list_image1_picture_header_size = 4;
    let cd_display_list_image1_picture_buffer_size = 6798;
    let cd_display_list_image1_picture_buffer_start =
        cd_display_list_image1_picture_offset + cd_display_list_image1_picture_header_size;
    let cd_display_list_image1_picture_buffer_end =
        cd_display_list_image1_picture_buffer_start + cd_display_list_image1_picture_buffer_size;
    let cd_display_list_image1_picture_buffer = resource_file_bytes
        [cd_display_list_image1_picture_buffer_start..cd_display_list_image1_picture_buffer_end]
        .to_vec();

    let cd_display_list_image2_picture_offset = 0x1921B;
    let cd_display_list_image2_picture_header_size = 4;
    let cd_display_list_image2_picture_buffer_size = 6798;
    let cd_display_list_image2_picture_buffer_start =
        cd_display_list_image2_picture_offset + cd_display_list_image2_picture_header_size;
    let cd_display_list_image2_picture_buffer_end =
        cd_display_list_image2_picture_buffer_start + cd_display_list_image2_picture_buffer_size;
    let cd_display_list_image2_picture_buffer = resource_file_bytes
        [cd_display_list_image2_picture_buffer_start..cd_display_list_image2_picture_buffer_end]
        .to_vec();

    let cd_display_list_image3_picture_offset = 0x1ACAD;
    let cd_display_list_image3_picture_header_size = 4;
    let cd_display_list_image3_picture_buffer_size = 6798;
    let cd_display_list_image3_picture_buffer_start =
        cd_display_list_image3_picture_offset + cd_display_list_image3_picture_header_size;
    let cd_display_list_image3_picture_buffer_end =
        cd_display_list_image3_picture_buffer_start + cd_display_list_image3_picture_buffer_size;
    let cd_display_list_image3_picture_buffer = resource_file_bytes
        [cd_display_list_image3_picture_buffer_start..cd_display_list_image3_picture_buffer_end]
        .to_vec();

    let cd_display_list_image4_picture_offset = 0x1C73F;
    let cd_display_list_image4_picture_header_size = 4;
    let cd_display_list_image4_picture_buffer_size = 6798;
    let cd_display_list_image4_picture_buffer_start =
        cd_display_list_image4_picture_offset + cd_display_list_image4_picture_header_size;
    let cd_display_list_image4_picture_buffer_end =
        cd_display_list_image4_picture_buffer_start + cd_display_list_image4_picture_buffer_size;
    let cd_display_list_image4_picture_buffer = resource_file_bytes
        [cd_display_list_image4_picture_buffer_start..cd_display_list_image4_picture_buffer_end]
        .to_vec();

    let cd_display_list_image5_picture_offset = 0x1E1D1;
    let cd_display_list_image5_picture_header_size = 4;
    let cd_display_list_image5_picture_buffer_size = 6798;
    let cd_display_list_image5_picture_buffer_start =
        cd_display_list_image5_picture_offset + cd_display_list_image5_picture_header_size;
    let cd_display_list_image5_picture_buffer_end =
        cd_display_list_image5_picture_buffer_start + cd_display_list_image5_picture_buffer_size;
    let cd_display_list_image5_picture_buffer = resource_file_bytes
        [cd_display_list_image5_picture_buffer_start..cd_display_list_image5_picture_buffer_end]
        .to_vec();

    let cd_animation_list_image1_picture_offset = 0x1FC63;
    let cd_animation_list_image1_picture_header_size = 4;
    let cd_animation_list_image1_picture_buffer_size = 33862;
    let cd_animation_list_image1_picture_buffer_start =
        cd_animation_list_image1_picture_offset + cd_animation_list_image1_picture_header_size;
    let cd_animation_list_image1_picture_buffer_end = cd_animation_list_image1_picture_buffer_start
        + cd_animation_list_image1_picture_buffer_size;
    let cd_animation_list_image1_picture_buffer = resource_file_bytes
        [cd_animation_list_image1_picture_buffer_start
            ..cd_animation_list_image1_picture_buffer_end]
        .to_vec();

    let cd_animation_list_image2_picture_offset = 0x280AD;
    let cd_animation_list_image2_picture_header_size = 4;
    let cd_animation_list_image2_picture_buffer_size = 65600;
    let cd_animation_list_image2_picture_buffer_start =
        cd_animation_list_image2_picture_offset + cd_animation_list_image2_picture_header_size;
    let cd_animation_list_image2_picture_buffer_end = cd_animation_list_image2_picture_buffer_start
        + cd_animation_list_image2_picture_buffer_size;
    let cd_animation_list_image2_picture_buffer = resource_file_bytes
        [cd_animation_list_image2_picture_buffer_start
            ..cd_animation_list_image2_picture_buffer_end]
        .to_vec();

    let cd_animation_list_image3_picture_offset = 0x380F1;
    let cd_animation_list_image3_picture_header_size = 4;
    let cd_animation_list_image3_picture_buffer_size = 65600;
    let cd_animation_list_image3_picture_buffer_start =
        cd_animation_list_image3_picture_offset + cd_animation_list_image3_picture_header_size;
    let cd_animation_list_image3_picture_buffer_end = cd_animation_list_image3_picture_buffer_start
        + cd_animation_list_image3_picture_buffer_size;
    let cd_animation_list_image3_picture_buffer = resource_file_bytes
        [cd_animation_list_image3_picture_buffer_start
            ..cd_animation_list_image3_picture_buffer_end]
        .to_vec();

    let cd_animation_list_image4_picture_offset = 0x48135;
    let cd_animation_list_image4_picture_header_size = 4;
    let cd_animation_list_image4_picture_buffer_size = 65600;
    let cd_animation_list_image4_picture_buffer_start =
        cd_animation_list_image4_picture_offset + cd_animation_list_image4_picture_header_size;
    let cd_animation_list_image4_picture_buffer_end = cd_animation_list_image4_picture_buffer_start
        + cd_animation_list_image4_picture_buffer_size;
    let cd_animation_list_image4_picture_buffer = resource_file_bytes
        [cd_animation_list_image4_picture_buffer_start
            ..cd_animation_list_image4_picture_buffer_end]
        .to_vec();

    let cd_animation_list_image5_picture_offset = 0x58179;
    let cd_animation_list_image5_picture_header_size = 4;
    let cd_animation_list_image5_picture_buffer_size = 65600;
    let cd_animation_list_image5_picture_buffer_start =
        cd_animation_list_image5_picture_offset + cd_animation_list_image5_picture_header_size;
    let cd_animation_list_image5_picture_buffer_end = cd_animation_list_image5_picture_buffer_start
        + cd_animation_list_image5_picture_buffer_size;
    let cd_animation_list_image5_picture_buffer = resource_file_bytes
        [cd_animation_list_image5_picture_buffer_start
            ..cd_animation_list_image5_picture_buffer_end]
        .to_vec();

    let cd_animation_list_image6_picture_offset = 0x681BD;
    let cd_animation_list_image6_picture_header_size = 4;
    let cd_animation_list_image6_picture_buffer_size = 65600;
    let cd_animation_list_image6_picture_buffer_start =
        cd_animation_list_image6_picture_offset + cd_animation_list_image6_picture_header_size;
    let cd_animation_list_image6_picture_buffer_end = cd_animation_list_image6_picture_buffer_start
        + cd_animation_list_image6_picture_buffer_size;
    let cd_animation_list_image6_picture_buffer = resource_file_bytes
        [cd_animation_list_image6_picture_buffer_start
            ..cd_animation_list_image6_picture_buffer_end]
        .to_vec();

    let cd_animation_list_image7_picture_offset = 0x78201;
    let cd_animation_list_image7_picture_header_size = 4;
    let cd_animation_list_image7_picture_buffer_size = 65600;
    let cd_animation_list_image7_picture_buffer_start =
        cd_animation_list_image7_picture_offset + cd_animation_list_image7_picture_header_size;
    let cd_animation_list_image7_picture_buffer_end = cd_animation_list_image7_picture_buffer_start
        + cd_animation_list_image7_picture_buffer_size;
    let cd_animation_list_image7_picture_buffer = resource_file_bytes
        [cd_animation_list_image7_picture_buffer_start
            ..cd_animation_list_image7_picture_buffer_end]
        .to_vec();

    let element2_picture_offset = 0x88245;
    let element2_picture_header_size = 12;
    let element2_picture_buffer_size = 203150;
    let element2_picture_buffer_start = element2_picture_offset + element2_picture_header_size;
    let element2_picture_buffer_end = element2_picture_buffer_start + element2_picture_buffer_size;
    let element2_picture_buffer =
        resource_file_bytes[element2_picture_buffer_start..element2_picture_buffer_end].to_vec();

    let switch_master_glyph_offset = 0xB9BDF;
    let switch_master_glyph_header_size = 4;
    let switch_master_glyph_buffer_size = 82;
    let switch_master_glyph_buffer_start =
        switch_master_glyph_offset + switch_master_glyph_header_size;
    let switch_master_glyph_buffer_end =
        switch_master_glyph_buffer_start + switch_master_glyph_buffer_size;
    let switch_master_glyph_buffer = resource_file_bytes
        [switch_master_glyph_buffer_start..switch_master_glyph_buffer_end]
        .to_vec();

    let switch_rec_glyph_offset = 0xB9C35;
    let switch_rec_glyph_header_size = 4;
    let switch_rec_glyph_buffer_size = 82;
    let switch_rec_glyph_buffer_start = switch_rec_glyph_offset + switch_rec_glyph_header_size;
    let switch_rec_glyph_buffer_end = switch_rec_glyph_buffer_start + switch_rec_glyph_buffer_size;
    let switch_rec_glyph_buffer =
        resource_file_bytes[switch_rec_glyph_buffer_start..switch_rec_glyph_buffer_end].to_vec();

    let switch_cd_glyph_offset = 0xB9C8B;
    let switch_cd_glyph_header_size = 4;
    let switch_cd_glyph_buffer_size = 82;
    let switch_cd_glyph_buffer_start = switch_cd_glyph_offset + switch_cd_glyph_header_size;
    let switch_cd_glyph_buffer_end = switch_cd_glyph_buffer_start + switch_cd_glyph_buffer_size;
    let switch_cd_glyph_buffer =
        resource_file_bytes[switch_cd_glyph_buffer_start..switch_cd_glyph_buffer_end].to_vec();

    let switch_dat_glyph_offset = 0xB9CE1;
    let switch_dat_glyph_header_size = 4;
    let switch_dat_glyph_buffer_size = 82;
    let switch_dat_glyph_buffer_start = switch_dat_glyph_offset + switch_dat_glyph_header_size;
    let switch_dat_glyph_buffer_end = switch_dat_glyph_buffer_start + switch_dat_glyph_buffer_size;
    let switch_dat_glyph_buffer =
        resource_file_bytes[switch_dat_glyph_buffer_start..switch_dat_glyph_buffer_end].to_vec();

    let switch_midi_glyph_offset = 0xB9D37;
    let switch_midi_glyph_header_size = 4;
    let switch_midi_glyph_buffer_size = 82;
    let switch_midi_glyph_buffer_start = switch_midi_glyph_offset + switch_midi_glyph_header_size;
    let switch_midi_glyph_buffer_end =
        switch_midi_glyph_buffer_start + switch_midi_glyph_buffer_size;
    let switch_midi_glyph_buffer =
        resource_file_bytes[switch_midi_glyph_buffer_start..switch_midi_glyph_buffer_end].to_vec();

    let image1_picture_offset = 0xB9D8D;
    let image1_picture_header_size = 12;
    let image1_picture_buffer_size = 5430;
    let image1_picture_buffer_start = image1_picture_offset + image1_picture_header_size;
    let image1_picture_buffer_end = image1_picture_buffer_start + image1_picture_buffer_size;
    let image1_picture_buffer =
        resource_file_bytes[image1_picture_buffer_start..image1_picture_buffer_end].to_vec();

    let image2_picture_offset = 0xBB2CF;
    let image2_picture_header_size = 12;
    let image2_picture_buffer_size = 4878;
    let image2_picture_buffer_start = image2_picture_offset + image2_picture_header_size;
    let image2_picture_buffer_end = image2_picture_buffer_start + image2_picture_buffer_size;
    let image2_picture_buffer =
        resource_file_bytes[image2_picture_buffer_start..image2_picture_buffer_end].to_vec();

    let media_player_animation_list_image1_picture_offset = 0xBC5E9;
    let media_player_animation_list_image1_picture_header_size = 4;
    let media_player_animation_list_image1_picture_buffer_size = 492;
    let media_player_animation_list_image1_picture_buffer_start =
        media_player_animation_list_image1_picture_offset
            + media_player_animation_list_image1_picture_header_size;
    let media_player_animation_list_image1_picture_buffer_end =
        media_player_animation_list_image1_picture_buffer_start
            + media_player_animation_list_image1_picture_buffer_size;
    let media_player_animation_list_image1_picture_buffer = resource_file_bytes
        [media_player_animation_list_image1_picture_buffer_start
            ..media_player_animation_list_image1_picture_buffer_end]
        .to_vec();

    let media_player_animation_list_image2_picture_offset = 0xBC7D9;
    let media_player_animation_list_image2_picture_header_size = 4;
    let media_player_animation_list_image2_picture_buffer_size = 492;
    let media_player_animation_list_image2_picture_buffer_start =
        media_player_animation_list_image2_picture_offset
            + media_player_animation_list_image2_picture_header_size;
    let media_player_animation_list_image2_picture_buffer_end =
        media_player_animation_list_image2_picture_buffer_start
            + media_player_animation_list_image2_picture_buffer_size;
    let media_player_animation_list_image2_picture_buffer = resource_file_bytes
        [media_player_animation_list_image2_picture_buffer_start
            ..media_player_animation_list_image2_picture_buffer_end]
        .to_vec();

    let media_player_animation_list_image3_picture_offset = 0xBC9C9;
    let media_player_animation_list_image3_picture_header_size = 4;
    let media_player_animation_list_image3_picture_buffer_size = 491;
    let media_player_animation_list_image3_picture_buffer_start =
        media_player_animation_list_image3_picture_offset
            + media_player_animation_list_image3_picture_header_size;
    let media_player_animation_list_image3_picture_buffer_end =
        media_player_animation_list_image3_picture_buffer_start
            + media_player_animation_list_image3_picture_buffer_size;
    let media_player_animation_list_image3_picture_buffer = resource_file_bytes
        [media_player_animation_list_image3_picture_buffer_start
            ..media_player_animation_list_image3_picture_buffer_end]
        .to_vec();

    let media_player_animation_list_image4_picture_offset = 0xBCBB8;
    let media_player_animation_list_image4_picture_header_size = 4;
    let media_player_animation_list_image4_picture_buffer_size = 488;
    let media_player_animation_list_image4_picture_buffer_start =
        media_player_animation_list_image4_picture_offset
            + media_player_animation_list_image4_picture_header_size;
    let media_player_animation_list_image4_picture_buffer_end =
        media_player_animation_list_image4_picture_buffer_start
            + media_player_animation_list_image4_picture_buffer_size;
    let media_player_animation_list_image4_picture_buffer = resource_file_bytes
        [media_player_animation_list_image4_picture_buffer_start
            ..media_player_animation_list_image4_picture_buffer_end]
        .to_vec();

    let media_player_animation_list_image5_picture_offset = 0xBCDA4;
    let media_player_animation_list_image5_picture_header_size = 4;
    let media_player_animation_list_image5_picture_buffer_size = 492;
    let media_player_animation_list_image5_picture_buffer_start =
        media_player_animation_list_image5_picture_offset
            + media_player_animation_list_image5_picture_header_size;
    let media_player_animation_list_image5_picture_buffer_end =
        media_player_animation_list_image5_picture_buffer_start
            + media_player_animation_list_image5_picture_buffer_size;
    let media_player_animation_list_image5_picture_buffer = resource_file_bytes
        [media_player_animation_list_image5_picture_buffer_start
            ..media_player_animation_list_image5_picture_buffer_end]
        .to_vec();

    let media_player_animation_list_image6_picture_offset = 0xBCF94;
    let media_player_animation_list_image6_picture_header_size = 4;
    let media_player_animation_list_image6_picture_buffer_size = 495;
    let media_player_animation_list_image6_picture_buffer_start =
        media_player_animation_list_image6_picture_offset
            + media_player_animation_list_image6_picture_header_size;
    let media_player_animation_list_image6_picture_buffer_end =
        media_player_animation_list_image6_picture_buffer_start
            + media_player_animation_list_image6_picture_buffer_size;
    let media_player_animation_list_image6_picture_buffer = resource_file_bytes
        [media_player_animation_list_image6_picture_buffer_start
            ..media_player_animation_list_image6_picture_buffer_end]
        .to_vec();

    let media_player_animation_list_image7_picture_offset = 0xBD187;
    let media_player_animation_list_image7_picture_header_size = 4;
    let media_player_animation_list_image7_picture_buffer_size = 492;
    let media_player_animation_list_image7_picture_buffer_start =
        media_player_animation_list_image7_picture_offset
            + media_player_animation_list_image7_picture_header_size;
    let media_player_animation_list_image7_picture_buffer_end =
        media_player_animation_list_image7_picture_buffer_start
            + media_player_animation_list_image7_picture_buffer_size;
    let media_player_animation_list_image7_picture_buffer = resource_file_bytes
        [media_player_animation_list_image7_picture_buffer_start
            ..media_player_animation_list_image7_picture_buffer_end]
        .to_vec();

    let media_player_animation_list_image8_picture_offset = 0xBD377;
    let media_player_animation_list_image8_picture_header_size = 4;
    let media_player_animation_list_image8_picture_buffer_size = 488;
    let media_player_animation_list_image8_picture_buffer_start =
        media_player_animation_list_image8_picture_offset
            + media_player_animation_list_image8_picture_header_size;
    let media_player_animation_list_image8_picture_buffer_end =
        media_player_animation_list_image8_picture_buffer_start
            + media_player_animation_list_image8_picture_buffer_size;
    let media_player_animation_list_image8_picture_buffer = resource_file_bytes
        [media_player_animation_list_image8_picture_buffer_start
            ..media_player_animation_list_image8_picture_buffer_end]
        .to_vec();

    let element3_picture_offset = 0xBD563;
    let element3_picture_header_size = 12;
    let element3_picture_buffer_size = 193570;
    let element3_picture_buffer_start = element3_picture_offset + element3_picture_header_size;
    let element3_picture_buffer_end = element3_picture_buffer_start + element3_picture_buffer_size;
    let element3_picture_buffer =
        resource_file_bytes[element3_picture_buffer_start..element3_picture_buffer_end].to_vec();

    let button_cd_player1_picture_offset = 0xEC991;
    let button_cd_player1_picture_header_size = 12;
    let button_cd_player1_picture_buffer_size = 1470;
    let button_cd_player1_picture_buffer_start =
        button_cd_player1_picture_offset + button_cd_player1_picture_header_size;
    let button_cd_player1_picture_buffer_end =
        button_cd_player1_picture_buffer_start + button_cd_player1_picture_buffer_size;
    let button_cd_player1_picture_buffer = resource_file_bytes
        [button_cd_player1_picture_buffer_start..button_cd_player1_picture_buffer_end]
        .to_vec();

    let button_cd_player2_picture_offset = 0xECF5B;
    let button_cd_player2_picture_header_size = 12;
    let button_cd_player2_picture_buffer_size = 1406;
    let button_cd_player2_picture_buffer_start =
        button_cd_player2_picture_offset + button_cd_player2_picture_header_size;
    let button_cd_player2_picture_buffer_end =
        button_cd_player2_picture_buffer_start + button_cd_player2_picture_buffer_size;
    let button_cd_player2_picture_buffer = resource_file_bytes
        [button_cd_player2_picture_buffer_start..button_cd_player2_picture_buffer_end]
        .to_vec();

    let button_cd_player3_picture_offset = 0xED4E5;
    let button_cd_player3_picture_header_size = 12;
    let button_cd_player3_picture_buffer_size = 1470;
    let button_cd_player3_picture_buffer_start =
        button_cd_player3_picture_offset + button_cd_player3_picture_header_size;
    let button_cd_player3_picture_buffer_end =
        button_cd_player3_picture_buffer_start + button_cd_player3_picture_buffer_size;
    let button_cd_player3_picture_buffer = resource_file_bytes
        [button_cd_player3_picture_buffer_start..button_cd_player3_picture_buffer_end]
        .to_vec();

    let button_cd_player4_picture_offset = 0xEDAAF;
    let button_cd_player4_picture_header_size = 12;
    let button_cd_player4_picture_buffer_size = 1406;
    let button_cd_player4_picture_buffer_start =
        button_cd_player4_picture_offset + button_cd_player4_picture_header_size;
    let button_cd_player4_picture_buffer_end =
        button_cd_player4_picture_buffer_start + button_cd_player4_picture_buffer_size;
    let button_cd_player4_picture_buffer = resource_file_bytes
        [button_cd_player4_picture_buffer_start..button_cd_player4_picture_buffer_end]
        .to_vec();

    let button_cd_player5_picture_offset = 0xEE039;
    let button_cd_player5_picture_header_size = 12;
    let button_cd_player5_picture_buffer_size = 1406;
    let button_cd_player5_picture_buffer_start =
        button_cd_player5_picture_offset + button_cd_player5_picture_header_size;
    let button_cd_player5_picture_buffer_end =
        button_cd_player5_picture_buffer_start + button_cd_player5_picture_buffer_size;
    let button_cd_player5_picture_buffer = resource_file_bytes
        [button_cd_player5_picture_buffer_start..button_cd_player5_picture_buffer_end]
        .to_vec();

    let button_cd_player6_picture_offset = 0xEE5C3;
    let button_cd_player6_picture_header_size = 12;
    let button_cd_player6_picture_buffer_size = 1406;
    let button_cd_player6_picture_buffer_start =
        button_cd_player6_picture_offset + button_cd_player6_picture_header_size;
    let button_cd_player6_picture_buffer_end =
        button_cd_player6_picture_buffer_start + button_cd_player6_picture_buffer_size;
    let button_cd_player6_picture_buffer = resource_file_bytes
        [button_cd_player6_picture_buffer_start..button_cd_player6_picture_buffer_end]
        .to_vec();

    let button_cd_player7_picture_offset = 0xEEB4D;
    let button_cd_player7_picture_header_size = 12;
    let button_cd_player7_picture_buffer_size = 1470;
    let button_cd_player7_picture_buffer_start =
        button_cd_player7_picture_offset + button_cd_player7_picture_header_size;
    let button_cd_player7_picture_buffer_end =
        button_cd_player7_picture_buffer_start + button_cd_player7_picture_buffer_size;
    let button_cd_player7_picture_buffer = resource_file_bytes
        [button_cd_player7_picture_buffer_start..button_cd_player7_picture_buffer_end]
        .to_vec();

    let button_cd_player8_picture_offset = 0xEF117;
    let button_cd_player8_picture_header_size = 12;
    let button_cd_player8_picture_buffer_size = 1470;
    let button_cd_player8_picture_buffer_start =
        button_cd_player8_picture_offset + button_cd_player8_picture_header_size;
    let button_cd_player8_picture_buffer_end =
        button_cd_player8_picture_buffer_start + button_cd_player8_picture_buffer_size;
    let button_cd_player8_picture_buffer = resource_file_bytes
        [button_cd_player8_picture_buffer_start..button_cd_player8_picture_buffer_end]
        .to_vec();

    let main_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        main_icon_offset,
    ) {
        Ok(main_icon) => main_icon,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let element1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        element1_picture_offset,
    ) {
        Ok(element1_picture) => element1_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_display_list_image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_display_list_image1_picture_offset,
    ) {
        Ok(cd_display_list_image1_picture) => cd_display_list_image1_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_display_list_image2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_display_list_image2_picture_offset,
    ) {
        Ok(cd_display_list_image2_picture) => cd_display_list_image2_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_display_list_image3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_display_list_image3_picture_offset,
    ) {
        Ok(cd_display_list_image3_picture) => cd_display_list_image3_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_display_list_image4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_display_list_image4_picture_offset,
    ) {
        Ok(cd_display_list_image4_picture) => cd_display_list_image4_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_display_list_image5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_display_list_image5_picture_offset,
    ) {
        Ok(cd_display_list_image5_picture) => cd_display_list_image5_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_animation_list_image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_animation_list_image1_picture_offset,
    ) {
        Ok(cd_animation_list_image1_picture) => cd_animation_list_image1_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_animation_list_image2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_animation_list_image2_picture_offset,
    ) {
        Ok(cd_animation_list_image2_picture) => cd_animation_list_image2_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_animation_list_image3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_animation_list_image3_picture_offset,
    ) {
        Ok(cd_animation_list_image3_picture) => cd_animation_list_image3_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_animation_list_image4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_animation_list_image4_picture_offset,
    ) {
        Ok(cd_animation_list_image4_picture) => cd_animation_list_image4_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_animation_list_image5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_animation_list_image5_picture_offset,
    ) {
        Ok(cd_animation_list_image5_picture) => cd_animation_list_image5_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_animation_list_image6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_animation_list_image6_picture_offset,
    ) {
        Ok(cd_animation_list_image6_picture) => cd_animation_list_image6_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let cd_animation_list_image7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        cd_animation_list_image7_picture_offset,
    ) {
        Ok(cd_animation_list_image7_picture) => cd_animation_list_image7_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let element2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        element2_picture_offset,
    ) {
        Ok(element2_picture) => element2_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let switch_master_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        switch_master_glyph_offset,
    ) {
        Ok(switch_master_glyph) => switch_master_glyph,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let switch_rec_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        switch_rec_glyph_offset,
    ) {
        Ok(switch_rec_glyph) => switch_rec_glyph,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let switch_cd_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        switch_cd_glyph_offset,
    ) {
        Ok(switch_cd_glyph) => switch_cd_glyph,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let switch_dat_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        switch_dat_glyph_offset,
    ) {
        Ok(switch_dat_glyph) => switch_dat_glyph,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let switch_midi_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        switch_midi_glyph_offset,
    ) {
        Ok(switch_midi_glyph) => switch_midi_glyph,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        image1_picture_offset,
    ) {
        Ok(image1_picture) => image1_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let image2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        image2_picture_offset,
    ) {
        Ok(image2_picture) => image2_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let media_player_animation_list_image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        media_player_animation_list_image1_picture_offset,
    ) {
        Ok(media_player_animation_list_image1_picture) => {
            media_player_animation_list_image1_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let media_player_animation_list_image2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        media_player_animation_list_image2_picture_offset,
    ) {
        Ok(media_player_animation_list_image2_picture) => {
            media_player_animation_list_image2_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let media_player_animation_list_image3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        media_player_animation_list_image3_picture_offset,
    ) {
        Ok(media_player_animation_list_image3_picture) => {
            media_player_animation_list_image3_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let media_player_animation_list_image4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        media_player_animation_list_image4_picture_offset,
    ) {
        Ok(media_player_animation_list_image4_picture) => {
            media_player_animation_list_image4_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let media_player_animation_list_image5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        media_player_animation_list_image5_picture_offset,
    ) {
        Ok(media_player_animation_list_image5_picture) => {
            media_player_animation_list_image5_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let media_player_animation_list_image6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        media_player_animation_list_image6_picture_offset,
    ) {
        Ok(media_player_animation_list_image6_picture) => {
            media_player_animation_list_image6_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let media_player_animation_list_image7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        media_player_animation_list_image7_picture_offset,
    ) {
        Ok(media_player_animation_list_image7_picture) => {
            media_player_animation_list_image7_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let media_player_animation_list_image8_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        media_player_animation_list_image8_picture_offset,
    ) {
        Ok(media_player_animation_list_image8_picture) => {
            media_player_animation_list_image8_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let element3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        element3_picture_offset,
    ) {
        Ok(element3_picture) => element3_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let button_cd_player1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        button_cd_player1_picture_offset,
    ) {
        Ok(button_cd_player1_picture) => button_cd_player1_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let button_cd_player2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        button_cd_player2_picture_offset,
    ) {
        Ok(button_cd_player2_picture) => button_cd_player2_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let button_cd_player3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        button_cd_player3_picture_offset,
    ) {
        Ok(button_cd_player3_picture) => button_cd_player3_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let button_cd_player4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        button_cd_player4_picture_offset,
    ) {
        Ok(button_cd_player4_picture) => button_cd_player4_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let button_cd_player5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        button_cd_player5_picture_offset,
    ) {
        Ok(button_cd_player5_picture) => button_cd_player5_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let button_cd_player6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        button_cd_player6_picture_offset,
    ) {
        Ok(button_cd_player6_picture) => button_cd_player6_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let button_cd_player7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        button_cd_player7_picture_offset,
    ) {
        Ok(button_cd_player7_picture) => button_cd_player7_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    let button_cd_player8_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx".to_owned(),
        button_cd_player8_picture_offset,
    ) {
        Ok(button_cd_player8_picture) => button_cd_player8_picture,
        Err(e) => panic!("Failed to resolve resource file: {}", e),
    };

    assert_eq!(main_icon.len(), main_icon_buffer_size);
    assert_eq!(element1_picture.len(), element1_picture_buffer_size);
    assert_eq!(
        cd_display_list_image1_picture.len(),
        cd_display_list_image1_picture_buffer_size
    );
    assert_eq!(
        cd_display_list_image2_picture.len(),
        cd_display_list_image2_picture_buffer_size
    );
    assert_eq!(
        cd_display_list_image3_picture.len(),
        cd_display_list_image3_picture_buffer_size
    );
    assert_eq!(
        cd_display_list_image4_picture.len(),
        cd_display_list_image4_picture_buffer_size
    );
    assert_eq!(
        cd_display_list_image5_picture.len(),
        cd_display_list_image5_picture_buffer_size
    );
    assert_eq!(
        cd_animation_list_image1_picture.len(),
        cd_animation_list_image1_picture_buffer_size
    );
    assert_eq!(
        cd_animation_list_image2_picture.len(),
        cd_animation_list_image2_picture_buffer_size
    );
    assert_eq!(
        cd_animation_list_image3_picture.len(),
        cd_animation_list_image3_picture_buffer_size
    );
    assert_eq!(
        cd_animation_list_image4_picture.len(),
        cd_animation_list_image4_picture_buffer_size
    );
    assert_eq!(
        cd_animation_list_image5_picture.len(),
        cd_animation_list_image5_picture_buffer_size
    );
    assert_eq!(
        cd_animation_list_image6_picture.len(),
        cd_animation_list_image6_picture_buffer_size
    );
    assert_eq!(
        cd_animation_list_image7_picture.len(),
        cd_animation_list_image7_picture_buffer_size
    );
    assert_eq!(element2_picture.len(), element2_picture_buffer_size);
    assert_eq!(switch_master_glyph.len(), switch_master_glyph_buffer_size);
    assert_eq!(switch_rec_glyph.len(), switch_rec_glyph_buffer_size);
    assert_eq!(switch_cd_glyph.len(), switch_cd_glyph_buffer_size);
    assert_eq!(switch_dat_glyph.len(), switch_dat_glyph_buffer_size);
    assert_eq!(switch_midi_glyph.len(), switch_midi_glyph_buffer_size);
    assert_eq!(image1_picture.len(), image1_picture_buffer_size);
    assert_eq!(image2_picture.len(), image2_picture_buffer_size);
    assert_eq!(
        media_player_animation_list_image1_picture.len(),
        media_player_animation_list_image1_picture_buffer_size
    );
    assert_eq!(
        media_player_animation_list_image2_picture.len(),
        media_player_animation_list_image2_picture_buffer_size
    );
    assert_eq!(
        media_player_animation_list_image3_picture.len(),
        media_player_animation_list_image3_picture_buffer_size
    );
    assert_eq!(
        media_player_animation_list_image4_picture.len(),
        media_player_animation_list_image4_picture_buffer_size
    );
    assert_eq!(
        media_player_animation_list_image5_picture.len(),
        media_player_animation_list_image5_picture_buffer_size
    );
    assert_eq!(
        media_player_animation_list_image6_picture.len(),
        media_player_animation_list_image6_picture_buffer_size
    );
    assert_eq!(
        media_player_animation_list_image7_picture.len(),
        media_player_animation_list_image7_picture_buffer_size
    );
    assert_eq!(
        media_player_animation_list_image8_picture.len(),
        media_player_animation_list_image8_picture_buffer_size
    );
    assert_eq!(element3_picture.len(), element3_picture_buffer_size);
    assert_eq!(
        button_cd_player1_picture.len(),
        button_cd_player1_picture_buffer_size
    );
    assert_eq!(
        button_cd_player2_picture.len(),
        button_cd_player2_picture_buffer_size
    );
    assert_eq!(
        button_cd_player3_picture.len(),
        button_cd_player3_picture_buffer_size
    );
    assert_eq!(
        button_cd_player4_picture.len(),
        button_cd_player4_picture_buffer_size
    );
    assert_eq!(
        button_cd_player5_picture.len(),
        button_cd_player5_picture_buffer_size
    );
    assert_eq!(
        button_cd_player6_picture.len(),
        button_cd_player6_picture_buffer_size
    );
    assert_eq!(
        button_cd_player7_picture.len(),
        button_cd_player7_picture_buffer_size
    );
    assert_eq!(
        button_cd_player8_picture.len(),
        button_cd_player8_picture_buffer_size
    );

    assert_eq!(main_icon, main_icon_buffer);
    assert_eq!(element1_picture, element1_picture_buffer);
    assert_eq!(
        cd_display_list_image1_picture,
        cd_display_list_image1_picture_buffer
    );
    assert_eq!(
        cd_display_list_image2_picture,
        cd_display_list_image2_picture_buffer
    );
    assert_eq!(
        cd_display_list_image3_picture,
        cd_display_list_image3_picture_buffer
    );
    assert_eq!(
        cd_display_list_image4_picture,
        cd_display_list_image4_picture_buffer
    );
    assert_eq!(
        cd_display_list_image5_picture,
        cd_display_list_image5_picture_buffer
    );
    assert_eq!(
        cd_animation_list_image1_picture,
        cd_animation_list_image1_picture_buffer
    );
    assert_eq!(
        cd_animation_list_image2_picture,
        cd_animation_list_image2_picture_buffer
    );
    assert_eq!(
        cd_animation_list_image3_picture,
        cd_animation_list_image3_picture_buffer
    );
    assert_eq!(
        cd_animation_list_image4_picture,
        cd_animation_list_image4_picture_buffer
    );
    assert_eq!(
        cd_animation_list_image5_picture,
        cd_animation_list_image5_picture_buffer
    );
    assert_eq!(
        cd_animation_list_image6_picture,
        cd_animation_list_image6_picture_buffer
    );
    assert_eq!(
        cd_animation_list_image7_picture,
        cd_animation_list_image7_picture_buffer
    );
    assert_eq!(element2_picture, element2_picture_buffer);
    assert_eq!(switch_master_glyph, switch_master_glyph_buffer);
    assert_eq!(switch_rec_glyph, switch_rec_glyph_buffer);
    assert_eq!(switch_cd_glyph, switch_cd_glyph_buffer);
    assert_eq!(switch_dat_glyph, switch_dat_glyph_buffer);
    assert_eq!(switch_midi_glyph, switch_midi_glyph_buffer);
    assert_eq!(image1_picture, image1_picture_buffer);
    assert_eq!(image2_picture, image2_picture_buffer);
    assert_eq!(
        media_player_animation_list_image1_picture,
        media_player_animation_list_image1_picture_buffer
    );
    assert_eq!(
        media_player_animation_list_image2_picture,
        media_player_animation_list_image2_picture_buffer
    );
    assert_eq!(
        media_player_animation_list_image3_picture,
        media_player_animation_list_image3_picture_buffer
    );
    assert_eq!(
        media_player_animation_list_image4_picture,
        media_player_animation_list_image4_picture_buffer
    );
    assert_eq!(
        media_player_animation_list_image5_picture,
        media_player_animation_list_image5_picture_buffer
    );
    assert_eq!(
        media_player_animation_list_image6_picture,
        media_player_animation_list_image6_picture_buffer
    );
    assert_eq!(
        media_player_animation_list_image7_picture,
        media_player_animation_list_image7_picture_buffer
    );
    assert_eq!(
        media_player_animation_list_image8_picture,
        media_player_animation_list_image8_picture_buffer
    );
    assert_eq!(element3_picture, element3_picture_buffer);
    assert_eq!(button_cd_player1_picture, button_cd_player1_picture_buffer);
    assert_eq!(button_cd_player2_picture, button_cd_player2_picture_buffer);
    assert_eq!(button_cd_player3_picture, button_cd_player3_picture_buffer);
    assert_eq!(button_cd_player4_picture, button_cd_player4_picture_buffer);
    assert_eq!(button_cd_player5_picture, button_cd_player5_picture_buffer);
    assert_eq!(button_cd_player6_picture, button_cd_player6_picture_buffer);
    assert_eq!(button_cd_player7_picture, button_cd_player7_picture_buffer);
    assert_eq!(button_cd_player8_picture, button_cd_player8_picture_buffer);
}
