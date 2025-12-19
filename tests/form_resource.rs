use vb6parse::parsers::resource::list_resolver;

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
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx",
        about_icon_offset,
    ) {
        Ok(about_icon) => about_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let about_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx",
        about_picture_offset,
    ) {
        Ok(about_picture) => about_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let label6_caption = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx",
        label6_caption_offset,
    ) {
        Ok(label6_caption) => label6_caption,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let label4_caption = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_About.frx",
        label4_caption_offset,
    ) {
        Ok(label4_caption) => label4_caption,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
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
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Busy.frx",
        busy_icon_offset,
    ) {
        Ok(busy_icon) => busy_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
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
    let alpha_img_ctl2_image_buffer_size = 113_643;
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
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx",
        init_icon_offset,
    ) {
        Ok(init_icon) => init_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let alpha_img_ctl2_image = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx",
        alpha_img_ctl2_image_offset,
    ) {
        Ok(alpha_img_ctl2_image) => alpha_img_ctl2_image,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let alpha_img_ctl2_effects = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx",
        alpha_img_ctl2_effects_offset,
    ) {
        Ok(alpha_img_ctl2_effects) => alpha_img_ctl2_effects,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let alpha_img_ctl1_image = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx",
        alpha_img_ctl1_image_offset,
    ) {
        Ok(alpha_img_ctl1_image) => alpha_img_ctl1_image,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let alpha_img_ctl1_effects = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Init.frx",
        alpha_img_ctl1_effects_offset,
    ) {
        Ok(alpha_img_ctl1_effects) => alpha_img_ctl1_effects,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
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
    let element2_picture_buffer_size = 203_150;
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

    let image3_picture_offset = 0xBB2CF;
    let image3_picture_header_size = 12;
    let image3_picture_buffer_size = 4878;
    let image3_picture_buffer_start = image3_picture_offset + image3_picture_header_size;
    let image3_picture_buffer_end = image3_picture_buffer_start + image3_picture_buffer_size;
    let image3_picture_buffer =
        resource_file_bytes[image3_picture_buffer_start..image3_picture_buffer_end].to_vec();

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
    let element3_picture_buffer_size = 193_570;
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

    let light_cd_play_on_picture_offset = 0xEF681;
    let light_cd_play_on_picture_header_size = 12;
    let light_cd_play_on_picture_buffer_size = 0;
    let light_cd_play_on_picture_buffer_start =
        light_cd_play_on_picture_offset + light_cd_play_on_picture_header_size;
    let light_cd_play_on_picture_buffer_end =
        light_cd_play_on_picture_buffer_start + light_cd_play_on_picture_buffer_size;
    let light_cd_play_on_picture_buffer = resource_file_bytes
        [light_cd_play_on_picture_buffer_start..light_cd_play_on_picture_buffer_end]
        .to_vec();

    let light_cd_pause_on_picture_offset = 0xEFBBB;
    let light_cd_pause_on_picture_header_size = 12;
    let light_cd_pause_on_picture_buffer_size = 1246;
    let light_cd_pause_on_picture_buffer_start =
        light_cd_pause_on_picture_offset + light_cd_pause_on_picture_header_size;
    let light_cd_pause_on_picture_buffer_end =
        light_cd_pause_on_picture_buffer_start + light_cd_pause_on_picture_buffer_size;
    let light_cd_pause_on_picture_buffer = resource_file_bytes
        [light_cd_pause_on_picture_buffer_start..light_cd_pause_on_picture_buffer_end]
        .to_vec();

    let element4_picture_offset = 0xF00A5;
    let element4_picture_header_size = 12;
    let element4_picture_buffer_size = 193_570;
    let element4_picture_buffer_start = element4_picture_offset + element4_picture_header_size;
    let element4_picture_buffer_end = element4_picture_buffer_start + element4_picture_buffer_size;
    let element4_picture_buffer =
        resource_file_bytes[element4_picture_buffer_start..element4_picture_buffer_end].to_vec();

    let button_open_stream_picture_offset = 0x11_F4D3;
    let button_open_stream_picture_header_size = 12;
    let button_open_stream_picture_buffer_size = 446;
    let button_open_stream_picture_buffer_start =
        button_open_stream_picture_offset + button_open_stream_picture_header_size;
    let button_open_stream_picture_buffer_end =
        button_open_stream_picture_buffer_start + button_open_stream_picture_buffer_size;
    let button_open_stream_picture_buffer = resource_file_bytes
        [button_open_stream_picture_buffer_start..button_open_stream_picture_buffer_end]
        .to_vec();

    let button_stop_stream_picture_offset = 0x11_F69D;
    let button_stop_stream_picture_header_size = 12;
    let button_stop_stream_picture_buffer_size = 1406;
    let button_stop_stream_picture_buffer_start =
        button_stop_stream_picture_offset + button_stop_stream_picture_header_size;
    let button_stop_stream_picture_buffer_end =
        button_stop_stream_picture_buffer_start + button_stop_stream_picture_buffer_size;
    let button_stop_stream_picture_buffer = resource_file_bytes
        [button_stop_stream_picture_buffer_start..button_stop_stream_picture_buffer_end]
        .to_vec();

    let button_play_stream_picture_offset = 0x11_FC27;
    let button_play_stream_picture_header_size = 12;
    let button_play_stream_picture_buffer_size = 1406;
    let button_play_stream_picture_buffer_start =
        button_play_stream_picture_offset + button_play_stream_picture_header_size;
    let button_play_stream_picture_buffer_end =
        button_play_stream_picture_buffer_start + button_play_stream_picture_buffer_size;
    let button_play_stream_picture_buffer = resource_file_bytes
        [button_play_stream_picture_buffer_start..button_play_stream_picture_buffer_end]
        .to_vec();

    let element5_picture_offset = 0x12_01B1;
    let element5_picture_header_size = 12;
    let element5_picture_buffer_size = 193_570;
    let element5_picture_buffer_start = element5_picture_offset + element5_picture_header_size;
    let element5_picture_buffer_end = element5_picture_buffer_start + element5_picture_buffer_size;
    let element5_picture_buffer =
        resource_file_bytes[element5_picture_buffer_start..element5_picture_buffer_end].to_vec();

    let recording_picture_offset = 0x14_F5DF;
    let recording_picture_header_size = 12;
    let recording_picture_buffer_size = 1406;
    let recording_picture_buffer_start = recording_picture_offset + recording_picture_header_size;
    let recording_picture_buffer_end =
        recording_picture_buffer_start + recording_picture_buffer_size;
    let recording_picture_buffer =
        resource_file_bytes[recording_picture_buffer_start..recording_picture_buffer_end].to_vec();

    let image4_picture_offset = 0x14_FB69;
    let image4_picture_header_size = 12;
    let image4_picture_buffer_size = 1406;
    let image4_picture_buffer_start = image4_picture_offset + image4_picture_header_size;
    let image4_picture_buffer_end = image4_picture_buffer_start + image4_picture_buffer_size;
    let image4_picture_buffer =
        resource_file_bytes[image4_picture_buffer_start..image4_picture_buffer_end].to_vec();

    let picture17_picture_offset = 0x15_00F3;
    let picture17_picture_header_size = 12;
    let picture17_picture_buffer_size = 36942;
    let picture17_picture_buffer_start = picture17_picture_offset + picture17_picture_header_size;
    let picture17_picture_buffer_end =
        picture17_picture_buffer_start + picture17_picture_buffer_size;
    let picture17_picture_buffer =
        resource_file_bytes[picture17_picture_buffer_start..picture17_picture_buffer_end].to_vec();

    let cmd_audioplayer1_picture_offset = 0x159_14D;
    let cmd_audioplayer1_picture_header_size = 12;
    let cmd_audioplayer1_picture_buffer_size = 1406;
    let cmd_audioplayer1_picture_buffer_start =
        cmd_audioplayer1_picture_offset + cmd_audioplayer1_picture_header_size;
    let cmd_audioplayer1_picture_buffer_end =
        cmd_audioplayer1_picture_buffer_start + cmd_audioplayer1_picture_buffer_size;
    let cmd_audioplayer1_picture_buffer = resource_file_bytes
        [cmd_audioplayer1_picture_buffer_start..cmd_audioplayer1_picture_buffer_end]
        .to_vec();

    let cmd_audioplayer2_picture_offset = 0x159_6D7;
    let cmd_audioplayer2_picture_header_size = 12;
    let cmd_audioplayer2_picture_buffer_size = 1470;
    let cmd_audioplayer2_picture_buffer_start =
        cmd_audioplayer2_picture_offset + cmd_audioplayer2_picture_header_size;
    let cmd_audioplayer2_picture_buffer_end =
        cmd_audioplayer2_picture_buffer_start + cmd_audioplayer2_picture_buffer_size;
    let cmd_audioplayer2_picture_buffer = resource_file_bytes
        [cmd_audioplayer2_picture_buffer_start..cmd_audioplayer2_picture_buffer_end]
        .to_vec();

    let cmd_audioplayer3_picture_offset = 0x159_CA1;
    let cmd_audioplayer3_picture_header_size = 12;
    let cmd_audioplayer3_picture_buffer_size = 1470;
    let cmd_audioplayer3_picture_buffer_start =
        cmd_audioplayer3_picture_offset + cmd_audioplayer3_picture_header_size;
    let cmd_audioplayer3_picture_buffer_end =
        cmd_audioplayer3_picture_buffer_start + cmd_audioplayer3_picture_buffer_size;
    let cmd_audioplayer3_picture_buffer = resource_file_bytes
        [cmd_audioplayer3_picture_buffer_start..cmd_audioplayer3_picture_buffer_end]
        .to_vec();

    let cmd_audioplayer4_picture_offset = 0x15A_26B;
    let cmd_audioplayer4_picture_header_size = 12;
    let cmd_audioplayer4_picture_buffer_size = 1406;
    let cmd_audioplayer4_picture_buffer_start =
        cmd_audioplayer4_picture_offset + cmd_audioplayer4_picture_header_size;
    let cmd_audioplayer4_picture_buffer_end =
        cmd_audioplayer4_picture_buffer_start + cmd_audioplayer4_picture_buffer_size;
    let cmd_audioplayer4_picture_buffer = resource_file_bytes
        [cmd_audioplayer4_picture_buffer_start..cmd_audioplayer4_picture_buffer_end]
        .to_vec();

    let cmd_audioplayer5_picture_offset = 0x15A_7F5;
    let cmd_audioplayer5_picture_header_size = 12;
    let cmd_audioplayer5_picture_buffer_size = 1406;
    let cmd_audioplayer5_picture_buffer_start =
        cmd_audioplayer5_picture_offset + cmd_audioplayer5_picture_header_size;
    let cmd_audioplayer5_picture_buffer_end =
        cmd_audioplayer5_picture_buffer_start + cmd_audioplayer5_picture_buffer_size;
    let cmd_audioplayer5_picture_buffer = resource_file_bytes
        [cmd_audioplayer5_picture_buffer_start..cmd_audioplayer5_picture_buffer_end]
        .to_vec();

    let cmd_audioplayer6_picture_offset = 0x15_AD7F;
    let cmd_audioplayer6_picture_header_size = 12;
    let cmd_audioplayer6_picture_buffer_size = 1406;
    let cmd_audioplayer6_picture_buffer_start =
        cmd_audioplayer6_picture_offset + cmd_audioplayer6_picture_header_size;
    let cmd_audioplayer6_picture_buffer_end =
        cmd_audioplayer6_picture_buffer_start + cmd_audioplayer6_picture_buffer_size;
    let cmd_audioplayer6_picture_buffer = resource_file_bytes
        [cmd_audioplayer6_picture_buffer_start..cmd_audioplayer6_picture_buffer_end]
        .to_vec();

    let cmd_audioplayer7_picture_offset = 0x15B_309;
    let cmd_audioplayer7_picture_header_size = 12;
    let cmd_audioplayer7_picture_buffer_size = 1470;
    let cmd_audioplayer7_picture_buffer_start =
        cmd_audioplayer7_picture_offset + cmd_audioplayer7_picture_header_size;
    let cmd_audioplayer7_picture_buffer_end =
        cmd_audioplayer7_picture_buffer_start + cmd_audioplayer7_picture_buffer_size;
    let cmd_audioplayer7_picture_buffer = resource_file_bytes
        [cmd_audioplayer7_picture_buffer_start..cmd_audioplayer7_picture_buffer_end]
        .to_vec();

    let cmd_audioplayer8_picture_offset = 0x15B_8D3;
    let cmd_audioplayer8_picture_header_size = 12;
    let cmd_audioplayer8_picture_buffer_size = 1470;
    let cmd_audioplayer8_picture_buffer_start =
        cmd_audioplayer8_picture_offset + cmd_audioplayer8_picture_header_size;
    let cmd_audioplayer8_picture_buffer_end =
        cmd_audioplayer8_picture_buffer_start + cmd_audioplayer8_picture_buffer_size;
    let cmd_audioplayer8_picture_buffer = resource_file_bytes
        [cmd_audioplayer8_picture_buffer_start..cmd_audioplayer8_picture_buffer_end]
        .to_vec();

    let cmd_audioplayer9_picture_offset = 0x15B_E9D;
    let cmd_audioplayer9_picture_header_size = 12;
    let cmd_audioplayer9_picture_buffer_size = 1406;
    let cmd_audioplayer9_picture_buffer_start =
        cmd_audioplayer9_picture_offset + cmd_audioplayer9_picture_header_size;
    let cmd_audioplayer9_picture_buffer_end =
        cmd_audioplayer9_picture_buffer_start + cmd_audioplayer9_picture_buffer_size;
    let cmd_audioplayer9_picture_buffer = resource_file_bytes
        [cmd_audioplayer9_picture_buffer_start..cmd_audioplayer9_picture_buffer_end]
        .to_vec();

    let light_dat_play_on_picture_offset = 0x15C_427;
    let light_dat_play_on_picture_header_size = 12;
    let light_dat_play_on_picture_buffer_size = 1230;
    let light_dat_play_on_picture_buffer_start =
        light_dat_play_on_picture_offset + light_dat_play_on_picture_header_size;
    let light_dat_play_on_picture_buffer_end =
        light_dat_play_on_picture_buffer_start + light_dat_play_on_picture_buffer_size;
    let light_dat_play_on_picture_buffer = resource_file_bytes
        [light_dat_play_on_picture_buffer_start..light_dat_play_on_picture_buffer_end]
        .to_vec();

    let light_dat_pause_on_picture_offset = 0x15C_901;
    let light_dat_pause_on_picture_header_size = 12;
    let light_dat_pause_on_picture_buffer_size = 1246;
    let light_dat_pause_on_picture_buffer_start =
        light_dat_pause_on_picture_offset + light_dat_pause_on_picture_header_size;
    let light_dat_pause_on_picture_buffer_end =
        light_dat_pause_on_picture_buffer_start + light_dat_pause_on_picture_buffer_size;
    let light_dat_pause_on_picture_buffer = resource_file_bytes
        [light_dat_pause_on_picture_buffer_start..light_dat_pause_on_picture_buffer_end]
        .to_vec();

    let element6_picture_offset = 0x15C_DEB;
    let element6_picture_header_size = 12;
    let element6_picture_buffer_size = 195486;
    let element6_picture_buffer_start = element6_picture_offset + element6_picture_header_size;
    let element6_picture_buffer_end = element6_picture_buffer_start + element6_picture_buffer_size;
    let element6_picture_buffer =
        resource_file_bytes[element6_picture_buffer_start..element6_picture_buffer_end].to_vec();

    let button_midi_player1_picture_offset = 0x18C_995;
    let button_midi_player1_picture_header_size = 12;
    let button_midi_player1_picture_buffer_size = 1470;
    let button_midi_player1_picture_buffer_start =
        button_midi_player1_picture_offset + button_midi_player1_picture_header_size;
    let button_midi_player1_picture_buffer_end =
        button_midi_player1_picture_buffer_start + button_midi_player1_picture_buffer_size;
    let button_midi_player1_picture_buffer = resource_file_bytes
        [button_midi_player1_picture_buffer_start..button_midi_player1_picture_buffer_end]
        .to_vec();

    let button_midi_player2_picture_offset = 0x18C_F5F;
    let button_midi_player2_picture_header_size = 12;
    let button_midi_player2_picture_buffer_size = 1406;
    let button_midi_player2_picture_buffer_start =
        button_midi_player2_picture_offset + button_midi_player2_picture_header_size;
    let button_midi_player2_picture_buffer_end =
        button_midi_player2_picture_buffer_start + button_midi_player2_picture_buffer_size;
    let button_midi_player2_picture_buffer = resource_file_bytes
        [button_midi_player2_picture_buffer_start..button_midi_player2_picture_buffer_end]
        .to_vec();

    let button_midi_player3_picture_offset = 0x18D_4E9;
    let button_midi_player3_picture_header_size = 12;
    let button_midi_player3_picture_buffer_size = 1470;
    let button_midi_player3_picture_buffer_start =
        button_midi_player3_picture_offset + button_midi_player3_picture_header_size;
    let button_midi_player3_picture_buffer_end =
        button_midi_player3_picture_buffer_start + button_midi_player3_picture_buffer_size;
    let button_midi_player3_picture_buffer = resource_file_bytes
        [button_midi_player3_picture_buffer_start..button_midi_player3_picture_buffer_end]
        .to_vec();

    let button_midi_player4_picture_offset = 0x18D_AB3;
    let button_midi_player4_picture_header_size = 12;
    let button_midi_player4_picture_buffer_size = 1470;
    let button_midi_player4_picture_buffer_start =
        button_midi_player4_picture_offset + button_midi_player4_picture_header_size;
    let button_midi_player4_picture_buffer_end =
        button_midi_player4_picture_buffer_start + button_midi_player4_picture_buffer_size;
    let button_midi_player4_picture_buffer = resource_file_bytes
        [button_midi_player4_picture_buffer_start..button_midi_player4_picture_buffer_end]
        .to_vec();

    let button_midi_player5_picture_offset = 0x18E_07D;
    let button_midi_player5_picture_header_size = 12;
    let button_midi_player5_picture_buffer_size = 1406;
    let button_midi_player5_picture_buffer_start =
        button_midi_player5_picture_offset + button_midi_player5_picture_header_size;
    let button_midi_player5_picture_buffer_end =
        button_midi_player5_picture_buffer_start + button_midi_player5_picture_buffer_size;
    let button_midi_player5_picture_buffer = resource_file_bytes
        [button_midi_player5_picture_buffer_start..button_midi_player5_picture_buffer_end]
        .to_vec();

    let button_midi_player6_picture_offset = 0x18E_607;
    let button_midi_player6_picture_header_size = 12;
    let button_midi_player6_picture_buffer_size = 1406;
    let button_midi_player6_picture_buffer_start =
        button_midi_player6_picture_offset + button_midi_player6_picture_header_size;
    let button_midi_player6_picture_buffer_end =
        button_midi_player6_picture_buffer_start + button_midi_player6_picture_buffer_size;
    let button_midi_player6_picture_buffer = resource_file_bytes
        [button_midi_player6_picture_buffer_start..button_midi_player6_picture_buffer_end]
        .to_vec();

    let button_midi_player7_picture_offset = 0x18E_B91;
    let button_midi_player7_picture_header_size = 12;
    let button_midi_player7_picture_buffer_size = 1470;
    let button_midi_player7_picture_buffer_start =
        button_midi_player7_picture_offset + button_midi_player7_picture_header_size;
    let button_midi_player7_picture_buffer_end =
        button_midi_player7_picture_buffer_start + button_midi_player7_picture_buffer_size;
    let button_midi_player7_picture_buffer = resource_file_bytes
        [button_midi_player7_picture_buffer_start..button_midi_player7_picture_buffer_end]
        .to_vec();

    let light_midi_floppy_drive_picture_offset = 0x18F_15B;
    let light_midi_floppy_drive_picture_header_size = 12;
    let light_midi_floppy_drive_picture_buffer_size = 1654;
    let light_midi_floppy_drive_picture_buffer_start =
        light_midi_floppy_drive_picture_offset + light_midi_floppy_drive_picture_header_size;
    let light_midi_floppy_drive_picture_buffer_end =
        light_midi_floppy_drive_picture_buffer_start + light_midi_floppy_drive_picture_buffer_size;
    let light_midi_floppy_drive_picture_buffer = resource_file_bytes
        [light_midi_floppy_drive_picture_buffer_start..light_midi_floppy_drive_picture_buffer_end]
        .to_vec();

    let image6_picture_offset = 0x18F_7DD;
    let image6_picture_header_size = 12;
    let image6_picture_buffer_size = 626;
    let image6_picture_buffer_start = image6_picture_offset + image6_picture_header_size;
    let image6_picture_buffer_end = image6_picture_buffer_start + image6_picture_buffer_size;
    let image6_picture_buffer =
        resource_file_bytes[image6_picture_buffer_start..image6_picture_buffer_end].to_vec();

    let floppy_in_picture_offset = 0x18F_A5B;
    let floppy_in_picture_header_size = 12;
    let floppy_in_picture_buffer_size = 11662;
    let floppy_in_picture_buffer_start = floppy_in_picture_offset + floppy_in_picture_header_size;
    let floppy_in_picture_buffer_end =
        floppy_in_picture_buffer_start + floppy_in_picture_buffer_size;
    let floppy_in_picture_buffer =
        resource_file_bytes[floppy_in_picture_buffer_start..floppy_in_picture_buffer_end].to_vec();

    let floppy_out_picture_offset = 0x192_7F5;
    let floppy_out_picture_header_size = 12;
    let floppy_out_picture_buffer_size = 11662;
    let floppy_out_picture_buffer_start =
        floppy_out_picture_offset + floppy_out_picture_header_size;
    let floppy_out_picture_buffer_end =
        floppy_out_picture_buffer_start + floppy_out_picture_buffer_size;
    let floppy_out_picture_buffer = resource_file_bytes
        [floppy_out_picture_buffer_start..floppy_out_picture_buffer_end]
        .to_vec();

    let light_midi_play_on_picture_offset = 0x195_58F;
    let light_midi_play_on_picture_header_size = 12;
    let light_midi_play_on_picture_buffer_size = 1230;
    let light_midi_play_on_picture_buffer_start =
        light_midi_play_on_picture_offset + light_midi_play_on_picture_header_size;
    let light_midi_play_on_picture_buffer_end =
        light_midi_play_on_picture_buffer_start + light_midi_play_on_picture_buffer_size;
    let light_midi_play_on_picture_buffer = resource_file_bytes
        [light_midi_play_on_picture_buffer_start..light_midi_play_on_picture_buffer_end]
        .to_vec();

    let light_midi_pause_on_picture_offset = 0x195_A69;
    let light_midi_pause_on_picture_header_size = 12;
    let light_midi_pause_on_picture_buffer_size = 1246;
    let light_midi_pause_on_picture_buffer_start =
        light_midi_pause_on_picture_offset + light_midi_pause_on_picture_header_size;
    let light_midi_pause_on_picture_buffer_end =
        light_midi_pause_on_picture_buffer_start + light_midi_pause_on_picture_buffer_size;
    let light_midi_pause_on_picture_buffer = resource_file_bytes
        [light_midi_pause_on_picture_buffer_start..light_midi_pause_on_picture_buffer_end]
        .to_vec();

    let element7_picture_offset = 0x195_F53;
    let element7_picture_header_size = 12;
    let element7_picture_buffer_size = 93938;
    let element7_picture_buffer_start = element7_picture_offset + element7_picture_header_size;
    let element7_picture_buffer_end = element7_picture_buffer_start + element7_picture_buffer_size;
    let element7_picture_buffer =
        resource_file_bytes[element7_picture_buffer_start..element7_picture_buffer_end].to_vec();

    let options_menu_button_picture_offset = 0x1AC_E51;
    let options_menu_button_picture_header_size = 12;
    let options_menu_button_picture_buffer_size = 2106;
    let options_menu_button_picture_buffer_start =
        options_menu_button_picture_offset + options_menu_button_picture_header_size;
    let options_menu_button_picture_buffer_end =
        options_menu_button_picture_buffer_start + options_menu_button_picture_buffer_size;
    let options_menu_button_picture_buffer = resource_file_bytes
        [options_menu_button_picture_buffer_start..options_menu_button_picture_buffer_end]
        .to_vec();

    let image5_picture_offset = 0x1AD_697;
    let image5_picture_header_size = 12;
    let image5_picture_buffer_size = 19134;
    let image5_picture_buffer_start = image5_picture_offset + image5_picture_header_size;
    let image5_picture_buffer_end = image5_picture_buffer_start + image5_picture_buffer_size;
    let image5_picture_buffer =
        resource_file_bytes[image5_picture_buffer_start..image5_picture_buffer_end].to_vec();

    let elements_disabled_image_list_image1_picture_offset = 0x1B2_161;
    let elements_disabled_image_list_image1_picture_header_size = 4;
    let elements_disabled_image_list_image1_picture_buffer_size = 92046;
    let elements_disabled_image_list_image1_picture_buffer_start =
        elements_disabled_image_list_image1_picture_offset
            + elements_disabled_image_list_image1_picture_header_size;
    let elements_disabled_image_list_image1_picture_buffer_end =
        elements_disabled_image_list_image1_picture_buffer_start
            + elements_disabled_image_list_image1_picture_buffer_size;
    let elements_disabled_image_list_image1_picture_buffer = resource_file_bytes
        [elements_disabled_image_list_image1_picture_buffer_start
            ..elements_disabled_image_list_image1_picture_buffer_end]
        .to_vec();

    let elements_disabled_image_list_image2_picture_offset = 0x1C8_8F3;
    let elements_disabled_image_list_image2_picture_header_size = 4;
    let elements_disabled_image_list_image2_picture_buffer_size = 191_678;
    let elements_disabled_image_list_image2_picture_buffer_start =
        elements_disabled_image_list_image2_picture_offset
            + elements_disabled_image_list_image2_picture_header_size;
    let elements_disabled_image_list_image2_picture_buffer_end =
        elements_disabled_image_list_image2_picture_buffer_start
            + elements_disabled_image_list_image2_picture_buffer_size;
    let elements_disabled_image_list_image2_picture_buffer = resource_file_bytes
        [elements_disabled_image_list_image2_picture_buffer_start
            ..elements_disabled_image_list_image2_picture_buffer_end]
        .to_vec();

    let elements_disabled_image_list_image3_picture_offset = 0x1F7_5B5;
    let elements_disabled_image_list_image3_picture_header_size = 4;
    let elements_disabled_image_list_image3_picture_buffer_size = 191_678;
    let elements_disabled_image_list_image3_picture_buffer_start =
        elements_disabled_image_list_image3_picture_offset
            + elements_disabled_image_list_image3_picture_header_size;
    let elements_disabled_image_list_image3_picture_buffer_end =
        elements_disabled_image_list_image3_picture_buffer_start
            + elements_disabled_image_list_image3_picture_buffer_size;
    let elements_disabled_image_list_image3_picture_buffer = resource_file_bytes
        [elements_disabled_image_list_image3_picture_buffer_start
            ..elements_disabled_image_list_image3_picture_buffer_end]
        .to_vec();

    let elements_disabled_image_list_image4_picture_offset = 0x226_277;
    let elements_disabled_image_list_image4_picture_header_size = 4;
    let elements_disabled_image_list_image4_picture_buffer_size = 191_678;
    let elements_disabled_image_list_image4_picture_buffer_start =
        elements_disabled_image_list_image4_picture_offset
            + elements_disabled_image_list_image4_picture_header_size;
    let elements_disabled_image_list_image4_picture_buffer_end =
        elements_disabled_image_list_image4_picture_buffer_start
            + elements_disabled_image_list_image4_picture_buffer_size;
    let elements_disabled_image_list_image4_picture_buffer = resource_file_bytes
        [elements_disabled_image_list_image4_picture_buffer_start
            ..elements_disabled_image_list_image4_picture_buffer_end]
        .to_vec();

    let elements_disabled_image_list_image5_picture_offset = 0x254_F39;
    let elements_disabled_image_list_image5_picture_header_size = 4;
    let elements_disabled_image_list_image5_picture_buffer_size = 390_942;
    let elements_disabled_image_list_image5_picture_buffer_start =
        elements_disabled_image_list_image5_picture_offset
            + elements_disabled_image_list_image5_picture_header_size;
    let elements_disabled_image_list_image5_picture_buffer_end =
        elements_disabled_image_list_image5_picture_buffer_start
            + elements_disabled_image_list_image5_picture_buffer_size;
    let elements_disabled_image_list_image5_picture_buffer = resource_file_bytes
        [elements_disabled_image_list_image5_picture_buffer_start
            ..elements_disabled_image_list_image5_picture_buffer_end]
        .to_vec();

    let elements_disabled_image_list_image6_picture_offset = 0x2B4_65B;
    let elements_disabled_image_list_image6_picture_header_size = 4;
    let elements_disabled_image_list_image6_picture_buffer_size = 390_942;
    let elements_disabled_image_list_image6_picture_buffer_start =
        elements_disabled_image_list_image6_picture_offset
            + elements_disabled_image_list_image6_picture_header_size;
    let elements_disabled_image_list_image6_picture_buffer_end =
        elements_disabled_image_list_image6_picture_buffer_start
            + elements_disabled_image_list_image6_picture_buffer_size;
    let elements_disabled_image_list_image6_picture_buffer = resource_file_bytes
        [elements_disabled_image_list_image6_picture_buffer_start
            ..elements_disabled_image_list_image6_picture_buffer_end]
        .to_vec();

    let elements_disabled_image_list_image7_picture_offset = 0x313_D7D;
    let elements_disabled_image_list_image7_picture_header_size = 4;
    let elements_disabled_image_list_image7_picture_buffer_size = 92046;
    let elements_disabled_image_list_image7_picture_buffer_start =
        elements_disabled_image_list_image7_picture_offset
            + elements_disabled_image_list_image7_picture_header_size;
    let elements_disabled_image_list_image7_picture_buffer_end =
        elements_disabled_image_list_image7_picture_buffer_start
            + elements_disabled_image_list_image7_picture_buffer_size;
    let elements_disabled_image_list_image7_picture_buffer = resource_file_bytes
        [elements_disabled_image_list_image7_picture_buffer_start
            ..elements_disabled_image_list_image7_picture_buffer_end]
        .to_vec();

    let main_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        main_icon_offset,
    ) {
        Ok(main_icon) => main_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let element1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        element1_picture_offset,
    ) {
        Ok(element1_picture) => element1_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_display_list_image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_display_list_image1_picture_offset,
    ) {
        Ok(cd_display_list_image1_picture) => cd_display_list_image1_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_display_list_image2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_display_list_image2_picture_offset,
    ) {
        Ok(cd_display_list_image2_picture) => cd_display_list_image2_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_display_list_image3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_display_list_image3_picture_offset,
    ) {
        Ok(cd_display_list_image3_picture) => cd_display_list_image3_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_display_list_image4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_display_list_image4_picture_offset,
    ) {
        Ok(cd_display_list_image4_picture) => cd_display_list_image4_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_display_list_image5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_display_list_image5_picture_offset,
    ) {
        Ok(cd_display_list_image5_picture) => cd_display_list_image5_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_animation_list_image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_animation_list_image1_picture_offset,
    ) {
        Ok(cd_animation_list_image1_picture) => cd_animation_list_image1_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_animation_list_image2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_animation_list_image2_picture_offset,
    ) {
        Ok(cd_animation_list_image2_picture) => cd_animation_list_image2_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_animation_list_image3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_animation_list_image3_picture_offset,
    ) {
        Ok(cd_animation_list_image3_picture) => cd_animation_list_image3_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_animation_list_image4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_animation_list_image4_picture_offset,
    ) {
        Ok(cd_animation_list_image4_picture) => cd_animation_list_image4_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_animation_list_image5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_animation_list_image5_picture_offset,
    ) {
        Ok(cd_animation_list_image5_picture) => cd_animation_list_image5_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_animation_list_image6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_animation_list_image6_picture_offset,
    ) {
        Ok(cd_animation_list_image6_picture) => cd_animation_list_image6_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cd_animation_list_image7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cd_animation_list_image7_picture_offset,
    ) {
        Ok(cd_animation_list_image7_picture) => cd_animation_list_image7_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let element2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        element2_picture_offset,
    ) {
        Ok(element2_picture) => element2_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let switch_master_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        switch_master_glyph_offset,
    ) {
        Ok(switch_master_glyph) => switch_master_glyph,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let switch_rec_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        switch_rec_glyph_offset,
    ) {
        Ok(switch_rec_glyph) => switch_rec_glyph,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let switch_cd_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        switch_cd_glyph_offset,
    ) {
        Ok(switch_cd_glyph) => switch_cd_glyph,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let switch_dat_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        switch_dat_glyph_offset,
    ) {
        Ok(switch_dat_glyph) => switch_dat_glyph,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let switch_midi_glyph = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        switch_midi_glyph_offset,
    ) {
        Ok(switch_midi_glyph) => switch_midi_glyph,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        image1_picture_offset,
    ) {
        Ok(image1_picture) => image1_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let image3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        image3_picture_offset,
    ) {
        Ok(image3_picture) => image3_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let media_player_animation_list_image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        media_player_animation_list_image1_picture_offset,
    ) {
        Ok(media_player_animation_list_image1_picture) => {
            media_player_animation_list_image1_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let media_player_animation_list_image2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        media_player_animation_list_image2_picture_offset,
    ) {
        Ok(media_player_animation_list_image2_picture) => {
            media_player_animation_list_image2_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let media_player_animation_list_image3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        media_player_animation_list_image3_picture_offset,
    ) {
        Ok(media_player_animation_list_image3_picture) => {
            media_player_animation_list_image3_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let media_player_animation_list_image4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        media_player_animation_list_image4_picture_offset,
    ) {
        Ok(media_player_animation_list_image4_picture) => {
            media_player_animation_list_image4_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let media_player_animation_list_image5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        media_player_animation_list_image5_picture_offset,
    ) {
        Ok(media_player_animation_list_image5_picture) => {
            media_player_animation_list_image5_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let media_player_animation_list_image6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        media_player_animation_list_image6_picture_offset,
    ) {
        Ok(media_player_animation_list_image6_picture) => {
            media_player_animation_list_image6_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let media_player_animation_list_image7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        media_player_animation_list_image7_picture_offset,
    ) {
        Ok(media_player_animation_list_image7_picture) => {
            media_player_animation_list_image7_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let media_player_animation_list_image8_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        media_player_animation_list_image8_picture_offset,
    ) {
        Ok(media_player_animation_list_image8_picture) => {
            media_player_animation_list_image8_picture
        }
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let element3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        element3_picture_offset,
    ) {
        Ok(element3_picture) => element3_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_cd_player1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_cd_player1_picture_offset,
    ) {
        Ok(button_cd_player1_picture) => button_cd_player1_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_cd_player2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_cd_player2_picture_offset,
    ) {
        Ok(button_cd_player2_picture) => button_cd_player2_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_cd_player3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_cd_player3_picture_offset,
    ) {
        Ok(button_cd_player3_picture) => button_cd_player3_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_cd_player4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_cd_player4_picture_offset,
    ) {
        Ok(button_cd_player4_picture) => button_cd_player4_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_cd_player5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_cd_player5_picture_offset,
    ) {
        Ok(button_cd_player5_picture) => button_cd_player5_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_cd_player6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_cd_player6_picture_offset,
    ) {
        Ok(button_cd_player6_picture) => button_cd_player6_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_cd_player7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_cd_player7_picture_offset,
    ) {
        Ok(button_cd_player7_picture) => button_cd_player7_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_cd_player8_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_cd_player8_picture_offset,
    ) {
        Ok(button_cd_player8_picture) => button_cd_player8_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let light_cd_play_on_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        light_cd_play_on_picture_offset,
    ) {
        Ok(light_cd_play_on_picture) => light_cd_play_on_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let light_cd_pause_on_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        light_cd_pause_on_picture_offset,
    ) {
        Ok(light_cd_pause_on_picture) => light_cd_pause_on_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let element4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        element4_picture_offset,
    ) {
        Ok(element4_picture) => element4_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_open_stream_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_open_stream_picture_offset,
    ) {
        Ok(button_open_stream_picture) => button_open_stream_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_stop_stream_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_stop_stream_picture_offset,
    ) {
        Ok(button_stop_stream_picture) => button_stop_stream_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_play_stream_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_play_stream_picture_offset,
    ) {
        Ok(button_play_stream_picture) => button_play_stream_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let element5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        element5_picture_offset,
    ) {
        Ok(element5_picture) => element5_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let image_recording_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        recording_picture_offset,
    ) {
        Ok(image_recording_picture) => image_recording_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let image4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        image4_picture_offset,
    ) {
        Ok(image4_picture) => image4_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let picture17_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        picture17_picture_offset,
    ) {
        Ok(picture17_picture) => picture17_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer1_picture_offset,
    ) {
        Ok(cmd_audioplayer1_picture) => cmd_audioplayer1_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer2_picture_offset,
    ) {
        Ok(cmd_audioplayer2_picture) => cmd_audioplayer2_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer3_picture_offset,
    ) {
        Ok(cmd_audioplayer3_picture) => cmd_audioplayer3_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer4_picture_offset,
    ) {
        Ok(cmd_audioplayer4_picture) => cmd_audioplayer4_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer5_picture_offset,
    ) {
        Ok(cmd_audioplayer5_picture) => cmd_audioplayer5_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer6_picture_offset,
    ) {
        Ok(cmd_audioplayer6_picture) => cmd_audioplayer6_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer7_picture_offset,
    ) {
        Ok(cmd_audioplayer7_picture) => cmd_audioplayer7_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer8_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer8_picture_offset,
    ) {
        Ok(cmd_audioplayer8_picture) => cmd_audioplayer8_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let cmd_audioplayer9_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        cmd_audioplayer9_picture_offset,
    ) {
        Ok(cmd_audioplayer9_picture) => cmd_audioplayer9_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let light_dat_play_on_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        light_dat_play_on_picture_offset,
    ) {
        Ok(light_dat_play_on_picture) => light_dat_play_on_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let light_dat_pause_on_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        light_dat_pause_on_picture_offset,
    ) {
        Ok(light_dat_pause_on_picture) => light_dat_pause_on_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let element6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        element6_picture_offset,
    ) {
        Ok(element6_picture) => element6_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_midi_player1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_midi_player1_picture_offset,
    ) {
        Ok(button_midi_player1_picture) => button_midi_player1_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_midi_player2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_midi_player2_picture_offset,
    ) {
        Ok(button_midi_player2_picture) => button_midi_player2_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_midi_player3_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_midi_player3_picture_offset,
    ) {
        Ok(button_midi_player3_picture) => button_midi_player3_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_midi_player4_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_midi_player4_picture_offset,
    ) {
        Ok(button_midi_player4_picture) => button_midi_player4_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_midi_player5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_midi_player5_picture_offset,
    ) {
        Ok(button_midi_player5_picture) => button_midi_player5_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_midi_player6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_midi_player6_picture_offset,
    ) {
        Ok(button_midi_player6_picture) => button_midi_player6_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let button_midi_player7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        button_midi_player7_picture_offset,
    ) {
        Ok(button_midi_player7_picture) => button_midi_player7_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let light_midi_floppy_drive_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        light_midi_floppy_drive_picture_offset,
    ) {
        Ok(light_midi_floppy_drive_picture) => light_midi_floppy_drive_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let image6_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        image6_picture_offset,
    ) {
        Ok(image6_picture) => image6_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let floppy_in_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        floppy_in_picture_offset,
    ) {
        Ok(floppy_in_picture) => floppy_in_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let floppy_out_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        floppy_out_picture_offset,
    ) {
        Ok(floppy_out_picture) => floppy_out_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let light_midi_play_on_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        light_midi_play_on_picture_offset,
    ) {
        Ok(light_midi_play_on_picture) => light_midi_play_on_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let light_midi_pause_on_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        light_midi_pause_on_picture_offset,
    ) {
        Ok(light_midi_pause_on_picture) => light_midi_pause_on_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let element7_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        element7_picture_offset,
    ) {
        Ok(element7_picture) => element7_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let options_menu_button_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        options_menu_button_picture_offset,
    ) {
        Ok(options_menu_button_picture) => options_menu_button_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let image5_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
        image5_picture_offset,
    ) {
        Ok(image5_picture) => image5_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let elements_disabled_image_list_image1_picture =
        match vb6parse::parsers::resource_file_resolver(
            "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
            elements_disabled_image_list_image1_picture_offset,
        ) {
            Ok(elements_disabled_image_list_image1_picture) => {
                elements_disabled_image_list_image1_picture
            }
            Err(e) => panic!("Failed to resolve resource file: {e}"),
        };

    let elements_disabled_image_list_image2_picture =
        match vb6parse::parsers::resource_file_resolver(
            "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
            elements_disabled_image_list_image2_picture_offset,
        ) {
            Ok(elements_disabled_image_list_image2_picture) => {
                elements_disabled_image_list_image2_picture
            }
            Err(e) => panic!("Failed to resolve resource file: {e}"),
        };

    let elements_disabled_image_list_image3_picture =
        match vb6parse::parsers::resource_file_resolver(
            "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
            elements_disabled_image_list_image3_picture_offset,
        ) {
            Ok(elements_disabled_image_list_image3_picture) => {
                elements_disabled_image_list_image3_picture
            }
            Err(e) => panic!("Failed to resolve resource file: {e}"),
        };

    let elements_disabled_image_list_image4_picture =
        match vb6parse::parsers::resource_file_resolver(
            "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
            elements_disabled_image_list_image4_picture_offset,
        ) {
            Ok(elements_disabled_image_list_image4_picture) => {
                elements_disabled_image_list_image4_picture
            }
            Err(e) => panic!("Failed to resolve resource file: {e}"),
        };

    let elements_disabled_image_list_image5_picture =
        match vb6parse::parsers::resource_file_resolver(
            "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
            elements_disabled_image_list_image5_picture_offset,
        ) {
            Ok(elements_disabled_image_list_image5_picture) => {
                elements_disabled_image_list_image5_picture
            }
            Err(e) => panic!("Failed to resolve resource file: {e}"),
        };

    let elements_disabled_image_list_image6_picture =
        match vb6parse::parsers::resource_file_resolver(
            "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
            elements_disabled_image_list_image6_picture_offset,
        ) {
            Ok(elements_disabled_image_list_image6_picture) => {
                elements_disabled_image_list_image6_picture
            }
            Err(e) => panic!("Failed to resolve resource file: {e}"),
        };

    let elements_disabled_image_list_image7_picture =
        match vb6parse::parsers::resource_file_resolver(
            "./tests/data/audiostation/Audiostation/src/Forms/Form_Main.frx",
            elements_disabled_image_list_image7_picture_offset,
        ) {
            Ok(elements_disabled_image_list_image7_picture) => {
                elements_disabled_image_list_image7_picture
            }
            Err(e) => panic!("Failed to resolve resource file: {e}"),
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
    assert_eq!(image3_picture.len(), image3_picture_buffer_size);
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
    assert_eq!(
        light_cd_play_on_picture.len(),
        light_cd_play_on_picture_buffer_size
    );
    assert_eq!(
        light_cd_pause_on_picture.len(),
        light_cd_pause_on_picture_buffer_size
    );
    assert_eq!(element4_picture.len(), element4_picture_buffer_size);
    assert_eq!(
        button_open_stream_picture.len(),
        button_open_stream_picture_buffer_size
    );
    assert_eq!(
        button_stop_stream_picture.len(),
        button_stop_stream_picture_buffer_size
    );
    assert_eq!(
        button_play_stream_picture.len(),
        button_play_stream_picture_buffer_size
    );
    assert_eq!(element5_picture.len(), element5_picture_buffer_size);
    assert_eq!(image_recording_picture.len(), recording_picture_buffer_size);
    assert_eq!(image4_picture.len(), image4_picture_buffer_size);
    assert_eq!(picture17_picture.len(), picture17_picture_buffer_size);
    assert_eq!(
        cmd_audioplayer1_picture.len(),
        cmd_audioplayer1_picture_buffer_size
    );
    assert_eq!(
        cmd_audioplayer2_picture.len(),
        cmd_audioplayer2_picture_buffer_size
    );
    assert_eq!(
        cmd_audioplayer3_picture.len(),
        cmd_audioplayer3_picture_buffer_size
    );
    assert_eq!(
        cmd_audioplayer4_picture.len(),
        cmd_audioplayer4_picture_buffer_size
    );
    assert_eq!(
        cmd_audioplayer5_picture.len(),
        cmd_audioplayer5_picture_buffer_size
    );
    assert_eq!(
        cmd_audioplayer6_picture.len(),
        cmd_audioplayer6_picture_buffer_size
    );
    assert_eq!(
        cmd_audioplayer7_picture.len(),
        cmd_audioplayer7_picture_buffer_size
    );
    assert_eq!(
        cmd_audioplayer8_picture.len(),
        cmd_audioplayer8_picture_buffer_size
    );
    assert_eq!(
        cmd_audioplayer9_picture.len(),
        cmd_audioplayer9_picture_buffer_size
    );
    assert_eq!(
        light_dat_play_on_picture.len(),
        light_dat_play_on_picture_buffer_size
    );
    assert_eq!(
        light_dat_pause_on_picture.len(),
        light_dat_pause_on_picture_buffer_size
    );
    assert_eq!(element6_picture.len(), element6_picture_buffer_size);
    assert_eq!(
        button_midi_player1_picture.len(),
        button_midi_player1_picture_buffer_size
    );
    assert_eq!(
        button_midi_player2_picture.len(),
        button_midi_player2_picture_buffer_size
    );
    assert_eq!(
        button_midi_player3_picture.len(),
        button_midi_player3_picture_buffer_size
    );
    assert_eq!(
        button_midi_player4_picture.len(),
        button_midi_player4_picture_buffer_size
    );
    assert_eq!(
        button_midi_player5_picture.len(),
        button_midi_player5_picture_buffer_size
    );
    assert_eq!(
        button_midi_player6_picture.len(),
        button_midi_player6_picture_buffer_size
    );
    assert_eq!(
        button_midi_player7_picture.len(),
        button_midi_player7_picture_buffer_size
    );
    assert_eq!(
        light_midi_floppy_drive_picture.len(),
        light_midi_floppy_drive_picture_buffer_size
    );
    assert_eq!(image6_picture.len(), image6_picture_buffer_size);
    assert_eq!(floppy_in_picture.len(), floppy_in_picture_buffer_size);
    assert_eq!(floppy_out_picture.len(), floppy_out_picture_buffer_size);
    assert_eq!(
        light_midi_play_on_picture.len(),
        light_midi_play_on_picture_buffer_size
    );
    assert_eq!(
        light_midi_pause_on_picture.len(),
        light_midi_pause_on_picture_buffer_size
    );
    assert_eq!(element7_picture.len(), element7_picture_buffer_size);
    assert_eq!(
        options_menu_button_picture.len(),
        options_menu_button_picture_buffer_size
    );
    assert_eq!(image5_picture.len(), image5_picture_buffer_size);
    assert_eq!(
        elements_disabled_image_list_image1_picture.len(),
        elements_disabled_image_list_image1_picture_buffer_size
    );
    assert_eq!(
        elements_disabled_image_list_image2_picture.len(),
        elements_disabled_image_list_image2_picture_buffer_size
    );
    assert_eq!(
        elements_disabled_image_list_image3_picture.len(),
        elements_disabled_image_list_image3_picture_buffer_size
    );
    assert_eq!(
        elements_disabled_image_list_image4_picture.len(),
        elements_disabled_image_list_image4_picture_buffer_size
    );
    assert_eq!(
        elements_disabled_image_list_image5_picture.len(),
        elements_disabled_image_list_image5_picture_buffer_size
    );
    assert_eq!(
        elements_disabled_image_list_image6_picture.len(),
        elements_disabled_image_list_image6_picture_buffer_size
    );
    assert_eq!(
        elements_disabled_image_list_image7_picture.len(),
        elements_disabled_image_list_image7_picture_buffer_size
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
    assert_eq!(image3_picture, image3_picture_buffer);
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
    assert_eq!(light_cd_play_on_picture, light_cd_play_on_picture_buffer);
    assert_eq!(light_cd_pause_on_picture, light_cd_pause_on_picture_buffer);
    assert_eq!(element4_picture, element4_picture_buffer);
    assert_eq!(
        button_open_stream_picture,
        button_open_stream_picture_buffer
    );
    assert_eq!(
        button_stop_stream_picture,
        button_stop_stream_picture_buffer
    );
    assert_eq!(
        button_play_stream_picture,
        button_play_stream_picture_buffer
    );
    assert_eq!(element5_picture, element5_picture_buffer);
    assert_eq!(image_recording_picture, recording_picture_buffer);
    assert_eq!(image4_picture, image4_picture_buffer);
    assert_eq!(picture17_picture, picture17_picture_buffer);
    assert_eq!(cmd_audioplayer1_picture, cmd_audioplayer1_picture_buffer);
    assert_eq!(cmd_audioplayer2_picture, cmd_audioplayer2_picture_buffer);
    assert_eq!(cmd_audioplayer3_picture, cmd_audioplayer3_picture_buffer);
    assert_eq!(cmd_audioplayer4_picture, cmd_audioplayer4_picture_buffer);
    assert_eq!(cmd_audioplayer5_picture, cmd_audioplayer5_picture_buffer);
    assert_eq!(cmd_audioplayer6_picture, cmd_audioplayer6_picture_buffer);
    assert_eq!(cmd_audioplayer7_picture, cmd_audioplayer7_picture_buffer);
    assert_eq!(cmd_audioplayer8_picture, cmd_audioplayer8_picture_buffer);
    assert_eq!(cmd_audioplayer9_picture, cmd_audioplayer9_picture_buffer);
    assert_eq!(light_dat_play_on_picture, light_dat_play_on_picture_buffer);
    assert_eq!(
        light_dat_pause_on_picture,
        light_dat_pause_on_picture_buffer
    );
    assert_eq!(element6_picture, element6_picture_buffer);
    assert_eq!(
        button_midi_player1_picture,
        button_midi_player1_picture_buffer
    );
    assert_eq!(
        button_midi_player2_picture,
        button_midi_player2_picture_buffer
    );
    assert_eq!(
        button_midi_player3_picture,
        button_midi_player3_picture_buffer
    );
    assert_eq!(
        button_midi_player4_picture,
        button_midi_player4_picture_buffer
    );
    assert_eq!(
        button_midi_player5_picture,
        button_midi_player5_picture_buffer
    );
    assert_eq!(
        button_midi_player6_picture,
        button_midi_player6_picture_buffer
    );
    assert_eq!(
        button_midi_player7_picture,
        button_midi_player7_picture_buffer
    );
    assert_eq!(
        light_midi_floppy_drive_picture,
        light_midi_floppy_drive_picture_buffer
    );
    assert_eq!(image6_picture, image6_picture_buffer);
    assert_eq!(floppy_in_picture, floppy_in_picture_buffer);
    assert_eq!(floppy_out_picture, floppy_out_picture_buffer);
    assert_eq!(
        light_midi_play_on_picture,
        light_midi_play_on_picture_buffer
    );
    assert_eq!(
        light_midi_pause_on_picture,
        light_midi_pause_on_picture_buffer
    );
    assert_eq!(element7_picture, element7_picture_buffer);
    assert_eq!(
        options_menu_button_picture,
        options_menu_button_picture_buffer
    );
    assert_eq!(image5_picture, image5_picture_buffer);
    assert_eq!(
        elements_disabled_image_list_image1_picture,
        elements_disabled_image_list_image1_picture_buffer
    );
    assert_eq!(
        elements_disabled_image_list_image2_picture,
        elements_disabled_image_list_image2_picture_buffer
    );
    assert_eq!(
        elements_disabled_image_list_image3_picture,
        elements_disabled_image_list_image3_picture_buffer
    );
    assert_eq!(
        elements_disabled_image_list_image4_picture,
        elements_disabled_image_list_image4_picture_buffer
    );
    assert_eq!(
        elements_disabled_image_list_image5_picture,
        elements_disabled_image_list_image5_picture_buffer
    );
    assert_eq!(
        elements_disabled_image_list_image6_picture,
        elements_disabled_image_list_image6_picture_buffer
    );
    assert_eq!(
        elements_disabled_image_list_image7_picture,
        elements_disabled_image_list_image7_picture_buffer
    );
}

#[test]
fn audiostation_normalize_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Normalize.frx");

    let normalize_icon_offset = 0x00;
    let normalize_icon_header_size = 12;
    let normalize_icon_buffer_size = 0;
    let normalize_icon_buffer_start = normalize_icon_offset + normalize_icon_header_size;
    let normalize_icon_buffer_end = normalize_icon_buffer_start + normalize_icon_buffer_size;
    let normalize_icon_buffer =
        resource_file_bytes[normalize_icon_buffer_start..normalize_icon_buffer_end].to_vec();

    let normalize_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Normalize.frx",
        normalize_icon_offset,
    ) {
        Ok(normalize_icon) => normalize_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    assert_eq!(normalize_icon.len(), normalize_icon_buffer_size);
    assert_eq!(normalize_icon, normalize_icon_buffer);
}

#[test]
fn audiostation_open_dialog_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_OpenDialog.frx");

    let open_stream_form_icon_offset = 0x0000;
    let open_stream_form_icon_header_size = 12;
    let open_stream_form_icon_buffer_size = 0;
    let open_stream_form_icon_buffer_start =
        open_stream_form_icon_offset + open_stream_form_icon_header_size;
    let open_stream_form_icon_buffer_end =
        open_stream_form_icon_buffer_start + open_stream_form_icon_buffer_size;
    let open_stream_form_icon_buffer = resource_file_bytes
        [open_stream_form_icon_buffer_start..open_stream_form_icon_buffer_end]
        .to_vec();

    let tab_strip_tab_picture0_offset = 0x000C;
    let tab_strip_tab_picture0_header_size = 4;
    let tab_strip_tab_picture0_buffer_size = 24;
    let tab_strip_tab_picture0_buffer_start =
        tab_strip_tab_picture0_offset + tab_strip_tab_picture0_header_size;
    let tab_strip_tab_picture0_buffer_end =
        tab_strip_tab_picture0_buffer_start + tab_strip_tab_picture0_buffer_size;
    let tab_strip_tab_picture0_buffer = resource_file_bytes
        [tab_strip_tab_picture0_buffer_start..tab_strip_tab_picture0_buffer_end]
        .to_vec();

    let list_view_list_image1_picture_offset = 0x0028;
    let list_view_list_image1_picture_header_size = 4;
    let list_view_list_image1_picture_buffer_size = 343;
    let list_view_list_image1_picture_buffer_start =
        list_view_list_image1_picture_offset + list_view_list_image1_picture_header_size;
    let list_view_list_image1_picture_buffer_end =
        list_view_list_image1_picture_buffer_start + list_view_list_image1_picture_buffer_size;
    let list_view_list_image1_picture_buffer = resource_file_bytes
        [list_view_list_image1_picture_buffer_start..list_view_list_image1_picture_buffer_end]
        .to_vec();

    let list_view_list_image2_picture_offset = 0x0183;
    let list_view_list_image2_picture_header_size = 4;
    let list_view_list_image2_picture_buffer_size = 1430;
    let list_view_list_image2_picture_buffer_start =
        list_view_list_image2_picture_offset + list_view_list_image2_picture_header_size;
    let list_view_list_image2_picture_buffer_end =
        list_view_list_image2_picture_buffer_start + list_view_list_image2_picture_buffer_size;
    let list_view_list_image2_picture_buffer = resource_file_bytes
        [list_view_list_image2_picture_buffer_start..list_view_list_image2_picture_buffer_end]
        .to_vec();

    let open_stream_form_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_OpenDialog.frx",
        open_stream_form_icon_offset,
    ) {
        Ok(open_stream_form_icon) => open_stream_form_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let tab_strip_tab_picture0 = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_OpenDialog.frx",
        tab_strip_tab_picture0_offset,
    ) {
        Ok(tab_strip_tab_picture0) => tab_strip_tab_picture0,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let list_view_list_image1_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_OpenDialog.frx",
        list_view_list_image1_picture_offset,
    ) {
        Ok(list_view_list_image1_picture) => list_view_list_image1_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let list_view_list_image2_picture = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_OpenDialog.frx",
        list_view_list_image2_picture_offset,
    ) {
        Ok(list_view_list_image2_picture) => list_view_list_image2_picture,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    assert_eq!(
        open_stream_form_icon.len(),
        open_stream_form_icon_buffer_size
    );
    assert_eq!(
        tab_strip_tab_picture0.len(),
        tab_strip_tab_picture0_buffer_size
    );
    assert_eq!(
        list_view_list_image1_picture.len(),
        list_view_list_image1_picture_buffer_size
    );
    assert_eq!(
        list_view_list_image2_picture.len(),
        list_view_list_image2_picture_buffer_size
    );

    assert_eq!(open_stream_form_icon, open_stream_form_icon_buffer);
    assert_eq!(tab_strip_tab_picture0, tab_strip_tab_picture0_buffer);
    assert_eq!(
        list_view_list_image1_picture,
        list_view_list_image1_picture_buffer
    );
    assert_eq!(
        list_view_list_image2_picture,
        list_view_list_image2_picture_buffer
    );
}

#[test]
fn audiostation_playlist_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Playlist.frx");

    let playlist_form_icon_offset = 0x0000;
    let playlist_form_icon_header_size = 12;
    let playlist_form_icon_buffer_size = 0;
    let playlist_form_icon_buffer_start =
        playlist_form_icon_offset + playlist_form_icon_header_size;
    let playlist_form_icon_buffer_end =
        playlist_form_icon_buffer_start + playlist_form_icon_buffer_size;
    let playlist_form_icon_buffer = resource_file_bytes
        [playlist_form_icon_buffer_start..playlist_form_icon_buffer_end]
        .to_vec();

    let playlist_form_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Playlist.frx",
        playlist_form_icon_offset,
    ) {
        Ok(playlist_form_icon) => playlist_form_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };
    assert_eq!(playlist_form_icon.len(), playlist_form_icon_buffer_size);
    assert_eq!(playlist_form_icon, playlist_form_icon_buffer);
}

#[test]
fn audiostation_plugins_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Plugins.frx");

    let plugins_form_icon_offset = 0x0000;
    let plugins_form_icon_header_size = 12;
    let plugins_form_icon_buffer_size = 0;
    let plugins_form_icon_buffer_start = plugins_form_icon_offset + plugins_form_icon_header_size;
    let plugins_form_icon_buffer_end =
        plugins_form_icon_buffer_start + plugins_form_icon_buffer_size;
    let plugins_form_icon_buffer =
        resource_file_bytes[plugins_form_icon_buffer_start..plugins_form_icon_buffer_end].to_vec();

    let plugins_form_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Plugins.frx",
        plugins_form_icon_offset,
    ) {
        Ok(plugins_form_icon) => plugins_form_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };
    assert_eq!(plugins_form_icon.len(), plugins_form_icon_buffer_size);
    assert_eq!(plugins_form_icon, plugins_form_icon_buffer);
}

#[test]
fn audiostation_settings_record_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Settings_Record.frx");

    let settings_record_form_icon_offset = 0x0000;
    let settings_record_form_icon_header_size = 12;
    let settings_record_form_icon_buffer_size = 0;
    let settings_record_form_icon_buffer_start =
        settings_record_form_icon_offset + settings_record_form_icon_header_size;
    let settings_record_form_icon_buffer_end =
        settings_record_form_icon_buffer_start + settings_record_form_icon_buffer_size;
    let settings_record_form_icon_buffer = resource_file_bytes
        [settings_record_form_icon_buffer_start..settings_record_form_icon_buffer_end]
        .to_vec();

    // lists are a bit special since we need the list item count and that's in the
    // resource header which means unlike the other resources we need to read the header
    // and include that in the buffer size, that is why the header_size is listed as 0.
    let combo_language_item_data_offset = 0x000C;
    let combo_language_item_data_header_size = 0;
    let combo_language_item_data_buffer_size = 13;
    let combo_language_item_data_buffer_start =
        combo_language_item_data_offset + combo_language_item_data_header_size;
    let combo_language_item_data_buffer_end =
        combo_language_item_data_buffer_start + combo_language_item_data_buffer_size;
    let combo_language_item_data_buffer = resource_file_bytes
        [combo_language_item_data_buffer_start..combo_language_item_data_buffer_end]
        .to_vec();

    let combo_language_data_items = list_resolver(&combo_language_item_data_buffer);

    let combo_language_list_offset = 0x0019;
    let combo_language_list_header_size = 0;
    let combo_language_list_buffer_size = 28;
    let combo_language_list_buffer_start =
        combo_language_list_offset + combo_language_list_header_size;
    let combo_language_list_buffer_end =
        combo_language_list_buffer_start + combo_language_list_buffer_size;
    let combo_language_list_buffer = resource_file_bytes
        [combo_language_list_buffer_start..combo_language_list_buffer_end]
        .to_vec();

    let combo_language_list_items = list_resolver(&combo_language_list_buffer);

    let combo_midi_device_item_data_offset = 0x0035;
    let combo_midi_device_item_data_header_size = 0;
    let combo_midi_device_item_data_buffer_size = 13;
    let combo_midi_device_item_data_buffer_start =
        combo_midi_device_item_data_offset + combo_midi_device_item_data_header_size;
    let combo_midi_device_item_data_buffer_end =
        combo_midi_device_item_data_buffer_start + combo_midi_device_item_data_buffer_size;
    let combo_midi_device_item_data_buffer = resource_file_bytes
        [combo_midi_device_item_data_buffer_start..combo_midi_device_item_data_buffer_end]
        .to_vec();

    let combo_midi_device_data_items = list_resolver(&combo_midi_device_item_data_buffer);

    let combo_midi_device_list_offset = 0x0042;
    let combo_midi_device_list_header_size = 0;
    let combo_midi_device_list_buffer_size = 28;
    let combo_midi_device_list_buffer_start =
        combo_midi_device_list_offset + combo_midi_device_list_header_size;
    let combo_midi_device_list_buffer_end =
        combo_midi_device_list_buffer_start + combo_midi_device_list_buffer_size;
    let combo_midi_device_list_buffer = resource_file_bytes
        [combo_midi_device_list_buffer_start..combo_midi_device_list_buffer_end]
        .to_vec();

    let combo_midi_device_list_items = list_resolver(&combo_midi_device_list_buffer);

    let settings_record_form_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Settings_Record.frx",
        settings_record_form_icon_offset,
    ) {
        Ok(settings_record_form_icon) => settings_record_form_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let combo_language_item_data = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Settings_Record.frx",
        combo_language_item_data_offset,
    ) {
        Ok(combo_language_item_data) => combo_language_item_data,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let combo_language_list = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Settings_Record.frx",
        combo_language_list_offset,
    ) {
        Ok(combo_language_list) => combo_language_list,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let combo_midi_device_item_data = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Settings_Record.frx",
        combo_midi_device_item_data_offset,
    ) {
        Ok(combo_midi_device_item_data) => combo_midi_device_item_data,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let combo_midi_device_list = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Settings_Record.frx",
        combo_midi_device_list_offset,
    ) {
        Ok(combo_midi_device_list) => combo_midi_device_list,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    assert_eq!(
        settings_record_form_icon.len(),
        settings_record_form_icon_buffer_size
    );
    assert_eq!(
        combo_language_item_data.len(),
        combo_language_item_data_buffer_size
    );
    assert_eq!(combo_language_list.len(), combo_language_list_buffer_size);
    assert_eq!(
        combo_midi_device_item_data.len(),
        combo_midi_device_item_data_buffer_size
    );
    assert_eq!(
        combo_midi_device_list.len(),
        combo_midi_device_list_buffer_size
    );

    assert_eq!(settings_record_form_icon, settings_record_form_icon_buffer);
    assert_eq!(combo_language_item_data, combo_language_item_data_buffer);
    assert_eq!(combo_language_list, combo_language_list_buffer);
    assert_eq!(
        combo_midi_device_item_data,
        combo_midi_device_item_data_buffer
    );
    assert_eq!(combo_midi_device_list, combo_midi_device_list_buffer);

    assert_eq!(combo_language_data_items.len(), 3);
    assert_eq!(combo_language_data_items[0], "0");
    assert_eq!(combo_language_data_items[1], "0");
    assert_eq!(combo_language_data_items[2], "0");

    assert_eq!(combo_language_list_items.len(), 3);
    assert_eq!(combo_language_list_items[0], "English");
    assert_eq!(combo_language_list_items[1], "Dutch");
    assert_eq!(combo_language_list_items[2], "German");

    assert_eq!(combo_midi_device_data_items.len(), 3);
    assert_eq!(combo_midi_device_data_items[0], "0");
    assert_eq!(combo_midi_device_data_items[1], "0");
    assert_eq!(combo_midi_device_data_items[2], "0");

    assert_eq!(combo_midi_device_list_items.len(), 3);
    assert_eq!(combo_midi_device_list_items[0], "English");
    assert_eq!(combo_midi_device_list_items[1], "Dutch");
    assert_eq!(combo_midi_device_list_items[2], "German");
}

#[test]
fn audiostation_settings_recorder_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Settings_Recorder.frx");

    let settings_recorder_form_icon_offset = 0x0000;
    let settings_recorder_form_icon_header_size = 12;
    let settings_recorder_form_icon_buffer_size = 0;
    let settings_recorder_form_icon_buffer_start =
        settings_recorder_form_icon_offset + settings_recorder_form_icon_header_size;
    let settings_recorder_form_icon_buffer_end =
        settings_recorder_form_icon_buffer_start + settings_recorder_form_icon_buffer_size;
    let settings_recorder_form_icon_buffer = resource_file_bytes
        [settings_recorder_form_icon_buffer_start..settings_recorder_form_icon_buffer_end]
        .to_vec();

    let settings_recorder_form_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Settings_Recorder.frx",
        settings_recorder_form_icon_offset,
    ) {
        Ok(settings_recorder_form_icon) => settings_recorder_form_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    assert_eq!(
        settings_recorder_form_icon.len(),
        settings_recorder_form_icon_buffer_size
    );
    assert_eq!(
        settings_recorder_form_icon,
        settings_recorder_form_icon_buffer
    );
}

#[test]
fn audiostation_streams_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Streams.frx");

    let streams_form_icon_offset = 0x0000;
    let streams_form_icon_header_size = 12;
    let streams_form_icon_buffer_size = 0;
    let streams_form_icon_buffer_start = streams_form_icon_offset + streams_form_icon_header_size;
    let streams_form_icon_buffer_end =
        streams_form_icon_buffer_start + streams_form_icon_buffer_size;
    let streams_form_icon_buffer =
        resource_file_bytes[streams_form_icon_buffer_start..streams_form_icon_buffer_end].to_vec();

    let streams_form_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Streams.frx",
        streams_form_icon_offset,
    ) {
        Ok(streams_form_icon) => streams_form_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    assert_eq!(streams_form_icon.len(), streams_form_icon_buffer_size);
    assert_eq!(streams_form_icon, streams_form_icon_buffer);
}

#[test]
fn audiostation_track_properties_frx_load() {
    let resource_file_bytes =
        include_bytes!("./data/audiostation/Audiostation/src/Forms/Form_Track_Properties.frx");

    let track_properties_form_icon_offset = 0x0000;
    let track_properties_form_icon_header_size = 12;
    let track_properties_form_icon_buffer_size = 0;
    let track_properties_form_icon_buffer_start =
        track_properties_form_icon_offset + track_properties_form_icon_header_size;
    let track_properties_form_icon_buffer_end =
        track_properties_form_icon_buffer_start + track_properties_form_icon_buffer_size;
    let track_properties_form_icon_buffer = resource_file_bytes
        [track_properties_form_icon_buffer_start..track_properties_form_icon_buffer_end]
        .to_vec();

    let track_properties_text_properties_text_offset = 0x000C;
    let track_properties_text_properties_text_header_size = 1;
    let track_properties_text_properties_text_buffer_size = 20;
    let track_properties_text_properties_text_buffer_start =
        track_properties_text_properties_text_offset
            + track_properties_text_properties_text_header_size;
    let track_properties_text_properties_text_buffer_end =
        track_properties_text_properties_text_buffer_start
            + track_properties_text_properties_text_buffer_size;
    let track_properties_text_properties_text_buffer = resource_file_bytes
        [track_properties_text_properties_text_buffer_start
            ..track_properties_text_properties_text_buffer_end]
        .to_vec();

    let track_properties_form_icon = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Track_Properties.frx",
        track_properties_form_icon_offset,
    ) {
        Ok(track_properties_form_icon) => track_properties_form_icon,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    let track_properties_text_properties_text = match vb6parse::parsers::resource_file_resolver(
        "./tests/data/audiostation/Audiostation/src/Forms/Form_Track_Properties.frx",
        track_properties_text_properties_text_offset,
    ) {
        Ok(track_properties_text_properties_text) => track_properties_text_properties_text,
        Err(e) => panic!("Failed to resolve resource file: {e}"),
    };

    assert_eq!(
        track_properties_form_icon.len(),
        track_properties_form_icon_buffer_size
    );
    assert_eq!(
        track_properties_text_properties_text.len(),
        track_properties_text_properties_text_buffer_size
    );

    assert_eq!(
        track_properties_form_icon,
        track_properties_form_icon_buffer
    );
    assert_eq!(
        track_properties_text_properties_text,
        track_properties_text_properties_text_buffer
    );

    assert_eq!(
        track_properties_text_properties_text,
        b"Textbox_Properties\r\n"
    );
}
