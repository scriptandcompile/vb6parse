use vb6parse::parsers::form::resource_file_resolver;
use vb6parse::parsers::VB6FormFile;

#[test]
fn artificial_life_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Artificial-life/frmMain.frm");

    let form_file = match VB6FormFile::parse(
        "frmMain.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Blacklight.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn blacklight_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Blacklight-effect/Blacklight.frm");

    let form_file = match VB6FormFile::parse(
        "Blacklight.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Blacklight.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn brightness_effect_part_1_form_load() {
    let form_file_bytes =
        include_bytes!("./data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.frm");

    let form_file = match VB6FormFile::parse(
        "Brightness.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Brightness.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn brightness_effect_part_2_form_load() {
    let form_file_bytes = include_bytes!(
        "./data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.frm"
    );

    let form_file = match VB6FormFile::parse(
        "Brightness2.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Brightness2.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn brightness_effect_part_3_form_load() {
    let form_file_bytes =
        include_bytes!("./data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.frm");

    let form_file = match VB6FormFile::parse(
        "Brightness3.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Brightness3.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn brightness_effect_part_4_form_load() {
    let form_file_bytes = include_bytes!(
        "./data/vb6-code/Brightness-effect/Part 4 - Even faster DIBs/Brightness.frm"
    );

    let form_file = match VB6FormFile::parse(
        "Brightness.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Brightness.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn color_shift_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Color-shift-effect/ShiftColors.frm");

    let form_file = match VB6FormFile::parse(
        "ShiftColors.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'ShiftColors.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn colorize_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Colorize-effect/Colorize.frm");

    let form_file = match VB6FormFile::parse(
        "Colorize.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Colorize.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn contrast_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Contrast-effect/Contrast.frm");

    let form_file = match VB6FormFile::parse(
        "Contrast.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Contrast.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn curves_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Curves-effect/Curves.frm");

    let form_file = match VB6FormFile::parse(
        "Curves.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Curves.frm' form file");
        }
    };
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn custom_image_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Custom-image-filters/CustomFilters.frm");

    let form_file = match VB6FormFile::parse(
        "CustomFilters.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'CustomFilters.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn diffuse_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Diffuse-effect/Diffuse.frm");

    let form_file = match VB6FormFile::parse(
        "Diffuse.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Diffuse.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn edge_detection_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Edge-detection/EdgeDetection.frm");

    let form_file = match VB6FormFile::parse(
        "EdgeDetection.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'EdgeDetection.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn emboss_engrave_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Emboss-engrave-effect/EmbossEngrave.frm");

    let form_file = match VB6FormFile::parse(
        "EmbossEngrave.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'EmbossEngrave.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn fill_image_region_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Fill-image-region/frmFill.frm");

    let form_file = match VB6FormFile::parse(
        "frmFill.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'frmFill.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn fire_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Fire-effect/frmFire.frm");

    let form_file = match VB6FormFile::parse(
        "frmFire.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'frmFire.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn game_physics_basic_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Game-physics-basic/FormPhysics.frm");

    let form_file = match VB6FormFile::parse(
        "frmPhysics.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'frmPhysics.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn gradient_2d_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Gradient-2D/Gradient.frm");

    let form_file = match VB6FormFile::parse(
        "Gradient.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Gradient.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn grayscale_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Grayscale-effect/Grayscale.frm");

    let form_file = match VB6FormFile::parse(
        "Grayscale.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Grayscale.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn hidden_markov_model_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Hidden-Markov-model/frmHMM.frm");

    let form_file = match VB6FormFile::parse(
        "frmHMM.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'frmHMM.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn histograms_advanced_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Histograms-advanced/Histogram.frm");

    let form_file = match VB6FormFile::parse(
        "Histogram.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Histogram.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn histograms_basic_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Histograms-basic/Histogram.frm");

    let form_file = match VB6FormFile::parse(
        "Histogram.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Histogram.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn levels_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Levels-effect/Main.frm");

    let form_file = match VB6FormFile::parse(
        "Main.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Main.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn mandelbrot_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Mandelbrot/Mandelbrot.frm");

    let form_file = match VB6FormFile::parse(
        "Mandelbrot.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Main.frm' form file:");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn map_editor_2d_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Map-editor-2D/Main Editor.frm");

    let form_file = match VB6FormFile::parse(
        "Main Editor.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Main Editor.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn nature_effects_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Nature-effects/NatureFilters.frm");

    let form_file = match VB6FormFile::parse(
        "NatureFilters.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'NatureFilters.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn randomize_effects_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Randomize-effects/RandomizationFX.frm");

    let form_file = match VB6FormFile::parse(
        "RandomizationFX.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'RandomizationFX.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn scanner_twain_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Scanner-TWAIN/frmScanner.frm");

    let form_file = match VB6FormFile::parse(
        "frmScanner.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'frmScanner.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn screen_capture_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Screen-capture/FormScreenCapture.frm");

    let form_file = match VB6FormFile::parse(
        "FormScreenCapture.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'FormScreenCapture.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn sepia_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Sepia-effect/Sepia.frm");

    let form_file = match VB6FormFile::parse(
        "Sepia.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Sepia.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn threshold_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Threshold-effect/Threshold.frm");

    let form_file = match VB6FormFile::parse(
        "Threshold.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Threshold.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn transparency_2d_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Transparency-2D/frmTransparency.frm");

    let form_file = match VB6FormFile::parse(
        "frmTransparency.frm".to_owned(),
        form_file_bytes,
        resource_file_resolver,
    ) {
        Ok(form_file) => form_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'frmTransparency.frm' form file");
        }
    };

    insta::assert_yaml_snapshot!(form_file);
}
