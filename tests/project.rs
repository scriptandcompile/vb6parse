use vb6parse::parsers::VB6Project;

#[test]
fn artificial_life_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Artificial-life/Artificial Life.vbp");

    let project = match VB6Project::parse("Artificial Life.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Artificial Life.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn blacklight_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Blacklight-effect/Blacklight.vbp");

    let project = match VB6Project::parse("Blacklight.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Blacklight.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn brightness_effect_part_1_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp");

    let project = match VB6Project::parse("Brightness.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Brightness.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn brightness_effect_part_2_project_load() {
    let project_file_bytes = include_bytes!(
        "./data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.vbp"
    );

    let project = match VB6Project::parse("Brightness2.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Brightness2.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn brightness_effect_part_3_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.vbp");

    let project = match VB6Project::parse("Brightness3.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Brightness3.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn color_shift_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Color-shift-effect/ShiftColor.vbp");

    let project = match VB6Project::parse("ShiftColor.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'ShiftColor.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn colorize_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Colorize-effect/Colorize.vbp");

    let project = match VB6Project::parse("Colorize.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Colorize.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn contrast_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Contrast-effect/Contrast.vbp");

    let project = match VB6Project::parse("Contrast.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Contrast.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn curves_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Curves-effect/Curves.vbp");

    let project = match VB6Project::parse("Curves.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Curves.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn custom_image_filters_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Custom-image-filters/CustomFilters.vbp");

    let project = match VB6Project::parse("CustomFilters.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'CustomFilters.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn diffuse_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Diffuse-effect/Diffuse.vbp");

    let project = match VB6Project::parse("Diffuse.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Diffuse.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn edge_detection_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Edge-detection/EdgeDetection.vbp");

    let project = match VB6Project::parse("EdgeDetection.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'EdgeDetection.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn emboss_engrave_effect_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Emboss-engrave-effect/EmbossEngrave.vbp");

    let project = match VB6Project::parse("EmbossEngrave.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'EmbossEngrave.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn fill_image_region_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Fill-image-region/Fill_Region.vbp");

    let project = match VB6Project::parse("Fill_Region.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Fill_Region.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn fire_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Fire-effect/FlameTest.vbp");

    let project = match VB6Project::parse("FlameTest.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'FlameTest.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn game_physics_basic_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Game-physics-basic/Physics.vbp");

    let project = match VB6Project::parse("Physics.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Physics.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn gradient_2d_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Gradient-2D/Gradient.vbp");

    let project = match VB6Project::parse("Gradient.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Gradient.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn grayscale_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Grayscale-effect/Grayscale.vbp");

    let project = match VB6Project::parse("Grayscale.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Grayscale.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn hidden_markov_model_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Hidden-Markov-model/HMM.vbp");

    let project = match VB6Project::parse("HMM.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'HMM.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn histograms_advanced_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Histograms-advanced/Advanced Histograms.vbp");

    let project = match VB6Project::parse("Advanced Histograms.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!(
                "Failed to parse class file 'Advanced Histograms.vbp': {}",
                e
            );
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn histogram_basic_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Histograms-basic/Basic Histograms.vbp");

    let project = match VB6Project::parse("Basic Histograms.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Basic Histograms.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn levels_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Levels-effect/Image Levels.vbp");

    let project = match VB6Project::parse("Image Levels.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Image Levels.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn mandelbrot_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Mandelbrot/Mandelbrot.vbp");

    let project = match VB6Project::parse("Mandelbrot.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Mandelbrot.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn map_editor_2d_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Map-editor-2D/Map Editor.vbp");

    let project = match VB6Project::parse("Map Editor.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Map Editor.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn nature_effects_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Nature-effects/NatureFilters.vbp");

    let project = match VB6Project::parse("NatureFilters.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'NatureFilters.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn randomize_effects_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Randomize-effects/RandomizationFX.vbp");

    let project = match VB6Project::parse("RandomizationFX.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'RandomizationFX.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn scanner_twain_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Scanner-TWAIN/VB_Scanner_Support.vbp");

    let project = match VB6Project::parse("VB_Scanner_Support.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'VB_Scanner_Support.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn screen_capture_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Screen-capture/ScreenCapture.vbp");

    let project = match VB6Project::parse("ScreenCapture.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'ScreenCapture.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn sepia_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Sepia-effect/Sepia.vbp");

    let project = match VB6Project::parse("Sepia.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Sepia.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn threshold_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Threshold-effect/Threshold.vbp");

    let project = match VB6Project::parse("Threshold.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Threshold.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}

#[test]
fn transparency_2d_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Transparency-2D/Transparency.vbp");

    let project = match VB6Project::parse("Transparency.vbp", project_file_bytes) {
        Ok(project) => project,
        Err(e) => {
            panic!("Failed to parse class file 'Transparency.vbp': {}", e);
        }
    };

    insta::assert_yaml_snapshot!(project);
}
