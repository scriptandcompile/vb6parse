use vb6parse::*;

#[test]
fn artificial_life_project_load() {
    let file_path = "./tests/data/vb6-code/Artificial-life/Artificial Life.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let result = SourceFile::decode_with_replacement(file_path, &project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn blacklight_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Blacklight-effect/Blacklight.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let result = SourceFile::decode_with_replacement(file_path, &project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn brightness_effect_part_1_project_load() {
    let file_path = "./tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let result = SourceFile::decode_with_replacement(file_path, &project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn brightness_effect_part_2_project_load() {
    let file_path =
        "./tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let result = SourceFile::decode_with_replacement(file_path, &project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn brightness_effect_part_3_project_load() {
    let file_path = "./tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let result = SourceFile::decode_with_replacement(file_path, &project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn color_shift_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Color-shift-effect/ShiftColor.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let result = SourceFile::decode_with_replacement(file_path, &project_file_bytes);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file '{file_path}': {e:?}"),
    };

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn colorize_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Colorize-effect/Colorize.vbp";

    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn contrast_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Contrast-effect/Contrast.vbp";

    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn curves_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Curves-effect/Curves.vbp";

    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn custom_image_filters_project_load() {
    let file_path = "./tests/data/vb6-code/Custom-image-filters/CustomFilters.vbp";

    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn diffuse_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Diffuse-effect/Diffuse.vbp";

    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn edge_detection_project_load() {
    let file_path = "./tests/data/vb6-code/Edge-detection/EdgeDetection.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn emboss_engrave_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Emboss-engrave-effect/EmbossEngrave.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn fill_image_region_project_load() {
    let file_path = "./tests/data/vb6-code/Fill-image-region/Fill_Region.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn fire_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Fire-effect/FlameTest.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn game_physics_basic_project_load() {
    let file_path = "./tests/data/vb6-code/Game-physics-basic/Physics.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn gradient_2d_project_load() {
    let file_path = "./tests/data/vb6-code/Gradient-2D/Gradient.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn grayscale_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Grayscale-effect/Grayscale.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn hidden_markov_model_project_load() {
    let file_path = "./tests/data/vb6-code/Hidden-Markov-model/HMM.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn histograms_advanced_project_load() {
    let file_path = "./tests/data/vb6-code/Histograms-advanced/Advanced Histograms.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn histogram_basic_project_load() {
    let file_path = "./tests/data/vb6-code/Histograms-basic/Basic Histograms.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn levels_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Levels-effect/Image Levels.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn mandelbrot_project_load() {
    let file_path = "./tests/data/vb6-code/Mandelbrot/Mandelbrot.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn map_editor_2d_project_load() {
    let file_path = "./tests/data/vb6-code/Map-editor-2D/Map Editor.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn nature_effects_project_load() {
    let file_path = "./tests/data/vb6-code/Nature-effects/NatureFilters.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn randomize_effects_project_load() {
    let file_path = "./tests/data/vb6-code/Randomize-effects/RandomizationFX.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn scanner_twain_project_load() {
    let file_path = "./tests/data/vb6-code/Scanner-TWAIN/VB_Scanner_Support.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn screen_capture_project_load() {
    let file_path = "./tests/data/vb6-code/Screen-capture/ScreenCapture.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn sepia_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Sepia-effect/Sepia.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn threshold_effect_project_load() {
    let file_path = "./tests/data/vb6-code/Threshold-effect/Threshold.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}

#[test]
fn transparency_2d_project_load() {
    let file_path = "./tests/data/vb6-code/Transparency-2D/Transparency.vbp";
    let project_file_bytes = std::fs::read(file_path).expect("Failed to read project file");

    let source_file = SourceFile::decode_with_replacement(file_path, &project_file_bytes)
        .expect("Failed to decode project file");

    let (project_file_opt, failures) = ProjectFile::parse(&source_file).unpack();

    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = project_file_opt.expect("Project should be present.");

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/project");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(project);
}
