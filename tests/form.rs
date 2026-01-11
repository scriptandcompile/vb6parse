use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

#[test]
fn artificial_life_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Artificial-life/frmMain.frm");

    let source_file = SourceFile::decode_with_replacement("frmMain.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'frmMain.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }

        panic!("Failed to parse 'frmMain.frm' form file");
    }

    let form_file = form_result.unwrap();

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn blacklight_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Blacklight-effect/Blacklight.frm");

    let source_file = SourceFile::decode_with_replacement("Blacklight.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Blacklight.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Blacklight.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn brightness_effect_part_1_form_load() {
    let form_file_bytes =
        include_bytes!("./data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.frm");

    let source_file = SourceFile::decode_with_replacement("Brightness.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Brightness.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Brightness.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn brightness_effect_part_2_form_load() {
    let form_file_bytes = include_bytes!(
        "./data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.frm"
    );

    let source_file = SourceFile::decode_with_replacement("Brightness2.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Brightness2.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Brightness2.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn brightness_effect_part_3_form_load() {
    let form_file_bytes =
        include_bytes!("./data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.frm");

    let source_file = SourceFile::decode_with_replacement("Brightness3.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Brightness3.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Brightness3.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn brightness_effect_part_4_form_load() {
    let form_file_bytes = include_bytes!(
        "./data/vb6-code/Brightness-effect/Part 4 - Even faster DIBs/Brightness.frm"
    );

    let source_file = SourceFile::decode_with_replacement("Brightness.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Brightness.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Brightness.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn color_shift_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Color-shift-effect/ShiftColors.frm");

    let source_file = SourceFile::decode_with_replacement("ShiftColors.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'ShiftColors.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'ShiftColors.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn colorize_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Colorize-effect/Colorize.frm");

    let source_file = SourceFile::decode_with_replacement("Colorize.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Colorize.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Colorize.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn contrast_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Contrast-effect/Contrast.frm");

    let source_file = SourceFile::decode_with_replacement("Contrast.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Contrast.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Contrast.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn curves_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Curves-effect/Curves.frm");

    let source_file = SourceFile::decode_with_replacement("Curves.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Curves.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Curves.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn custom_image_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Custom-image-filters/CustomFilters.frm");

    let source_file = SourceFile::decode_with_replacement("CustomFilters.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'CustomFilters.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'CustomFilters.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn diffuse_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Diffuse-effect/Diffuse.frm");

    let source_file = SourceFile::decode_with_replacement("Diffuse.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Diffuse.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Diffuse.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn edge_detection_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Edge-detection/EdgeDetection.frm");

    let source_file = SourceFile::decode_with_replacement("EdgeDetection.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'EdgeDetection.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'EdgeDetection.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn emboss_engrave_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Emboss-engrave-effect/EmbossEngrave.frm");

    let source_file = SourceFile::decode_with_replacement("EmbossEngrave.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'EmbossEngrave.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'EmbossEngrave.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn fill_image_region_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Fill-image-region/frmFill.frm");

    let source_file = SourceFile::decode_with_replacement("frmFill.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'frmFill.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'frmFill.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn fire_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Fire-effect/frmFire.frm");

    let source_file = SourceFile::decode_with_replacement("frmFire.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'frmFire.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'frmFire.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn game_physics_basic_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Game-physics-basic/FormPhysics.frm");

    let source_file = SourceFile::decode_with_replacement("FormPhysics.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'formPhysics.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'formPhysics.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn gradient_2d_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Gradient-2D/Gradient.frm");

    let source_file = SourceFile::decode_with_replacement("Gradient.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Gradient.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Gradient.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn grayscale_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Grayscale-effect/Grayscale.frm");

    let source_file = SourceFile::decode_with_replacement("Grayscale.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Grayscale.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Grayscale.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn hidden_markov_model_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Hidden-Markov-model/frmHMM.frm");

    let source_file = SourceFile::decode_with_replacement("frmHMM.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'frmHMM.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'frmHMM.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn histograms_advanced_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Histograms-advanced/Histogram.frm");

    let source_file = SourceFile::decode_with_replacement("Histogram.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Histogram.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Histogram.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn histograms_basic_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Histograms-basic/Histogram.frm");

    let source_file = SourceFile::decode_with_replacement("Histogram.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Histogram.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Histogram.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn levels_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Levels-effect/Main.frm");

    let source_file = SourceFile::decode_with_replacement("Main.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Main.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Main.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn mandelbrot_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Mandelbrot/Mandelbrot.frm");

    let source_file = SourceFile::decode_with_replacement("Mandelbrot.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Mandelbrot.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Mandelbrot.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn map_editor_2d_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Map-editor-2D/Main Editor.frm");

    let source_file = SourceFile::decode_with_replacement("Main Editor.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Main Editor.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Main Editor.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn nature_effects_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Nature-effects/NatureFilters.frm");

    let source_file = SourceFile::decode_with_replacement("NatureFilters.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'NatureFilters.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'NatureFilters.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn randomize_effects_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Randomize-effects/RandomizationFX.frm");

    let source_file = SourceFile::decode_with_replacement("RandomizationFX.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'RandomizationFX.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'RandomizationFX.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn scanner_twain_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Scanner-TWAIN/frmScanner.frm");

    let source_file = SourceFile::decode_with_replacement("frmScanner.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'frmScanner.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'frmScanner.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn screen_capture_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Screen-capture/FormScreenCapture.frm");

    let source_file = SourceFile::decode_with_replacement("FormScreenCapture.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'FormScreenCapture.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'FormScreenCapture.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn sepia_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Sepia-effect/Sepia.frm");

    let source_file = SourceFile::decode_with_replacement("Sepia.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Sepia.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Sepia.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn threshold_effect_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Threshold-effect/Threshold.frm");

    let source_file = SourceFile::decode_with_replacement("Threshold.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'Threshold.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'Threshold.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn transparency_2d_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Transparency-2D/frmTransparency.frm");

    let source_file = SourceFile::decode_with_replacement("frmTransparency.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            eprintln!("{e:?}");
            panic!("Failed to decode 'frmTransparency.frm' source file");
        }
    };

    let form_result = FormFile::parse(&source_file);

    if form_result.has_failures() {
        for failure in form_result.failures() {
            failure.eprint();
        }
        panic!("Failed to parse 'frmTransparency.frm' form file");
    }

    let form_file = form_result.unwrap();
    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../snapshots/tests/form");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}
