use vb6parse::parsers::VB6ClassFile;
use vb6parse::parsers::VB6Project;
use vb6parse::parsers::VB6Stream;

use criterion::{black_box, criterion_group, criterion_main, Criterion};

fn project_benchmark(c: &mut Criterion) {
    let project_names = vec![
        "Artificial Life.vbp".to_owned(),
        "Blacklight.vbp".to_owned(),
        "Brightness.vbp".to_owned(),
        "Brightness2.vbp".to_owned(),
        "Brightness3.vbp".to_owned(),
        "ShiftColor.vbp".to_owned(),
        "Colorize.vbp".to_owned(),
        "Contrast.vbp".to_owned(),
        "Curves.vbp".to_owned(),
        "CustomFilters.vbp".to_owned(),
        "Diffuse.vbp".to_owned(),
        "EdgeDetection.vbp".to_owned(),
        "EmbossEngrave.vbp".to_owned(),
        "Fill_Region.vbp".to_owned(),
        "FlameTest.vbp".to_owned(),
        "Physics.vbp".to_owned(),
        "Gradient.vbp".to_owned(),
        "Grayscale.vbp".to_owned(),
        "HMM.vbp".to_owned(),
        "Advanced Histograms.vbp".to_owned(),
        "Basic Histograms.vbp".to_owned(),
        "Image Levels.vbp".to_owned(),
        "Mandelbrot.vbp".to_owned(),
        "Map Editor.vbp".to_owned(),
        "NatureFilters.vbp".to_owned(),
        "RandomizationFX.vbp".to_owned(),
        "VB_Scanner_Support.vbp".to_owned(),
        "ScreenCapture.vbp".to_owned(),
        "Sepia.vbp".to_owned(),
        "Threshold.vbp".to_owned(),
        "Transparency.vbp".to_owned(),
    ];

    let projects = vec![
        include_bytes!("../tests/data/vb6-code/Artificial-life/Artificial Life.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Blacklight-effect/Blacklight.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp")
            .to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Color-shift-effect/ShiftColor.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Colorize-effect/Colorize.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Contrast-effect/Contrast.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Curves-effect/Curves.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Custom-image-filters/CustomFilters.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Diffuse-effect/Diffuse.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Edge-detection/EdgeDetection.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Emboss-engrave-effect/EmbossEngrave.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Fill-image-region/Fill_Region.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Fire-effect/FlameTest.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Game-physics-basic/Physics.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Gradient-2D/Gradient.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Grayscale-effect/Grayscale.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Hidden-Markov-model/HMM.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-advanced/Advanced Histograms.vbp")
            .to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-basic/Basic Histograms.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Levels-effect/Image Levels.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Mandelbrot/Mandelbrot.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Map-editor-2D/Map Editor.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Nature-effects/NatureFilters.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Randomize-effects/RandomizationFX.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Scanner-TWAIN/VB_Scanner_Support.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Screen-capture/ScreenCapture.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Sepia-effect/Sepia.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Threshold-effect/Threshold.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Transparency-2D/Transparency.vbp").to_vec(),
    ];

    let project_pairs: Vec<(_, _)> = project_names.iter().zip(projects.iter()).collect();

    c.bench_function("load multiple projects", |b| {
        b.iter(|| {
            for project_pair in &project_pairs {
                black_box({
                    let mut stream = VB6Stream::new(project_pair.0, project_pair.1.as_slice());
                    let _proj = VB6Project::parse(&mut stream);
                });
            }
        })
    });
}

fn class_benchmark(c: &mut Criterion) {
    let class_names = vec![
        "FastDrawing.cls".to_owned(),
        "pdOpenSaveDialog.cls".to_owned(),
        "cCommonDialog.cls".to_owned(),
        "cSystemColorDialog.cls".to_owned(),
    ];

    let classes = vec![
        include_bytes!("../tests/data/vb6-code/Levels-effect/FastDrawing.cls").to_vec(),
        include_bytes!("../tests/data/vb6-code/Levels-effect/pdOpenSaveDialog.cls").to_vec(),
        include_bytes!("../tests/data/vb6-code/Randomize-effects/cCommonDialog.cls").to_vec(),
        include_bytes!("../tests/data/vb6-code/Emboss-engrave-effect/cSystemColorDialog.cls")
            .to_vec(),
        include_bytes!("../tests/data/vb6-code/Artificial-life/Organism.cls").to_vec(),
    ];

    let class_pairs: Vec<(_, _)> = class_names.iter().zip(classes.iter()).collect();

    c.bench_function("load multiple classes", |b| {
        b.iter(|| {
            for class_pair in &class_pairs {
                black_box({
                    let _class =
                        VB6ClassFile::parse(class_pair.0.to_string(), &mut class_pair.1.as_slice());
                });
            }
        })
    });
}

criterion_group!(benches, project_benchmark, class_benchmark);
criterion_main!(benches);
