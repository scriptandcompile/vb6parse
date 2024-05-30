use vb6parse::project::VB6Project;

use criterion::{black_box, criterion_group, criterion_main, Criterion};

fn criterion_benchmark(c: &mut Criterion) {
    let projects = vec![
        include_bytes!("../tests/data/vb6-code/Artificial-life/Artificial Life.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Blacklight-effect/Blacklight.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp")
            .to_vec(),
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
        include_bytes!("../tests/data/vb6-code/Histograms-advanced/Advanced Histograms.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-basic/Basic Histograms.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Levels-effect/Image Levels.vbp").to_vec(),
    ];

    c.bench_function("load multiple projects", |b| {
        b.iter(|| {
            for project in &projects {
                black_box({
                    let _proj = VB6Project::parse(project.as_slice());
                });
            }
        })
    });
}

criterion_group!(benches, criterion_benchmark);
criterion_main!(benches);
