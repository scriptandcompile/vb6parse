use vb6parse::{ProjectFile, SourceFile};

use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use std::hint::black_box;

// Define file size categories
#[derive(Debug)]
enum FileSize {
    Small,  // < 2KB
    Medium, // 2-10KB
    Large,  // > 10KB
}

impl FileSize {
    fn categorize(bytes: &[u8]) -> Self {
        let size = bytes.len();
        if size < 2048 {
            FileSize::Small
        } else if size < 10240 {
            FileSize::Medium
        } else {
            FileSize::Large
        }
    }

    fn as_str(&self) -> &str {
        match self {
            FileSize::Small => "small",
            FileSize::Medium => "medium",
            FileSize::Large => "large",
        }
    }
}

struct ProjectBenchmark {
    name: String,
    data: Vec<u8>,
    size_category: FileSize,
}

impl ProjectBenchmark {
    fn new(name: &str, data: &[u8]) -> Self {
        let size_category = FileSize::categorize(data);
        Self {
            name: name.to_owned(),
            data: data.to_vec(),
            size_category,
        }
    }
}

fn project_benchmarks(criterion: &mut Criterion) {
    let projects = vec![
        ProjectBenchmark::new(
            "Artificial Life.vbp",
            include_bytes!("../tests/data/vb6-code/Artificial-life/Artificial Life.vbp"),
        ),
        ProjectBenchmark::new(
            "Blacklight.vbp",
            include_bytes!("../tests/data/vb6-code/Blacklight-effect/Blacklight.vbp"),
        ),
        ProjectBenchmark::new(
            "Brightness.vbp",
            include_bytes!(
                "../tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp"
            ),
        ),
        ProjectBenchmark::new(
            "Brightness2.vbp",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.vbp"),
        ),
        ProjectBenchmark::new(
            "Brightness3.vbp",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.vbp"),
        ),
        ProjectBenchmark::new(
            "ShiftColor.vbp",
            include_bytes!("../tests/data/vb6-code/Color-shift-effect/ShiftColor.vbp"),
        ),
        ProjectBenchmark::new(
            "Colorize.vbp",
            include_bytes!("../tests/data/vb6-code/Colorize-effect/Colorize.vbp"),
        ),
        ProjectBenchmark::new(
            "Contrast.vbp",
            include_bytes!("../tests/data/vb6-code/Contrast-effect/Contrast.vbp"),
        ),
        ProjectBenchmark::new(
            "Curves.vbp",
            include_bytes!("../tests/data/vb6-code/Curves-effect/Curves.vbp"),
        ),
        ProjectBenchmark::new(
            "CustomFilters.vbp",
            include_bytes!("../tests/data/vb6-code/Custom-image-filters/CustomFilters.vbp"),
        ),
        ProjectBenchmark::new(
            "Diffuse.vbp",
            include_bytes!("../tests/data/vb6-code/Diffuse-effect/Diffuse.vbp"),
        ),
        ProjectBenchmark::new(
            "EdgeDetection.vbp",
            include_bytes!("../tests/data/vb6-code/Edge-detection/EdgeDetection.vbp"),
        ),
        ProjectBenchmark::new(
            "EmbossEngrave.vbp",
            include_bytes!("../tests/data/vb6-code/Emboss-engrave-effect/EmbossEngrave.vbp"),
        ),
        ProjectBenchmark::new(
            "Fill_Region.vbp",
            include_bytes!("../tests/data/vb6-code/Fill-image-region/Fill_Region.vbp"),
        ),
        ProjectBenchmark::new(
            "FlameTest.vbp",
            include_bytes!("../tests/data/vb6-code/Fire-effect/FlameTest.vbp"),
        ),
        ProjectBenchmark::new(
            "Physics.vbp",
            include_bytes!("../tests/data/vb6-code/Game-physics-basic/Physics.vbp"),
        ),
        ProjectBenchmark::new(
            "Gradient.vbp",
            include_bytes!("../tests/data/vb6-code/Gradient-2D/Gradient.vbp"),
        ),
        ProjectBenchmark::new(
            "Grayscale.vbp",
            include_bytes!("../tests/data/vb6-code/Grayscale-effect/Grayscale.vbp"),
        ),
        ProjectBenchmark::new(
            "HMM.vbp",
            include_bytes!("../tests/data/vb6-code/Hidden-Markov-model/HMM.vbp"),
        ),
        ProjectBenchmark::new(
            "Advanced Histograms.vbp",
            include_bytes!("../tests/data/vb6-code/Histograms-advanced/Advanced Histograms.vbp"),
        ),
        ProjectBenchmark::new(
            "Basic Histograms.vbp",
            include_bytes!("../tests/data/vb6-code/Histograms-basic/Basic Histograms.vbp"),
        ),
        ProjectBenchmark::new(
            "Image Levels.vbp",
            include_bytes!("../tests/data/vb6-code/Levels-effect/Image Levels.vbp"),
        ),
        ProjectBenchmark::new(
            "Mandelbrot.vbp",
            include_bytes!("../tests/data/vb6-code/Mandelbrot/Mandelbrot.vbp"),
        ),
        ProjectBenchmark::new(
            "Map Editor.vbp",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Map Editor.vbp"),
        ),
        ProjectBenchmark::new(
            "NatureFilters.vbp",
            include_bytes!("../tests/data/vb6-code/Nature-effects/NatureFilters.vbp"),
        ),
        ProjectBenchmark::new(
            "RandomizationFX.vbp",
            include_bytes!("../tests/data/vb6-code/Randomize-effects/RandomizationFX.vbp"),
        ),
        ProjectBenchmark::new(
            "VB_Scanner_Support.vbp",
            include_bytes!("../tests/data/vb6-code/Scanner-TWAIN/VB_Scanner_Support.vbp"),
        ),
        ProjectBenchmark::new(
            "ScreenCapture.vbp",
            include_bytes!("../tests/data/vb6-code/Screen-capture/ScreenCapture.vbp"),
        ),
        ProjectBenchmark::new(
            "Sepia.vbp",
            include_bytes!("../tests/data/vb6-code/Sepia-effect/Sepia.vbp"),
        ),
        ProjectBenchmark::new(
            "Threshold.vbp",
            include_bytes!("../tests/data/vb6-code/Threshold-effect/Threshold.vbp"),
        ),
        ProjectBenchmark::new(
            "Transparency.vbp",
            include_bytes!("../tests/data/vb6-code/Transparency-2D/Transparency.vbp"),
        ),
    ];

    let mut group = criterion.benchmark_group("project_files");

    for project in &projects {
        let benchmark_name = format!("{}/{}", project.size_category.as_str(), project.name);

        group.throughput(Throughput::Bytes(project.data.len() as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(&benchmark_name),
            &project,
            |bench, project| {
                let source_file =
                    SourceFile::decode_with_replacement(&project.name, project.data.as_slice())
                        .expect("Failed to decode project file");

                bench.iter(|| {
                    black_box(ProjectFile::parse(black_box(&source_file)));
                });
            },
        );
    }

    group.finish();
}

criterion_group!(benches, project_benchmarks);
criterion_main!(benches);
