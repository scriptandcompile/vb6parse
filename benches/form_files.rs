use vb6parse::{FormFile, SourceFile};

use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use std::hint::black_box;

// Define file size categories
#[derive(Debug)]
enum FileSize {
    Small,  // < 5KB
    Medium, // 5-20KB
    Large,  // > 20KB
}

impl FileSize {
    fn categorize(bytes: &[u8]) -> Self {
        let size = bytes.len();
        if size < 5120 {
            FileSize::Small
        } else if size < 20480 {
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

struct FormBenchmark {
    name: String,
    data: Vec<u8>,
    size_category: FileSize,
}

impl FormBenchmark {
    fn new(name: &str, data: &[u8]) -> Self {
        let size_category = FileSize::categorize(data);
        Self {
            name: name.to_owned(),
            data: data.to_vec(),
            size_category,
        }
    }
}

fn form_benchmarks(criterion: &mut Criterion) {
    let forms = vec![
        FormBenchmark::new(
            "FormPhysics.frm",
            include_bytes!("../tests/data/vb6-code/Game-physics-basic/FormPhysics.frm"),
        ),
        FormBenchmark::new(
            "Levels/Histogram.frm",
            include_bytes!("../tests/data/vb6-code/Levels-effect/Histogram.frm"),
        ),
        FormBenchmark::new(
            "Levels/Main.frm",
            include_bytes!("../tests/data/vb6-code/Levels-effect/Main.frm"),
        ),
        FormBenchmark::new(
            "Contrast.frm",
            include_bytes!("../tests/data/vb6-code/Contrast-effect/Contrast.frm"),
        ),
        FormBenchmark::new(
            "Colorize.frm",
            include_bytes!("../tests/data/vb6-code/Colorize-effect/Colorize.frm"),
        ),
        FormBenchmark::new(
            "CustomFilters.frm",
            include_bytes!("../tests/data/vb6-code/Custom-image-filters/CustomFilters.frm"),
        ),
        FormBenchmark::new(
            "frmHMM.frm",
            include_bytes!("../tests/data/vb6-code/Hidden-Markov-model/frmHMM.frm"),
        ),
        FormBenchmark::new(
            "Grayscale.frm",
            include_bytes!("../tests/data/vb6-code/Grayscale-effect/Grayscale.frm"),
        ),
        FormBenchmark::new(
            "frmScanner.frm",
            include_bytes!("../tests/data/vb6-code/Scanner-TWAIN/frmScanner.frm"),
        ),
        FormBenchmark::new(
            "ShiftColors.frm",
            include_bytes!("../tests/data/vb6-code/Color-shift-effect/ShiftColors.frm"),
        ),
        FormBenchmark::new(
            "frmFill.frm",
            include_bytes!("../tests/data/vb6-code/Fill-image-region/frmFill.frm"),
        ),
        FormBenchmark::new(
            "frmTransparency.frm",
            include_bytes!("../tests/data/vb6-code/Transparency-2D/frmTransparency.frm"),
        ),
        FormBenchmark::new(
            "EdgeDetection.frm",
            include_bytes!("../tests/data/vb6-code/Edge-detection/EdgeDetection.frm"),
        ),
        FormBenchmark::new(
            "NatureFilters.frm",
            include_bytes!("../tests/data/vb6-code/Nature-effects/NatureFilters.frm"),
        ),
        FormBenchmark::new(
            "FormScreenCapture.frm",
            include_bytes!("../tests/data/vb6-code/Screen-capture/FormScreenCapture.frm"),
        ),
        FormBenchmark::new(
            "HistogramsAdv/Histogram.frm",
            include_bytes!("../tests/data/vb6-code/Histograms-advanced/Histogram.frm"),
        ),
        FormBenchmark::new(
            "HistogramsAdv/Main.frm",
            include_bytes!("../tests/data/vb6-code/Histograms-advanced/Main.frm"),
        ),
        FormBenchmark::new(
            "Sepia.frm",
            include_bytes!("../tests/data/vb6-code/Sepia-effect/Sepia.frm"),
        ),
        FormBenchmark::new(
            "frmFire.frm",
            include_bytes!("../tests/data/vb6-code/Fire-effect/frmFire.frm"),
        ),
        FormBenchmark::new(
            "RandomizationFX.frm",
            include_bytes!("../tests/data/vb6-code/Randomize-effects/RandomizationFX.frm"),
        ),
        FormBenchmark::new(
            "Brightness/DIBs/Brightness3.frm",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.frm"),
        ),
        FormBenchmark::new(
            "Brightness/FasterDIBs/Brightness.frm",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 4 - Even faster DIBs/Brightness.frm"),
        ),
        FormBenchmark::new(
            "Brightness/API/Brightness2.frm",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.frm"),
        ),
        FormBenchmark::new(
            "Brightness/VB6/Brightness.frm",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.frm"),
        ),
        FormBenchmark::new(
            "MapEditor/FrmResize.frm",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/FrmResize.frm"),
        ),
        FormBenchmark::new(
            "MapEditor/Main Editor.frm",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Main Editor.frm"),
        ),
        FormBenchmark::new(
            "Histograms/Histogram.frm",
            include_bytes!("../tests/data/vb6-code/Histograms-basic/Histogram.frm"),
        ),
        FormBenchmark::new(
            "Histograms/Main.frm",
            include_bytes!("../tests/data/vb6-code/Histograms-basic/Main.frm"),
        ),
        FormBenchmark::new(
            "EmbossEngrave.frm",
            include_bytes!("../tests/data/vb6-code/Emboss-engrave-effect/EmbossEngrave.frm"),
        ),
        FormBenchmark::new(
            "Blacklight.frm",
            include_bytes!("../tests/data/vb6-code/Blacklight-effect/Blacklight.frm"),
        ),
        FormBenchmark::new(
            "Gradient.frm",
            include_bytes!("../tests/data/vb6-code/Gradient-2D/Gradient.frm"),
        ),
        FormBenchmark::new(
            "Mandelbrot.frm",
            include_bytes!("../tests/data/vb6-code/Mandelbrot/Mandelbrot.frm"),
        ),
        FormBenchmark::new(
            "Curves.frm",
            include_bytes!("../tests/data/vb6-code/Curves-effect/Curves.frm"),
        ),
        FormBenchmark::new(
            "frmMain.frm",
            include_bytes!("../tests/data/vb6-code/Artificial-life/frmMain.frm"),
        ),
    ];

    let mut group = criterion.benchmark_group("form_files");

    for form in &forms {
        let benchmark_name = format!(
            "{}/{}",
            form.size_category.as_str(),
            form.name
        );

        group.throughput(Throughput::Bytes(form.data.len() as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(&benchmark_name),
            &form,
            |bench, form| {
                let source_file = SourceFile::decode_with_replacement(
                    &form.name,
                    form.data.as_slice(),
                )
                .expect("Failed to decode form file");

                bench.iter(|| {
                    black_box(FormFile::parse(black_box(&source_file)));
                });
            },
        );
    }

    group.finish();
}

criterion_group!(benches, form_benchmarks);
criterion_main!(benches);
