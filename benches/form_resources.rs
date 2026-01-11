use vb6parse::FormResourceFile;

use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use std::hint::black_box;

// Define file size categories
#[derive(Debug)]
enum FileSize {
    Small,  // < 500 bytes
    Medium, // 500B-5KB
    Large,  // > 5KB
}

impl FileSize {
    fn categorize(bytes: &[u8]) -> Self {
        let size = bytes.len();
        if size < 500 {
            FileSize::Small
        } else if size < 5120 {
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

struct ResourceBenchmark {
    name: String,
    data: Vec<u8>,
    size_category: FileSize,
}

impl ResourceBenchmark {
    fn new(name: &str, data: &[u8]) -> Self {
        let size_category = FileSize::categorize(data);
        Self {
            name: name.to_owned(),
            data: data.to_vec(),
            size_category,
        }
    }
}

fn form_resource_benchmarks(criterion: &mut Criterion) {
    let files = vec![
        // Small resources - minimal data
        ResourceBenchmark::new(
            "FormPhysics.frx",
            include_bytes!("../tests/data/vb6-code/Game-physics-basic/FormPhysics.frx"),
        ),
        ResourceBenchmark::new(
            "Gradient.frx",
            include_bytes!("../tests/data/vb6-code/Gradient-2D/Gradient.frx"),
        ),
        ResourceBenchmark::new(
            "frmTransparency.frx",
            include_bytes!("../tests/data/vb6-code/Transparency-2D/frmTransparency.frx"),
        ),
        // Medium resources - typical icons/images
        ResourceBenchmark::new(
            "Brightness/VB6/Brightness.frx",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.frx"),
        ),
        ResourceBenchmark::new(
            "Brightness/API/Brightness2.frx",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.frx"),
        ),
        ResourceBenchmark::new(
            "Brightness/DIBs/Brightness3.frx",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.frx"),
        ),
        ResourceBenchmark::new(
            "Brightness/FasterDIBs/Brightness.frx",
            include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 4 - Even faster DIBs/Brightness.frx"),
        ),
        // Large resources - complex forms with many images
        ResourceBenchmark::new(
            "MapEditor/Main Editor.frx",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Main Editor.frx"),
        ),
        ResourceBenchmark::new(
            "frmHMM.frx",
            include_bytes!("../tests/data/vb6-code/Hidden-Markov-model/frmHMM.frx"),
        ),
        ResourceBenchmark::new(
            "frmFire.frx",
            include_bytes!("../tests/data/vb6-code/Fire-effect/frmFire.frx"),
        ),
    ];

    let mut group = criterion.benchmark_group("form_resources");

    for file in &files {
        let benchmark_name = format!("{}/{}", file.size_category.as_str(), file.name);

        group.throughput(Throughput::Bytes(file.data.len() as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(&benchmark_name),
            &file,
            |bench, file| {
                bench.iter(|| {
                    black_box(FormResourceFile::parse(
                        black_box(&file.name),
                        black_box(file.data.clone()),
                    ));
                });
            },
        );
    }

    group.finish();
}

criterion_group!(benches, form_resource_benchmarks);
criterion_main!(benches);
