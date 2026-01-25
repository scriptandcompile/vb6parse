use vb6parse::{ModuleFile, SourceFile};

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

struct ModuleBenchmark {
    name: String,
    data: Vec<u8>,
    size_category: FileSize,
}

impl ModuleBenchmark {
    fn new(name: &str, data: &[u8]) -> Self {
        let size_category = FileSize::categorize(data);
        Self {
            name: name.to_owned(),
            data: data.to_vec(),
            size_category,
        }
    }
}

fn module_benchmarks(criterion: &mut Criterion) {
    let modules = vec![
        ModuleBenchmark::new(
            "Physics_Logic.bas",
            include_bytes!("../tests/data/vb6-code/Game-physics-basic/Physics_Logic.bas"),
        ),
        ModuleBenchmark::new(
            "Levels/mod_PublicVars.bas",
            include_bytes!("../tests/data/vb6-code/Levels-effect/mod_PublicVars.bas"),
        ),
        ModuleBenchmark::new(
            "HistogramsAdv/mod_PublicVars.bas",
            include_bytes!("../tests/data/vb6-code/Histograms-advanced/mod_PublicVars.bas"),
        ),
        ModuleBenchmark::new(
            "MapEditor/Declarations.bas",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Declarations.bas"),
        ),
        ModuleBenchmark::new(
            "MapEditor/Subs.bas",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Subs.bas"),
        ),
        ModuleBenchmark::new(
            "Histograms/mod_PublicVars.bas",
            include_bytes!("../tests/data/vb6-code/Histograms-basic/mod_PublicVars.bas"),
        ),
        ModuleBenchmark::new(
            "ArtificialLife/Declarations.bas",
            include_bytes!("../tests/data/vb6-code/Artificial-life/Declarations.bas"),
        ),
    ];

    let mut group = criterion.benchmark_group("module_files");

    for module in &modules {
        let benchmark_name = format!("{}/{}", module.size_category.as_str(), module.name);

        group.throughput(Throughput::Bytes(module.data.len() as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(&benchmark_name),
            &module,
            |bench, module| {
                let source_file =
                    SourceFile::decode_with_replacement(&module.name, module.data.as_slice())
                        .expect("Failed to decode module file");

                bench.iter(|| {
                    black_box(ModuleFile::parse(black_box(&source_file)));
                });
            },
        );
    }

    group.finish();
}

criterion_group!(benches, module_benchmarks);
criterion_main!(benches);
