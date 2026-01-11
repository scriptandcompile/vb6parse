use vb6parse::{ClassFile, SourceFile};

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

struct ClassBenchmark {
    name: String,
    data: Vec<u8>,
    size_category: FileSize,
}

impl ClassBenchmark {
    fn new(name: &str, data: &[u8]) -> Self {
        let size_category = FileSize::categorize(data);
        Self {
            name: name.to_owned(),
            data: data.to_vec(),
            size_category,
        }
    }
}

fn class_benchmarks(criterion: &mut Criterion) {
    let classes = vec![
        ClassBenchmark::new(
            "FastDrawing.cls",
            include_bytes!("../tests/data/vb6-code/Levels-effect/FastDrawing.cls"),
        ),
        ClassBenchmark::new(
            "pdOpenSaveDialog.cls",
            include_bytes!("../tests/data/vb6-code/Levels-effect/pdOpenSaveDialog.cls"),
        ),
        ClassBenchmark::new(
            "cCommonDialog.cls",
            include_bytes!("../tests/data/vb6-code/Randomize-effects/cCommonDialog.cls"),
        ),
        ClassBenchmark::new(
            "cSystemColorDialog.cls",
            include_bytes!("../tests/data/vb6-code/Emboss-engrave-effect/cSystemColorDialog.cls"),
        ),
        ClassBenchmark::new(
            "Organism.cls",
            include_bytes!("../tests/data/vb6-code/Artificial-life/Organism.cls"),
        ),
    ];

    let mut group = criterion.benchmark_group("class_files");

    for class in &classes {
        let benchmark_name = format!("{}/{}", class.size_category.as_str(), class.name);

        group.throughput(Throughput::Bytes(class.data.len() as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(&benchmark_name),
            &class,
            |bench, class| {
                let source_file =
                    SourceFile::decode_with_replacement(&class.name, class.data.as_slice())
                        .expect("Failed to decode class file");

                bench.iter(|| {
                    black_box(ClassFile::parse(black_box(&source_file)));
                });
            },
        );
    }

    group.finish();
}

criterion_group!(benches, class_benchmarks);
criterion_main!(benches);
