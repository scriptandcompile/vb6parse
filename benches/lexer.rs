use vb6parse::{tokenize, SourceStream};

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

struct LexerBenchmark {
    name: String,
    data: Vec<u8>,
    size_category: FileSize,
}

impl LexerBenchmark {
    fn new(name: &str, data: &[u8]) -> Self {
        let size_category = FileSize::categorize(data);
        Self {
            name: name.to_owned(),
            data: data.to_vec(),
            size_category,
        }
    }
}

fn lexer_benchmarks(criterion: &mut Criterion) {
    let files = vec![
        // Small files - simple tokenization
        LexerBenchmark::new(
            "Physics_Logic.bas",
            include_bytes!("../tests/data/vb6-code/Game-physics-basic/Physics_Logic.bas"),
        ),
        LexerBenchmark::new(
            "Declarations.bas",
            include_bytes!("../tests/data/vb6-code/Artificial-life/Declarations.bas"),
        ),
        LexerBenchmark::new(
            "Subs.bas",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Subs.bas"),
        ),
        // Medium files - typical code complexity
        LexerBenchmark::new(
            "mod_PublicVars.bas",
            include_bytes!("../tests/data/vb6-code/Levels-effect/mod_PublicVars.bas"),
        ),
        LexerBenchmark::new(
            "MapEditor/Declarations.bas",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Declarations.bas"),
        ),
        LexerBenchmark::new(
            "FastDrawing.cls",
            include_bytes!("../tests/data/vb6-code/Levels-effect/FastDrawing.cls"),
        ),
        LexerBenchmark::new(
            "FormPhysics.frm",
            include_bytes!("../tests/data/vb6-code/Game-physics-basic/FormPhysics.frm"),
        ),
        // Large files - complex with many tokens
        LexerBenchmark::new(
            "Organism.cls",
            include_bytes!("../tests/data/vb6-code/Artificial-life/Organism.cls"),
        ),
        LexerBenchmark::new(
            "cCommonDialog.cls",
            include_bytes!("../tests/data/vb6-code/Randomize-effects/cCommonDialog.cls"),
        ),
        LexerBenchmark::new(
            "pdOpenSaveDialog.cls",
            include_bytes!("../tests/data/vb6-code/Levels-effect/pdOpenSaveDialog.cls"),
        ),
        LexerBenchmark::new(
            "Levels/Main.frm",
            include_bytes!("../tests/data/vb6-code/Levels-effect/Main.frm"),
        ),
        LexerBenchmark::new(
            "Curves.frm",
            include_bytes!("../tests/data/vb6-code/Curves-effect/Curves.frm"),
        ),
    ];

    let mut group = criterion.benchmark_group("lexer");

    for file in &files {
        let benchmark_name = format!("{}/{}", file.size_category.as_str(), file.name);

        // Create source string from bytes
        let source = String::from_utf8_lossy(&file.data);

        group.throughput(Throughput::Bytes(file.data.len() as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(&benchmark_name),
            &source,
            |bench, source| {
                bench.iter(|| {
                    let mut source_stream = SourceStream::new(&file.name, source);
                    black_box(tokenize(black_box(&mut source_stream)));
                });
            },
        );
    }

    group.finish();
}

criterion_group!(benches, lexer_benchmarks);
criterion_main!(benches);
