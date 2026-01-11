use vb6parse::{parse, tokenize, SourceStream, TokenStream};

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

struct ParserBenchmark {
    name: String,
    data: Vec<u8>,
    size_category: FileSize,
}

impl ParserBenchmark {
    fn new(name: &str, data: &[u8]) -> Self {
        let size_category = FileSize::categorize(data);
        Self {
            name: name.to_owned(),
            data: data.to_vec(),
            size_category,
        }
    }
}

fn parser_benchmarks(criterion: &mut Criterion) {
    let files = vec![
        // Small files - simple CST construction
        ParserBenchmark::new(
            "Physics_Logic.bas",
            include_bytes!("../tests/data/vb6-code/Game-physics-basic/Physics_Logic.bas"),
        ),
        ParserBenchmark::new(
            "Declarations.bas",
            include_bytes!("../tests/data/vb6-code/Artificial-life/Declarations.bas"),
        ),
        ParserBenchmark::new(
            "Subs.bas",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Subs.bas"),
        ),
        // Medium files - typical complexity
        ParserBenchmark::new(
            "mod_PublicVars.bas",
            include_bytes!("../tests/data/vb6-code/Levels-effect/mod_PublicVars.bas"),
        ),
        ParserBenchmark::new(
            "MapEditor/Declarations.bas",
            include_bytes!("../tests/data/vb6-code/Map-editor-2D/Declarations.bas"),
        ),
        ParserBenchmark::new(
            "FastDrawing.cls",
            include_bytes!("../tests/data/vb6-code/Levels-effect/FastDrawing.cls"),
        ),
        ParserBenchmark::new(
            "FormPhysics.frm",
            include_bytes!("../tests/data/vb6-code/Game-physics-basic/FormPhysics.frm"),
        ),
        // Large files - complex CST
        ParserBenchmark::new(
            "Organism.cls",
            include_bytes!("../tests/data/vb6-code/Artificial-life/Organism.cls"),
        ),
        ParserBenchmark::new(
            "cCommonDialog.cls",
            include_bytes!("../tests/data/vb6-code/Randomize-effects/cCommonDialog.cls"),
        ),
        ParserBenchmark::new(
            "pdOpenSaveDialog.cls",
            include_bytes!("../tests/data/vb6-code/Levels-effect/pdOpenSaveDialog.cls"),
        ),
        ParserBenchmark::new(
            "Levels/Main.frm",
            include_bytes!("../tests/data/vb6-code/Levels-effect/Main.frm"),
        ),
        ParserBenchmark::new(
            "Curves.frm",
            include_bytes!("../tests/data/vb6-code/Curves-effect/Curves.frm"),
        ),
    ];

    let mut group = criterion.benchmark_group("parser");

    for file in &files {
        let benchmark_name = format!("{}/{}", file.size_category.as_str(), file.name);

        // Pre-tokenize for parser benchmarks (isolates parser performance)
        let source = String::from_utf8_lossy(&file.data);
        let mut source_stream = SourceStream::new(&file.name, &source);
        let tokenize_result = tokenize(&mut source_stream);
        let (tokens_opt, _failures) = tokenize_result.unpack();
        let token_stream = tokens_opt.expect("Failed to tokenize");
        let tokens = token_stream.tokens().clone();

        group.throughput(Throughput::Bytes(file.data.len() as u64));
        group.bench_with_input(
            BenchmarkId::from_parameter(&benchmark_name),
            &(&file.name, &tokens),
            |bench, (name, tokens)| {
                bench.iter(|| {
                    // Clone tokens for each iteration since parse consumes TokenStream
                    let token_stream = TokenStream::new((*name).clone(), (*tokens).clone());
                    black_box(parse(black_box(token_stream)));
                });
            },
        );
    }

    group.finish();
}

criterion_group!(benches, parser_benchmarks);
criterion_main!(benches);
