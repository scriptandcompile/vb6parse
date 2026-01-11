# VB6Parse Benchmarks

This directory contains performance benchmarks for the VB6Parse library using the Criterion framework.

## Benchmark Organization

The benchmarks are organized by file type, with each benchmark file focusing on a specific aspect of the parser:

- **`project_files.rs`** - Benchmarks for `.vbp` (VB6 project) file parsing
- **`class_files.rs`** - Benchmarks for `.cls` (VB6 class) file parsing
- **`module_files.rs`** - Benchmarks for `.bas` (VB6 module) file parsing
- **`form_files.rs`** - Benchmarks for `.frm` (VB6 form) file parsing

## File Size Categories

Benchmarks are categorized by file size to help identify performance characteristics across different input sizes:

### Projects, Classes, and Modules
- **Small**: < 2KB (simple files with minimal code)
- **Medium**: 2-10KB (typical VB6 files)
- **Large**: > 10KB (complex files with extensive code)

### Forms
- **Small**: < 5KB (simple forms with few controls)
- **Medium**: 5-20KB (typical forms with moderate complexity)
- **Large**: > 20KB (complex forms with many controls)

## Running Benchmarks

```bash
# Run all benchmarks
cargo bench

# Run benchmarks for a specific file type
cargo bench --bench project_files
cargo bench --bench class_files
cargo bench --bench module_files
cargo bench --bench form_files

# Run benchmarks matching a pattern
cargo bench -- small    # Only small files
cargo bench -- medium   # Only medium files
cargo bench -- large    # Only large files

# Save results as a baseline for comparison
cargo bench -- --save-baseline main

# Compare against a baseline
cargo bench -- --baseline main

# Generate HTML reports (enabled by default)
# Reports are generated in target/criterion/
```

## Understanding the Results

Each benchmark measures:
- **Mean**: Average execution time
- **Median**: Middle value of execution times
- **Std Dev**: Standard deviation (consistency indicator)
- **Throughput**: Bytes processed per second

Lower times and higher throughput indicate better performance.

## Benchmark Naming Convention

Benchmarks follow the naming pattern: `{file_type}/{size_category}/{file_name}`

Examples:
- `project_files/small/Blacklight.vbp`
- `class_files/medium/FastDrawing.cls`
- `module_files/large/Physics_Logic.bas`
- `form_files/large/MapEditor/Main Editor.frm`

## Viewing Results

After running benchmarks, view the HTML report:
```bash
# On Linux/macOS
open target/criterion/report/index.html

# Or navigate to the file in your browser
```

The report includes:
- Performance comparisons between runs
- Statistical analysis
- Violin plots showing distribution
- Detailed timing breakdowns

## Adding New Benchmarks

To add a new benchmark:

1. Choose the appropriate benchmark file based on file type
2. Add your test file to the data structure:
   ```rust
   ClassBenchmark::new(
       "MyNewClass.cls",
       include_bytes!("../tests/data/path/to/MyNewClass.cls"),
   ),
   ```
3. The file will be automatically categorized by size
4. Run `cargo bench` to verify

## CI Integration

Benchmarks can be run in CI to detect performance regressions:
```bash
# Establish baseline on main branch
cargo bench -- --save-baseline main

# On feature branch, compare against baseline
cargo bench -- --baseline main
```

## Notes

- All test data is embedded at compile time using `include_bytes!()`
- This ensures deterministic benchmarking
- `black_box()` is used to prevent compiler optimizations from skewing results
- Throughput measurements help identify bottlenecks in parsing large files
