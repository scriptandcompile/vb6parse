# VB6Parse Fuzzing

This directory contains fuzz testing targets for the vb6parse library using `cargo-fuzz` and libFuzzer.

## Setup

1. Install cargo-fuzz (requires nightly Rust):
```bash
cargo install cargo-fuzz
```

2. Ensure you have the nightly toolchain:
```bash
rustup install nightly
```

## Running Fuzzers

List available fuzz targets:
```bash
cargo +nightly fuzz list
```

Run a specific fuzzer:
```bash
cargo +nightly fuzz run sourcefile_decode
cargo +nightly fuzz run sourcestream
cargo +nightly fuzz run tokenize
cargo +nightly fuzz run cst_parse
```

Run with a time limit (e.g., 60 seconds):
```bash
cargo +nightly fuzz run sourcefile_decode -- -max_total_time=60
```

Run with specific options:
```bash
cargo +nightly fuzz run sourcefile_decode -- \
    -max_total_time=300 \
    -timeout=60 \
    -rss_limit_mb=4096 \
    -print_final_stats=1
```

## Fuzz Targets

### sourcefile_decode
Tests Windows-1252 decoding robustness with arbitrary byte sequences.

### sourcestream
Tests low-level character stream operations including:
- Character peeking
- Pattern matching
- Forward navigation

### tokenize
Tests the tokenizer with arbitrary text input to find:
- Invalid VB6 syntax handling
- Unterminated string literals
- Edge cases in token parsing

### cst_parse
Tests Concrete Syntax Tree construction with:
- Invalid VB6 syntax patterns
- Mismatched control structures (If/End If, For/Next, etc.)
- Deeply nested code structures
- Incomplete statements
- Complex expressions

## Corpus

The `corpus/` directory contains seed inputs for each fuzzer. These are automatically:
- Expanded during fuzzing as new interesting inputs are discovered
- Used as starting points for mutation-based fuzzing

Initial corpus is seeded from the `tests/data/` directory.

## Crashes

If a fuzzer discovers a crash, it will be saved in `artifacts/<fuzzer_name>/`.

To reproduce a crash:
```bash
cargo +nightly fuzz run <fuzzer_name> artifacts/<fuzzer_name>/<crash_file>
```

To minimize a crash test case:
```bash
cargo +nightly fuzz tmin <fuzzer_name> artifacts/<fuzzer_name>/<crash_file>
```

## Continuous Fuzzing

For continuous fuzzing in CI/CD, see the fuzzing plan in `../Fuzzing.md`.

Recommended continuous fuzzing durations:
- Development: 5-10 minutes per target
- CI/CD (nightly): 30 minutes per target
- Corpus building: 24-48 hours per target

## Coverage

To view coverage from fuzzing:
```bash
cargo +nightly fuzz coverage <fuzzer_name>
```

## References

- [cargo-fuzz book](https://rust-fuzz.github.io/book/cargo-fuzz.html)
- [libFuzzer documentation](https://llvm.org/docs/LibFuzzer.html)
- [Fuzzing.md](../Fuzzing.md) - Complete fuzzing strategy and plan
