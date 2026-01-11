# VB6Parse Fuzzing

This directory contains comprehensive fuzz testing for the vb6parse library using `cargo-fuzz` and libFuzzer. Fuzzing helps discover edge cases, malformed input handling, and potential panics that traditional unit tests might miss.

## Why Fuzz VB6Parse?

VB6Parse processes untrusted input from legacy VB6 projects, making robustness critical:

- **Legacy code variety**: Real-world VB6 projects contain encoding quirks, malformed syntax, and IDE-generated edge cases
- **Binary formats**: FRX files are binary with multiple header formats that must be parsed safely
- **Partial success model**: Parsers should handle malformed input gracefully, returning partial results rather than panicking
- **Complex state machines**: Tokenization and CST construction involve intricate state transitions that benefit from mutation testing

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

List all available fuzz targets:
```bash
cargo +nightly fuzz list
```

Run a specific fuzzer (replace `<target>` with any target below):
```bash
# I/O Layer targets
cargo +nightly fuzz run sourcefile_decode
cargo +nightly fuzz run sourcestream

# Lexer Layer target
cargo +nightly fuzz run tokenize

# Parsers Layer target
cargo +nightly fuzz run cst_parse

# Files Layer targets (high-level parsers)
cargo +nightly fuzz run project_file
cargo +nightly fuzz run class_file
cargo +nightly fuzz run module_file
cargo +nightly fuzz run form_file
cargo +nightly fuzz run form_resource
```

Run with a time limit (recommended for development):
```bash
# Run for 60 seconds
cargo +nightly fuzz run sourcefile_decode -- -max_total_time=60

# Run for 5 minutes
cargo +nightly fuzz run cst_parse -- -max_total_time=300
```

Run with specific options for deeper fuzzing:
```bash
cargo +nightly fuzz run form_resource -- \
    -max_total_time=3600 \
    -timeout=60 \
    -rss_limit_mb=4096 \
    -print_final_stats=1 \
    -jobs=4
```

Common libFuzzer options:
- `-max_total_time=<seconds>`: Stop after N seconds
- `-timeout=<seconds>`: Timeout for individual test cases (default: 1200)
- `-rss_limit_mb=<MB>`: Memory limit per test case
- `-jobs=<N>`: Number of parallel fuzzing jobs
- `-print_final_stats=1`: Show statistics when done
- `-dict=<file>`: Use a dictionary for mutation guidance

## Fuzz Targets

VB6Parse has 9 fuzz targets covering all layers of the parsing pipeline:

### Layer 1: I/O Layer

#### `sourcefile_decode`
Tests Windows-1252 decoding and character encoding robustness.

**What it tests:**
- Arbitrary byte sequences that may not be valid Windows-1252
- Invalid UTF-8 sequences
- Null bytes and control characters
- Encoding boundary conditions
- Replacement character insertion for invalid bytes

**Why it matters:** VB6 projects use Windows-1252 encoding, which has undefined bytes in certain ranges. The decoder must handle these gracefully.

#### `sourcestream`
Tests low-level character stream navigation and pattern matching.

**What it tests:**
- Character peeking at arbitrary offsets
- Pattern matching with malformed patterns
- Forward/backward navigation edge cases
- Case-insensitive comparisons
- Line/column tracking accuracy
- Offset boundary conditions

**Why it matters:** SourceStream is the foundation for all parsing. Bugs here affect everything built on top.

### Layer 2: Lexer Layer

#### `tokenize`
Tests tokenization of arbitrary VB6-like text input.

**What it tests:**
- Invalid VB6 syntax combinations
- Unterminated string literals
- Malformed numeric literals (e.g., `&H`, `&O` with no digits)
- Line continuation edge cases (`_` at unexpected positions)
- Comment handling (single quote, `Rem` statement)
- Keyword vs identifier ambiguity
- Whitespace and newline handling

**Why it matters:** The tokenizer must never panic on malformed source, even when syntax is completely invalid.

### Layer 3: Parsers Layer

#### `cst_parse`
Tests Concrete Syntax Tree construction from token streams.

**What it tests:**
- Invalid VB6 syntax patterns
- Mismatched control structures (`If` without `End If`, `For` without `Next`)
- Deeply nested code structures (potential stack overflow)
- Incomplete statements and expressions
- Complex expressions with unusual operator combinations
- Unexpected token sequences
- Missing required tokens

**Why it matters:** CST construction involves complex state machines. Fuzzing helps find unexpected token sequences that could cause panics or infinite loops.

### Layer 4: Files Layer (High-Level Parsers)

#### `project_file`
Tests VB6 project file (`.vbp`) parsing.

**What it tests:**
- Malformed project file syntax
- Invalid property names and values
- Missing required sections (e.g., `Type=`, `Form=`)
- Duplicate entries
- Incorrect reference formats
- Version number edge cases

**Why it matters:** Project files are the entry point to VB6 codebases. Robust parsing ensures the library can handle projects from any VB6 version or IDE quirk.

#### `class_file`
Tests VB6 class module (`.cls`) parsing.

**What it tests:**
- Malformed `VERSION` lines
- Invalid `Attribute` statements
- Missing or duplicate `CLASS` attribute
- Properties: `MultiUse`, `Persistable`, `DataBindingBehavior`, `DataSourceBehavior`
- Combination of header and code parsing
- Invalid VB6 code in class body

**Why it matters:** Class files have a unique header structure that differs from modules. Fuzzing ensures robust handling of all class-specific properties.

#### `module_file`
Tests VB6 standard module (`.bas`) parsing.

**What it tests:**
- Malformed `VERSION` lines in modules
- Invalid `Attribute VB_Name` statements
- Module-level variable declarations
- Public/Private procedure definitions
- Option statements (`Option Explicit`, `Option Base`)
- Invalid code in module body

**Why it matters:** Modules are the simplest VB6 file type, but still have header/body structure that must be parsed correctly.

#### `form_file`
Tests VB6 form file (`.frm`) parsing - the most complex file type.

**What it tests:**
- Form header with control hierarchy
- Nested control structures (Forms → Frames → Controls)
- Property parsing for 50+ control types
- Menu control definitions
- Begin/End block matching
- Combination of visual designer output and VB6 code
- Missing or malformed control properties
- Invalid control types

**Why it matters:** Form files are the most complex VB6 file type, containing both visual designer output and code. They have the most edge cases and IDE-generated quirks.

#### `form_resource`
Tests VB6 form resource file (`.frx`) parsing - pure binary format.

**What it tests:**
- Invalid binary data sequences
- Multiple FRX header formats (12-byte, 8-byte, 4-byte, 3-byte, 1-byte)
- Corrupted header fields
- Entry size mismatches
- Property GUID lookups
- String data with invalid encoding
- Binary blob handling (icons, images, etc.)
- Truncated files and incomplete entries

**Why it matters:** FRX files are binary with multiple header formats used across VB6 versions. This is the most crash-prone area, making fuzzing essential.

## Corpus Management

The `corpus/` directory contains seed inputs for each fuzzer. The corpus is crucial for effective fuzzing:

### How It Works

1. **Seed corpus**: Initial inputs in `corpus/<target>/` provide starting points for mutation
2. **Automatic expansion**: LibFuzzer discovers new "interesting" inputs during fuzzing and adds them to the corpus
3. **Coverage-guided**: Inputs that trigger new code paths are kept; redundant ones are discarded
4. **Persistent**: Corpus grows over time, improving fuzzing effectiveness in future runs

### Corpus Sources

Initial corpus is seeded from:
- `tests/data/`: Real VB6 project files (submodules)
- Hand-crafted edge cases
- Previously discovered crash cases (minimized)

### Corpus Growth

As you fuzz, the corpus automatically grows:
```bash
# Before fuzzing
$ ls corpus/form_file/ | wc -l
15

# After fuzzing for 1 hour
$ ls corpus/form_file/ | wc -l
247
```

### Managing Corpus

View corpus statistics:
```bash
# Count corpus entries
ls -1 corpus/<target>/ | wc -l

# Show total corpus size
du -sh corpus/<target>/
```

Minimize corpus (remove redundant entries):
```bash
cargo +nightly fuzz cmin <target>
```

Merge corpus from multiple runs:
```bash
# Merge corpus from another machine or CI run
cargo +nightly fuzz cmin <target> -- corpus/<target>/ other_corpus/<target>/
```

## Handling Crashes and Failures

### When a Crash Occurs

If a fuzzer discovers a crash or timeout, artifacts are saved in `artifacts/<fuzzer_name>/`:

```
artifacts/
├── form_resource/
│   ├── crash-da39a3ee5e6b4b0d  # Crash-causing input
│   ├── timeout-8b3f9c1a7d2e4f  # Input that caused timeout
│   └── ...
```

### Reproducing Crashes

Run the fuzzer with the crash file to reproduce:
```bash
cargo +nightly fuzz run form_resource artifacts/form_resource/crash-da39a3ee5e6b4b0d
```

This will:
1. Load the crash-causing input
2. Re-run the fuzzer with that exact input
3. Show the panic/error message and stack trace

### Minimizing Crashes

Crash inputs often contain redundant bytes. Minimize them for easier debugging:

```bash
cargo +nightly fuzz tmin form_resource artifacts/form_resource/crash-da39a3ee5e6b4b0d
```

This produces the smallest input that still triggers the crash, making it easier to:
- Understand root cause
- Write a minimal reproduction test case
- Fix the bug

### Analyzing Crashes

1. **Reproduce** to confirm the crash
2. **Minimize** to get the smallest crashing input
3. **Debug** with the minimized input:
   ```bash
   # Run with debugger
   rust-lldb target/x86_64-unknown-linux-gnu/release/form_resource artifacts/form_resource/crash-minimized
   
   # Or add print debugging to the fuzz target
   ```
4. **Create test case** from the crash to prevent regression
5. **Fix the bug** in the parser
6. **Verify** the fix:
   ```bash
   cargo +nightly fuzz run form_resource artifacts/form_resource/crash-da39a3ee5e6b4b0d -- -runs=1
   ```
7. **Delete artifact** once fixed

### Recent Failures

See `Recent_Failures.md` for a log of recent fuzzing discoveries and their fixes.

## Fuzzing Strategy

### Ad-Hoc Development Fuzzing

For day-to-day development, quick fuzzing sessions help catch issues early:

```bash
# Quick smoke test (1 minute per target)
for target in sourcefile_decode sourcestream tokenize cst_parse project_file class_file module_file form_file form_resource; do
    echo "Fuzzing $target..."
    cargo +nightly fuzz run $target -- -max_total_time=60
done
```

**When to fuzz during development:**
- After implementing a new parser or significant refactor
- Before committing changes to critical parsing code
- When fixing a bug to ensure the fix doesn't introduce new issues
- After updating dependencies that affect parsing logic

### Deep Fuzzing Sessions

For thorough testing, run longer sessions on specific targets:

```bash
# Focus on the most complex parsers
cargo +nightly fuzz run form_file -- -max_total_time=3600      # 1 hour
cargo +nightly fuzz run form_resource -- -max_total_time=3600  # 1 hour
cargo +nightly fuzz run cst_parse -- -max_total_time=1800      # 30 minutes
```

**Recommended priorities:**
1. **form_resource** - Binary parsing, highest crash risk
2. **form_file** - Most complex file format
3. **cst_parse** - Complex state machine
4. **project_file** - Entry point, critical for library users
5. **tokenize** - Foundation for all parsing

### Continuous Fuzzing

If running in CI/CD or overnight, use longer durations:

```bash
# Overnight session (8 hours per target)
cargo +nightly fuzz run form_resource -- -max_total_time=28800 -jobs=4
```

**Benefits of longer runs:**
- Discover rare edge cases
- Build more comprehensive corpus
- Achieve deeper code coverage
- Find timing-dependent issues

## Coverage Analysis

View code coverage achieved by fuzzing:

```bash
# Generate coverage report
cargo +nightly fuzz coverage form_resource

# View HTML coverage report
open fuzz/coverage/form_resource/index.html
```

This shows:
- Which code paths the fuzzer exercised
- Uncovered branches that might need seed inputs
- Comparison with unit test coverage

**Note:** Fuzzing coverage often differs from unit test coverage:
- Fuzzers discover edge cases unit tests miss
- Some code paths may require specific seeds to reach
- Coverage complements but doesn't replace traditional testing

## Performance Monitoring

Monitor fuzzing performance during runs:

```bash
# Run with statistics
cargo +nightly fuzz run form_file -- -print_final_stats=1

# Watch live progress
cargo +nightly fuzz run form_file -- -print_progress=1
```

**Key metrics:**
- **exec/s**: Executions per second (higher is better, indicates fuzzer efficiency)
- **cov**: Coverage (unique code paths found)
- **corp**: Corpus size (interesting inputs discovered)

**Typical exec/s rates:**
- `sourcefile_decode`: 50,000+ exec/s (simple, fast)
- `tokenize`: 10,000-20,000 exec/s (moderate complexity)
- `form_file`: 1,000-5,000 exec/s (complex parsing)
- `form_resource`: 5,000-10,000 exec/s (binary format)

If exec/s is very low (<100), the fuzzer may be hitting timeouts or slow paths frequently.

## Best Practices

### 1. Start with Quick Runs
Don't commit to long fuzzing sessions initially. Run 1-5 minutes first to catch obvious issues.

### 2. Monitor Memory Usage
Some inputs can cause excessive memory allocation. Set reasonable limits:
```bash
cargo +nightly fuzz run form_file -- -rss_limit_mb=2048
```

### 3. Use Parallel Jobs Carefully
Multiple jobs speed up fuzzing but increase resource usage:
```bash
# Good for overnight runs
cargo +nightly fuzz run form_resource -- -jobs=4 -max_total_time=28800

# Bad: too many jobs can thrash CPU
cargo +nightly fuzz run form_resource -- -jobs=32  # Probably overkill
```

### 4. Preserve Interesting Crashes
When you find a crash:
1. Copy it to a safe location (artifacts can be overwritten)
2. Minimize it immediately
3. Create a test case before deleting

### 5. Build Corpus Over Time
Don't delete corpus entries unless they're truly redundant. A rich corpus makes future fuzzing more effective.

### 6. Focus on High-Risk Areas
Not all fuzz targets need equal attention:
- **High priority**: Binary formats (form_resource), complex parsers (form_file, cst_parse)
- **Medium priority**: File parsers (project_file, class_file, module_file)
- **Lower priority**: Foundation layers (sourcefile_decode, sourcestream) - these are simpler and well-tested

### 7. Combine with Other Testing
Fuzzing complements but doesn't replace:
- Unit tests (specific scenarios)
- Integration tests (real-world files)
- Property-based tests (invariants)
- Manual testing (usability)

## Interpreting Results

### Success Indicators
- ✅ No crashes after reasonable fuzzing time (5+ minutes)
- ✅ Corpus grows steadily then plateaus (coverage maximized)
- ✅ High exec/s rate (fuzzer is efficient)
- ✅ Good coverage of target code

### Warning Signs
- ⚠️ Repeated timeouts (infinite loops or very slow paths)
- ⚠️ Memory limit hits (unbounded allocation)
- ⚠️ Corpus grows without bound (fuzzer finding too many "interesting" inputs)
- ⚠️ Very low exec/s (<100) (fuzzer spending too much time per input)

### When to Stop Fuzzing
- Corpus size stabilizes (no new inputs for 10+ minutes)
- Coverage plateaus (no new code paths discovered)
- Time limit reached
- Acceptable exec count achieved (e.g., 1M+ executions)

## Troubleshooting

### Fuzzer is very slow
- Check if inputs are triggering slow code paths
- Add timeouts: `-timeout=10`
- Profile the fuzz target to find bottlenecks

### Out of memory errors
- Reduce memory limit: `-rss_limit_mb=1024`
- Check for unbounded allocations in parser
- Minimize inputs before fuzzing

### No new coverage
- Corpus may be exhausted for this seed set
- Try running a different target
- Add new seed inputs from real VB6 projects

### Fuzzer finds too many "crashes" that aren't bugs
- Check if these are expected panics (e.g., `unimplemented!()`)
- Adjust parser to return errors instead of panicking
- Use `std::panic::catch_unwind` if intentional panics are acceptable
## Quick Reference

### Common Commands

```bash
# List all targets
cargo +nightly fuzz list

# Quick test (1 minute)
cargo +nightly fuzz run <target> -- -max_total_time=60

# Deep test (1 hour)
cargo +nightly fuzz run <target> -- -max_total_time=3600

# Reproduce crash
cargo +nightly fuzz run <target> artifacts/<target>/<crash_file>

# Minimize crash
cargo +nightly fuzz tmin <target> artifacts/<target>/<crash_file>

# View coverage
cargo +nightly fuzz coverage <target>

# Minimize corpus
cargo +nightly fuzz cmin <target>
```

### Target Priority Order

For ad-hoc fuzzing sessions, test in this order:

1. `form_resource` - Binary format, highest risk
2. `form_file` - Most complex text format
3. `cst_parse` - Core parsing logic
4. `project_file` - Library entry point
5. `class_file` - Common file type
6. `module_file` - Common file type
7. `tokenize` - Foundation layer
8. `sourcestream` - Low-level operations
9. `sourcefile_decode` - Character decoding

## References

- [cargo-fuzz book](https://rust-fuzz.github.io/book/cargo-fuzz.html)
- [libFuzzer documentation](https://llvm.org/docs/LibFuzzer.html)
- [Rust Fuzz Book](https://rust-fuzz.github.io/book/)
- [VB6Parse Repository](https://github.com/scriptandcompile/vb6parse)

## Contributing

Found a bug with fuzzing? Great! Please:

1. **Minimize** the crash input
2. **Create** a regression test from the minimized input
3. **File** an issue with:
   - The minimized crash input (or attach it)
   - Fuzzer target that found it
   - Error message/stack trace
4. **Submit** a PR with the fix and regression test

Crashes found by fuzzing are valuable - they represent real edge cases that could affect users with legacy VB6 projects.
