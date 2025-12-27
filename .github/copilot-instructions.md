# VB6Parse Copilot Instructions

## Project Overview
VB6Parse is a Rust library (v0.5.1) for parsing Visual Basic 6 code and project files. It's designed for offline analysis, legacy utilities, and code conversion tools (not real-time highlighting or LSP). The parser handles multiple VB6 file types: `.vbp` (projects), `.cls` (classes), `.bas` (modules), `.frm` (forms), `.frx` (form resources), `.ctl` (user controls), `.dob` (user documents).

## Architecture: Seven-Layer Parsing Pipeline

```
Bytes/String/File → SourceFile → SourceStream → TokenStream → CST → Object Layer
                    (Windows-1252) (Characters)   (Tokens)    (Tree) (Structured)
```

**Layers:**
1. **I/O Layer** (`io/`): Character decoding and stream access (SourceFile, SourceStream)
2. **Lexer Layer** (`lexer/`): Tokenization with keyword lookup tables (phf-based), creates TokenStream
3. **Syntax Layer** (`syntax/`): VB6 language constructs - 160+ library functions (14 categories) and 42 statements (9 categories)
4. **Parsers Layer** (`parsers/`): CST construction from tokens, wraps rowan's red-green tree
5. **Files Layer** (`files/`): High-level file format parsers (ProjectFile, ClassFile, ModuleFile, FormFile, FormResource)
6. **Language Layer** (`language/`): VB6 types, colors (24 predefined), controls (50+ types)
7. **Errors Layer** (`errors/`): Comprehensive error types for each layer

## Error Handling Pattern

All parsers return `ParseResult<'a, T, E>` which contains both:
- `result: Option<T>` - The parsed output
- `failures: Vec<ErrorDetails<'a, E>>` - Non-fatal parsing errors

**Fields are now private** - use accessor methods:
- `result()` or `into_result()` - Get the parsed output
- `failures()` or `into_failures()` - Get error list
- `unpack()` - Consume and get `(Option<T>, Vec<ErrorDetails>)`
- `has_failures()` - Check if any errors occurred
- `ok_or_errors()` - Convert to `Result<T, Vec<ErrorDetails>>`

Always check for failures after parsing by unpacking and handling the results.

Example:
```rust
let result = ProjectFile::parse(&source_file);
let (project, failures) = result.unpack();

// Check both result and failures
if let Some(project) = project {
    println!("Parsed successfully");
}
for failure in failures {
    failure.print();  // Print error with context
}
```

## Testing Conventions

- **Snapshot testing**: Uses `insta` for all integration tests. Run `cargo insta test` and `cargo insta review` to update snapshots
- **Test data**: Lives in `tests/data/` and includes git submodules of real VB6 projects. Run `git submodule update --init --recursive` before testing
- **Test coverage**: 5,467 library tests + 83 documentation tests + 31 integration tests
- **Benchmarking**: Uses criterion. Benchmarks in `benches/` use `include_bytes!()` to embed test files at compile time
- **Fuzz testing**: Comprehensive coverage-guided fuzzing with cargo-fuzz and libFuzzer. 9 fuzz targets covering all layers
- **Pattern**: Tests call `SourceFile::decode_with_replacement()` → `TypeFile::parse()` → `insta::assert_yaml_snapshot!()`
- **unwrap()**: Tests often call `unpack()` on `ParseResult` to separate parsed output and failures for snapshotting. Never use `unwrap()` on `ParseResult` directly in production code or tests.

## Key File Types & Entry Points

- **Projects** (`*.vbp`): `ProjectFile::parse(&source_file)` - Lists references, modules, forms, etc. without loading them
- **Classes** (`*.cls`): `ClassFile::parse(&source_file)` - Returns header + CST of code
- **Modules** (`*.bas`): `ModuleFile::parse(&source_file)` - Returns header + CST of code  
- **Forms** (`*.frm`): `FormFile::parse(&source_file)` - Special: UI controls in header, code in body. Forms have resource files (`.frx`)
- **Form Resources** (`*.frx`): `FormResourceFile::load_from_bytes(&bytes)` - Parses binary blobs for control strings and binary data. Supports multiple FRX header formats (4-byte, 8-byte, 12-byte)
- **User Controls** (`*.ctl`): Parsed as forms with `FormFile::parse()`
- **User Documents** (`*.dob`): Parsed as forms with `FormFile::parse()`

## CST Navigation

The CST preserves all tokens (whitespace, comments). Navigate via `CstNode`:
- `child_count()`, `children()` - Iterate child nodes
- `text()` - Get source text span
- `kind()` - Get `SyntaxKind` enum value
- Internal: Uses rowan's red-green tree for memory efficiency

## Common Patterns

1. **Case-insensitive parsing**: VB6 is case-insensitive. Use `Comparator::CaseInsensitive` when calling `SourceStream` methods
2. **Keyword lookup**: Keywords use a static `phf_ordered_map` for fast lookups (see [src/lexer/mod.rs](src/lexer/mod.rs))
3. **Property enums**: VB6 properties are modeled as Rust enums (see [src/language/controls/](src/language/controls/) for 50+ property types like `Alignment`, `BorderStyle`)
4. **Header parsing**: `.cls`, `.bas`, `.frm` files start with `VERSION` line and `Attribute` statements before code
5. **Token vs Text**: Prefer reading from tokens in `TokenStream` for parsing logic. Use `CstNode::text()` only when exact source text is needed (e.g., for error messages or snapshots). Do not mix raw string operations with token-based parsing.
6. **VB6 Library**: 160+ built-in functions organized in 14 categories (array, conversion, datetime, file_system, financial, formatting, interaction, inspection, math, string, etc.). 42 statements in 9 categories (control_flow, declarations, error_handling, file_operations, objects, etc.).

## Build & Run

- `cargo test` - Run tests (requires submodules: `git submodule update --init --recursive`)
- `cargo bench` - Run criterion benchmarks
- `cargo doc --open` - Generate and view docs
- `cargo +nightly fuzz run <target>` - Run fuzz tests (requires `cargo install cargo-fuzz`)
- `cargo clippy` - Check for linting issues
- `cargo insta review` - Review snapshot test changes
- No special build flags or features currently

## Things to Avoid

- Don't expose rowan types in public APIs (CST keeps them internal)
- Don't assume UTF-8 - always use `SourceFile::decode_with_replacement()` for Windows-1252
- Don't skip `has_failures()` checks - parsers can partially succeed with errors
- Don't mutate `SourceStream` offset without accounting for `forward()` semantics
- Don't use `unwrap()` on `ParseResult` - use `unpack()` instead
- Don't batch multiple state changes to `ParseResult` - mark todos completed immediately
- Don't use get_ prefix for parameterless getters (C-GETTER violation)

## Current Limitations

- Form resource (`.frx`) loading doesn't fully map binary blobs to all control properties yet
- CST exists but AST (Abstract Syntax Tree) is not yet implemented
- Focus is on "predominantly english" source code due to encoding detection limitations
- Not optimized for real-time highlighting or LSP (focus is on offline analysis)
- Some file types not yet implemented: `.vbw` (IDE window state), `.dsx` (data environments), `.dsr` (data env. resources), `.ttx` (Crystal reports)

## Recent Changes (v0.5.1 - Unreleased)

### Changed
- Removed `winnow` dependency - no longer used in the codebase
- Renamed `source_file` field to `file_name` in `TokenStream`
- Made `TokenStream` fields private with accessor methods (`file_name()`, `offset()`)
- All C-GETTER violations fixed - removed `get_` prefix from parameterless getters
- ProjectFile fields are now mostly private (except `project_type`, `other_properties`, `properties`)
- ParseResult fields are now private with accessor methods

### Key Features (v0.5.0 - v0.5.1)
- Full FRX (form resource) file support with multiple header formats (4-byte, 8-byte, 12-byte)
- Support for 40+ VB6 statements (RmDir, Resume, Randomize, RaiseEvent, Put, Print, Open, Mid, Lock, Load, Let, Kill, Input, Get, FileCopy, Event, Error, Erase, Enum, DeleteSettings, Declare, Date, Close, Property, Exit, GoTo, Select Case, With, Set, For Each, Do...Loop, Call, and more)
- Support for Public/Private variable declarations including 'WithEvents' keyword
- Improved line continuation support and conditional parsing
- Enhanced test coverage with 5,467 library tests + 83 doc tests + 31 integration tests
- Comprehensive fuzzing with 9 fuzz targets
