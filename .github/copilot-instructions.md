# VB6Parse Copilot Instructions

## Project Overview
VB6Parse is a Rust library for parsing Visual Basic 6 code and project files. It's designed for offline analysis, legacy utilities, and code conversion tools (not real-time highlighting or LSP). The parser handles multiple VB6 file types: `.vbp` (projects), `.cls` (classes), `.bas` (modules), `.frm` (forms), `.frx` (form resources).

## Architecture: Three-Layer Parsing Pipeline

```
Bytes/String/File → SourceFile → SourceStream → TokenStream → CST → Object Layer
```

1. **SourceFile**: Handles Windows-1252 encoding (VB6 era). Always use `SourceFile::decode_with_replacement(filename, bytes)` - never raw strings
2. **SourceStream**: Low-level character stream with offset tracking and line/column info
3. **TokenStream**: Tokenized via `tokenize()` using keyword lookup tables (see [src/tokenize.rs](src/tokenize.rs#L8))
4. **CST (Concrete Syntax Tree)**: Full syntax tree including whitespace/comments, wraps rowan library (see [src/parsers/cst/mod.rs](src/parsers/cst/mod.rs))
5. **Object Layer**: High-level structs like `ProjectFile`, `ClassFile`, `ModuleFile`, `FormFile`

## Error Handling Pattern

All parsers return `ParseResult<'a, T, E>` which contains both:
- `result: Option<T>` - The parsed output
- `failures: Vec<ErrorDetails<'a, E>>` - Non-fatal parsing errors

Always check for failures after parsing by unpacking and handling the results.


Example:
```rust
let result = ProjectFile::parse(&source_file);

let (project, failures) = result.unpack();

```

## Testing Conventions

- **Snapshot testing**: Uses `insta` for all integration tests. Run `cargo insta test` and `cargo insta review` to update snapshots
- **Test data**: Lives in `tests/data/` and includes git submodules of real VB6 projects. Run `git submodule update --init --recursive` before testing
- **Benchmarking**: Uses criterion. Benchmarks in `benches/` use `include_bytes!()` to embed test files at compile time
- **Pattern**: Tests call `SourceFile::decode_with_replacement()` → `TypeFile::parse()` → `insta::assert_yaml_snapshot!()`
- **unwrap()**: Tests often call `unpack()` on `ParseResult` to separate parsed output and failures for snapshotting. Never use `unwrap()` on `ParseResult` directly in production code or tests.

## Key File Types & Entry Points

- **Projects** (`*.vbp`): `ProjectFile::parse(&source_file)` - Lists references, modules, forms, etc. without loading them
- **Classes** (`*.cls`): `ClassFile::parse(&source_file)` - Returns header + CST of code
- **Modules** (`*.bas`): `ModuleFile::parse(&source_file)` - Returns header + CST of code  
- **Forms** (`*.frm`): `FormFile::parse(&source_file)` - Special: UI controls in header, code in body. Forms have resource files (`.frx`)
- **Form Resources** (`*.frx`): `FormResource::load_from_bytes(&bytes)` - Parses binary blobs for control strings and binary data.

## CST Navigation

The CST preserves all tokens (whitespace, comments). Navigate via `CstNode`:
- `child_count()`, `children()` - Iterate child nodes
- `text()` - Get source text span
- `kind()` - Get `SyntaxKind` enum value
- Internal: Uses rowan's red-green tree for memory efficiency

## Common Patterns

1. **Case-insensitive parsing**: VB6 is case-insensitive. Use `Comparator::CaseInsensitive` when calling `SourceStream` methods
2. **Keyword lookup**: Keywords use a static `phf_ordered_map` for fast lookups ([src/tokenize.rs](src/tokenize.rs#L8))
3. **Property enums**: VB6 properties are modeled as Rust enums (see [src/language/controls/](src/language/controls/) for 50+ property types like `Alignment`, `BorderStyle`)
4. **Header parsing**: `.cls`, `.bas`, `.frm` files start with `VERSION` line and `Attribute` statements before code
5. **Token vs Text**: Prefer reading from tokens in `TokenStream` for parsing logic. Use `CstNode::text()` only when exact source text is needed (e.g., for error messages or snapshots). Do not mix raw string operations with token-based parsing.

## Build & Run

- `cargo test` - Run tests (requires submodules)
- `cargo bench` - Run criterion benchmarks
- `cargo doc --open` - Generate and view docs
- No special build flags or features currently

## Things to Avoid

- Don't expose rowan types in public APIs (CST keeps them internal)
- Don't assume UTF-8 - always use `SourceFile::decode_with_replacement()` for Windows-1252
- Don't skip `has_failures()` checks - parsers can partially succeed with errors
- Don't mutate `SourceStream` offset without accounting for `forward()` semantics

## Current Limitations

- Form resource (`.frx`) loading doesn't fully map binary blobs to all control properties yet
- CST exists but AST (Abstract Syntax Tree) is not yet implemented
- Focus is on "predominantly english" source code due to encoding detection limitations
