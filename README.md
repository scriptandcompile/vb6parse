# VB6Parse

A complete, high-performance parser library for Visual Basic 6 code and project files.

[![Crates.io](https://img.shields.io/crates/v/vb6parse.svg)](https://crates.io/crates/vb6parse)
[![Documentation](https://docs.rs/vb6parse/badge.svg)](https://docs.rs/vb6parse)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Overview

VB6Parse is designed as a foundational library for tools that analyze, convert, or process Visual Basic 6 code. While capable of supporting real-time syntax highlighting and language servers, its primary focus is on offline analysis, legacy code utilities, and migration tools.

**Key Features:**
- Fast, efficient parsing with minimal allocations
- Full support for VB6 project files, modules, classes, forms, and resources
- Concrete Syntax Tree (CST) with complete source fidelity
- 160+ built-in VB6 library functions and 42 statements
- Comprehensive error handling with detailed failure information
- Zero-copy tokenization and streaming parsing

## Quick Start

Add VB6Parse to your `Cargo.toml`:

```toml
[dependencies]
vb6parse = "0.5.1"
```

### Parse a VB6 Project File

```rust
use vb6parse::*;

let input = r#"Type=Exe
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#...\stdole2.tlb#OLE Automation
Module=Module1; Module1.bas
Form=Form1.frm
"#;

// Decode source with Windows-1252 encoding (VB6 default)
let source = SourceFile::from_string("Project1.vbp", input);

// Parse the project
let result = ProjectFile::parse(&source);

// Handle results
let (project, failures) = result.unpack();

if let Some(project) = project {
    println!("Project type: {:?}", project.project_type);
    println!("Modules: {}", project.modules().count());
    println!("Forms: {}", project.forms().count());
}

// Print any parsing errors
for failure in failures {
    failure.print();
}
```

### Parse a VB6 Module

```rust
use vb6parse::*;

let code = r#"Attribute VB_Name = "MyModule"
Public Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
"#;

let source = SourceFile::from_string("MyModule.bas", code);
let result = ModuleFile::parse(&source);

let (module, failures) = result.unpack();
if let Some(module) = module {
    println!("Module name: {}", module.attributes.name);
}
```

### Tokenize VB6 Code

```rust
use vb6parse::*;

let source = SourceFile::from_string("test.bas", "Dim x As Integer");
let result = tokenize(&source);

let (token_stream, failures) = result.unpack();
if let Some(tokens) = token_stream {
    for token in tokens.iter() {
        println!("{:?}: {:?}", token.kind(), token.text());
    }
}
```

### Parse to Concrete Syntax Tree

```rust
use vb6parse::*;

let code = "Sub Test()\n    x = 5\nEnd Sub";
let source = SourceFile::from_string("test.bas", code);

// Tokenize first
let (tokens, _) = tokenize(&source).unpack();

// Parse to CST
if let Some(tokens) = tokens {
    let (cst, _) = parse(&tokens).unpack();
    
    if let Some(tree) = cst {
        // Navigate the syntax tree
        println!("Root children: {}", tree.root().child_count());
    }
}
```

### Navigating the CST

The CST provides rich navigation capabilities for traversing and querying the tree structure:

```rust
use vb6parse::*;
use vb6parse::parsers::SyntaxKind;

let source = "Sub Test()\n    Dim x As Integer\n    x = 42\nEnd Sub";
let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
let root = cst.to_serializable().root;

// Basic navigation
let child_count = root.child_count();
let first = root.first_child();

// Find by kind
let sub_stmt = root.find(SyntaxKind::SubStatement);  // First match
let all_dims = root.find_all(SyntaxKind::DimStatement);  // All matches

// Filter children
let non_tokens: Vec<_> = root.non_token_children().collect();
let significant: Vec<_> = root.significant_children().collect();

// Custom search with predicates
let keywords = root.find_all_if(|n| n.kind.to_string().ends_with("Keyword"));
let complex = root.find_all_if(|n| !n.is_token && n.children.len() > 5);

// Iterate all nodes depth-first
for node in root.descendants() {
    if node.is_significant() {
        println!("{:?}: {}", node.kind, node.text);
    }
}

// Convenience checkers
if node.is_comment() || node.is_whitespace() {
    // Skip trivia
}
```

**Available Navigation Methods:**

Both `ConcreteSyntaxTree` and `CstNode` provide:
- **Basic:** `child_count()`, `first_child()`, `last_child()`, `child_at()`
- **By Kind:** `children_by_kind()`, `first_child_by_kind()`, `contains_kind()`
- **Recursive:** `find()`, `find_all()`
- **Filtering:** `non_token_children()`, `token_children()`, `significant_children()`
- **Predicates:** `find_if()`, `find_all_if()`
- **Traversal:** `descendants()`, `depth_first_iter()`

CstNode also provides: `is_whitespace()`, `is_newline()`, `is_comment()`, `is_trivia()`, `is_significant()`

**See also:** [examples/cst_navigation.rs](examples/cst_navigation.rs) for comprehensive examples.

## API Surface

### Top-Level Imports

For common use cases, import everything with:

```rust
use vb6parse::*;
```

This brings in:
- **I/O Layer:** `SourceFile`, `SourceStream`
- **Lexer:** `tokenize()`, `Token`, `TokenStream`
- **File Parsers:** `ProjectFile`, `ClassFile`, `ModuleFile`, `FormFile`, `FormResourceFile`
- **Syntax Parsers:** `parse()`, `ConcreteSyntaxTree`, `SyntaxKind`, `SerializableTree`
- **Error Handling:** `ErrorDetails`, `ParseResult`, all error kind enums

### Layer Modules (Advanced Usage)

For advanced use cases, access specific layers:

```rust
use vb6parse::io::{SourceFile, SourceStream, Comparator};
use vb6parse::lexer::{tokenize, Token, TokenStream};
use vb6parse::parsers::{parse, ConcreteSyntaxTree};
use vb6parse::language::controls::{Control, ControlKind};
use vb6parse::errors::{ProjectErrorKind, FormErrorKind};
```

### Parsing Architecture

```
Bytes/String/File → SourceFile → SourceStream → TokenStream → CST → Object Layer
                    (Windows-1252) (Characters)   (Tokens)    (Tree) (Structured)
```

**Layers:**

1. **I/O Layer** (`io`): Character decoding and stream access
2. **Lexer Layer** (`lexer`): Tokenization with keyword lookup
3. **Syntax Layer** (`syntax`): VB6 language constructs and library functions
4. **Parsers Layer** (`parsers`): CST construction from tokens
5. **Files Layer** (`files`): High-level file format parsers
6. **Language Layer** (`language`): VB6 types, colors, controls
7. **Errors Layer** (`errors`): Comprehensive error types

## Source Code Organization

```
src/
├── io/                          # I/O Layer - Character streams and decoding
│   ├── mod.rs                   # SourceFile, SourceStream
│   ├── comparator.rs            # Case-sensitive/insensitive comparison
│   └── decode.rs                # Windows-1252 decoding
│
├── lexer/                       # Lexer Layer - Tokenization
│   ├── mod.rs                   # tokenize() function, keyword lookup
│   └── token_stream.rs          # TokenStream implementation
│
├── syntax/                      # Syntax Layer - VB6 Language constructs
│   ├── library/                 # VB6 built-in library
│   │   ├── functions/           # 160+ VB6 functions (14 categories)
│   │   │   ├── array/           # Array, Filter, Join, Split, etc.
│   │   │   ├── conversion/      # CBool, CInt, CLng, Str, Val, etc.
│   │   │   ├── datetime/        # Date, Now, Time, Year, Month, etc.
│   │   │   ├── file_system/     # Dir, EOF, FileLen, LOF, etc.
│   │   │   ├── financial/       # FV, IPmt, IRR, NPV, PV, Rate, etc.
│   │   │   ├── interaction/     # MsgBox, InputBox, Shell, etc.
│   │   │   ├── math/            # Abs, Cos, Sin, Tan, Log, Sqr, etc.
│   │   │   ├── miscellaneous/   # Environ, RGB, QBColor, etc.
│   │   │   ├── string/          # Left, Right, Mid, Len, Trim, etc.
│   │   │   └── ...
│   │   └── statements/          # 42 VB6 statements (9 categories)
│   │       ├── control_flow/    # If, Select Case, For, While, etc.
│   │       ├── declarations/    # Dim, ReDim, Const, Enum, etc.
│   │       ├── error_handling/  # On Error, Resume, Err, etc.
│   │       ├── file_operations/ # Open, Close, Get, Put, etc.
│   │       ├── objects/         # Set, With, RaiseEvent, etc.
│   │       └── ...
│   └── expressions/             # Expression parsing utilities
│
├── parsers/                     # Parsers Layer - CST construction
│   ├── cst/                     # Concrete Syntax Tree implementation
│   │   ├── mod.rs               # parse(), ConcreteSyntaxTree, CstNode
│   │   └── rowan_wrapper.rs     # Red-green tree wrapper
│   ├── parseresults.rs          # ParseResult<T, E> type
│   └── syntaxkind.rs            # SyntaxKind enum (all token types)
│
├── files/                       # Files Layer - VB6 file format parsers
│   ├── common/                  # Shared parsing utilities
│   │   ├── properties.rs        # Property bag, PropertyGroup
│   │   ├── attributes.rs        # Attribute statement parsing
│   │   └── references.rs        # Object reference parsing
│   ├── project/                 # VBP - Project files
│   │   ├── mod.rs               # ProjectFile struct and parser
│   │   ├── properties.rs        # Project properties
│   │   ├── references.rs        # Reference types
│   │   └── compilesettings.rs   # Compilation settings
│   ├── class/                   # CLS - Class modules
│   ├── module/                  # BAS - Code modules
│   ├── form/                    # FRM - Forms
│   └── resource/                # FRX - Form resources
│
├── language/                    # Language Layer - VB6 types and definitions
│   ├── color.rs                 # VB6 color constants and Color type
│   ├── controls/                # VB6 control definitions (50+ controls)
│   │   ├── mod.rs               # Control, ControlKind enums
│   │   ├── form.rs              # FormProperties
│   │   ├── textbox.rs           # TextBoxProperties
│   │   ├── label.rs             # LabelProperties
│   │   └── ...                  # 50+ control types
│   └── tokens.rs                # Token enum definition
│
├── errors/                      # Errors Layer - Error types
│   ├── mod.rs                   # ErrorDetails, error printing
│   ├── decode.rs                # SourceFileErrorKind
│   ├── tokenize.rs              # CodeErrorKind
│   ├── project.rs               # ProjectErrorKind
│   ├── class.rs                 # ClassErrorKind
│   ├── module.rs                # ModuleErrorKind
│   ├── form.rs                  # FormErrorKind
│   ├── property.rs              # PropertyError
│   └── resource.rs              # ResourceErrorKind
│
└── lib.rs                       # Public API surface
```

## Common Tasks

### 1. Load and Validate a VB6 Project

```rust
use vb6parse::*;
use std::fs;

fn load_project(path: &str) -> Result<ProjectFile, String> {
    let bytes = fs::read(path).map_err(|e| e.to_string())?;
    
    let source = SourceFile::decode_with_replacement(path, &bytes)
        .map_err(|e| format!("Decode error: {:?}", e))?;
    
    let result = ProjectFile::parse(&source);
    let (project, failures) = result.unpack();
    
    if !failures.is_empty() {
        for failure in &failures {
            failure.print();
        }
        return Err(format!("Parse had {} failures", failures.len()));
    }
    
    project.ok_or_else(|| "Parse returned no result".to_string())
}
```

### 2. Extract All Form Controls

```rust
use vb6parse::*;
use vb6parse::language::controls::{Control, ControlKind};

fn extract_controls(form_path: &str) -> Vec<String> {
    let source = SourceFile::from_file(form_path).unwrap();
    let result = FormFile::parse(&source);
    let (form, _) = result.unpack();
    
    let mut control_names = Vec::new();
    
    if let Some(form) = form {
        fn visit_control(control: &Control, names: &mut Vec<String>) {
            names.push(control.name.clone());
            
            // Recursively visit children
            if let Some(children) = control.kind.children() {
                for child in children {
                    visit_control(child, names);
                }
            }
        }
        
        for control in &form.controls {
            visit_control(control, &mut control_names);
        }
    }
    
    control_names
}
```

### 3. Analyze Code Without Full Parsing

```rust
use vb6parse::*;

fn count_identifiers(code: &str, function_name: &str) -> usize {
    let source = SourceFile::from_string("temp.bas", code);
    let result = tokenize(&source);
    let (tokens, _) = result.unpack();
    
    tokens
        .map(|ts| {
            ts.iter()
                .filter(|t| {
                    matches!(t.kind(), SyntaxKind::IdentifierToken)
                        && t.text().eq_ignore_ascii_case(function_name)
                })
                .count()
        })
        .unwrap_or(0)
}
```

### 4. Convert VB6 to Syntax Tree and Back

```rust
use vb6parse::*;

fn roundtrip_code(code: &str) -> String {
    let source = SourceFile::from_string("temp.bas", code);
    let (tokens, _) = tokenize(&source).unpack();
    
    if let Some(tokens) = tokens {
        let (cst, _) = parse(&tokens).unpack();
        
        if let Some(tree) = cst {
            // CST preserves all whitespace and comments
            return tree.root().text().to_string();
        }
    }
    
    String::new()
}
```

## Advanced Topics

### Error Handling

VB6Parse uses a custom `ParseResult<T, E>` type that separates successful results from recoverable errors:

```rust
use vb6parse::*;

let result = ProjectFile::parse(&source);

// Option 1: Unpack into result and failures
let (project_opt, failures) = result.unpack();

// Option 2: Check for failures first
if result.has_failures() {
    for failure in result.failures() {
        eprintln!("Error at line {}: {:?}", 
                  failure.offset.line_number,
                  failure.kind);
    }
}

// Option 3: Convert to Result<T, Vec<ErrorDetails>>
let std_result = result.ok_or_errors();
```

**See also:**
- [src/parsers/parseresults.rs](src/parsers/parseresults.rs) - ParseResult implementation
- [src/errors/mod.rs](src/errors/mod.rs) - Error types and ErrorDetails

### Working with the CST

The Concrete Syntax Tree preserves all source information including whitespace and comments:

```rust
use vb6parse::*;

let (tree, _) = parse(&tokens).unpack().0.unwrap();

// Navigate the tree
let root = tree.root();
for child in root.children() {
    println!("Node: {:?}", child.kind());
    println!("Text: {}", child.text());
}

// Serialize for debugging
let serializable = tree.to_serializable();
println!("{:#?}", serializable);
```

**See also:**
- [src/parsers/cst/mod.rs](src/parsers/cst/mod.rs) - CST documentation
- [examples/cst_parse.rs](examples/cst_parse.rs) - CST parsing example
- [examples/debug_cst.rs](examples/debug_cst.rs) - CST debugging

### Character Encoding

VB6 uses Windows-1252 encoding. Always use `decode_with_replacement()` for file content:

```rust
use vb6parse::*;

// From bytes (e.g., file read)
let bytes = std::fs::read("file.bas")?;
let source = SourceFile::decode_with_replacement("file.bas", &bytes)?;

// From UTF-8 string (testing/programmatic)
let source = SourceFile::from_string("test.bas", "Dim x As Integer");
```

**See also:**
- [src/io/decode.rs](src/io/decode.rs) - Decoding implementation
- [examples/parse_class.rs](examples/parse_class.rs) - Byte-level parsing

### VB6 Library Functions

VB6Parse includes full definitions for 160+ VB6 library functions organized into 14 categories:

```rust
// Access function metadata
use vb6parse::syntax::library::functions::string::left;
use vb6parse::syntax::library::functions::math::sin;
use vb6parse::syntax::library::functions::conversion::cint;

// Each module includes:
// - Full VB6 documentation
// - Function signatures
// - Parameter descriptions
// - Usage examples
// - Related functions
```

**Categories:**
- Array manipulation (Array, Filter, Join, Split, UBound, LBound)
- Conversion (CBool, CDate, CInt, CLng, CStr, Val, Str)
- Date/Time (Date, Time, Now, Year, Month, Day, Hour, DateAdd, DateDiff)
- File System (Dir, EOF, FileLen, FreeFile, LOF, Seek)
- Financial (FV, IPmt, IRR, NPV, PV, Rate)
- Formatting (Format, FormatCurrency, FormatDateTime, FormatNumber, FormatPercent)
- Interaction (MsgBox, InputBox, Shell, CreateObject, GetObject)
- Inspection (IsArray, IsDate, IsEmpty, IsNull, IsNumeric, TypeName, VarType)
- Math (Abs, Atn, Cos, Exp, Log, Rnd, Sgn, Sin, Sqr, Tan)
- String (Left, Right, Mid, Len, InStr, Replace, Trim, UCase, LCase)
- And more...

**See also:** [src/syntax/library/functions/](src/syntax/library/functions/)

### Form Resources (FRX Files)

Form resource files contain binary data for controls (images, icons, property blobs):

```rust
use vb6parse::*;

let bytes = std::fs::read("Form1.frx")?;
let result = FormResource::load_from_bytes(&bytes);

let (resource, failures) = result.unpack();
if let Some(resource) = resource {
    for (offset, data) in &resource.resources {
        println!("Resource at offset {}: {} bytes", offset, data.len());
    }
}
```

**See also:**
- [documents/FRX_format.md](documents/FRX_format.md) - FRX format specification
- [examples/debug_resource.rs](examples/debug_resource.rs) - Resource file debugging

## Testing

VB6Parse has comprehensive test coverage:

- **5,467 library tests** - Testing VB6 library functions and statements
- **83 documentation tests** - Ensuring examples work correctly
- **31 integration tests** - Parsing real-world VB6 projects

### Running Tests

```bash
# Clone test data (required for integration tests)
git submodule update --init --recursive

# Run all tests
cargo test

# Run only library tests
cargo test --lib

# Run only integration tests
cargo test --test '*'

# Run documentation tests
cargo test --doc
```

### Snapshot Testing

Integration tests use [insta](https://docs.rs/insta) for snapshot testing:

```bash
# Review snapshot changes
cargo insta review

# Accept all snapshots
cargo insta accept
```

**Test data location:** `tests/data/` (git submodules of real VB6 projects)

**See also:**
- [tests/](tests/) - Test files
- [tests/snapshots/](tests/snapshots/) - Snapshot files

## Benchmarking

VB6Parse includes criterion benchmarks for performance testing:

```bash
# Run all benchmarks
cargo bench

# Run specific benchmark
cargo bench bulk_parser_load

# Generate HTML reports
# Results saved to target/criterion/
```

**Benchmarks:**
- `bulk_parser_load` - Parsing multiple large VB6 projects
- Token stream generation
- CST construction

**See also:** [benches/](benches/)

## Code Coverage

VB6Parse uses `cargo-llvm-cov` to track test coverage and ensure comprehensive testing across all modules.

### Installation

```bash
# Install cargo-llvm-cov
cargo install cargo-llvm-cov
```

### Running Coverage

```bash
# Generate coverage report (terminal output)
cargo llvm-cov

# Generate HTML report
cargo llvm-cov --html
# Open target/llvm-cov/html/index.html in your browser

# Generate coverage with open HTML report
cargo llvm-cov --open

# Generate detailed coverage for specific packages
cargo llvm-cov --package vb6parse

# Include tests in coverage
cargo llvm-cov --all-targets

# Generate LCOV format (for CI/CD integration)
cargo llvm-cov --lcov --output-path lcov.info
```

### Coverage Reports

Coverage reports are saved to:
- **HTML reports:** `target/llvm-cov/html/`
- **Terminal summary:** Displays percentage coverage after running `cargo llvm-cov`
- **LCOV files:** `lcov.info` (when using `--lcov` flag)

**Current Coverage:**
- **Library tests:** 5,467 tests covering VB6 library functions
- **Integration tests:** 31 tests with real-world VB6 projects
- **Documentation tests:** 83 tests ensuring examples work
- **Coverage focus:** Parsers, tokenization, error handling, and file format support

## Contributing to VB6Parse

Contributions are welcome! Please see the [CONTRIBUTING.md](CONTRIBUTING.md) file for more information.

### Development Setup

```bash
# Clone repository
git clone https://github.com/scriptandcompile/vb6parse
cd vb6parse

# Get test data
git submodule update --init --recursive

# Run tests
cargo test

# Run benchmarks
cargo bench

# Check for issues
cargo clippy

# Format code
cargo fmt
```

### Code Organization Guidelines

1. **Layer Separation:** Keep clear boundaries between layers
2. **Windows-1252 Handling:** Always use `SourceFile::decode_with_replacement()`
3. **Error Recovery:** Parsers should recover from errors when possible
4. **CST Fidelity:** Preserve all source text including whitespace and comments
5. **Documentation:** Include doc tests for public APIs

### Adding New Features

**VB6 Library Functions:**
- Add to appropriate category in `src/syntax/library/functions/`
- Include full VB6 documentation
- Add comprehensive tests
- Update category mod.rs

**Control Types:**
- Add to `src/language/controls/`
- Define properties struct
- Add to ControlKind enum
- Include property validation

**Error Types:**
- Add to appropriate error module in `src/errors/`
- Ensure Display implementation
- Add context information

### Performance Considerations

- Use zero-copy where possible (string slices, not String)
- Avoid unnecessary allocations (use iterators)
- Leverage rowan's red-green tree for CST memory efficiency
- Use `phf` crate for compile-time lookup tables

**See also:**
- [CHANGELOG.md](CHANGELOG.md) - Version history

## Supported File Types

| Extension | Description | Status |
|-----------|-------------|--------|
| `.vbp` | Project files | ✅ Complete |
| `.cls` | Class modules | ✅ Complete |
| `.bas` | Code modules | ✅ Complete |
| `.frm` | Forms | ⚠️ PArtial (font, some icons, etc) |
| `.frx` | Form resources | ⚠️ Partial (binary blobs loaded, not all mapped to properties) |
| `.ctl` | User controls | ✅ Parsed as forms |
| `.dob` | User documents | ✅ Parsed as forms |
| `.vbw` | IDE window state | ❌ Not yet implemented |
| `.dsx` | Data environments | ❌ Not yet implemented |
| `.dsr` | Data env. resources | ❌ Not yet implemented |
| `.ttx` | Crystal reports | ❌ Not yet implemented |

## Project Status

- ✅ **Core Parsing:** Fully implemented for VBP, CLS, BAS files
- ✅ **Tokenization:** Complete with keyword lookup
- ✅ **CST Construction:** Full syntax tree with source fidelity
- ✅ **Error Handling:** Comprehensive error types and recovery
- ✅ **VB6 Library:** 160+ functions, 42 statements documented
- ⚠️ **FRX Resources:** Binary loading complete, property mapping partial
- ⚠️ **FRM Properties:** Majority of FRM properties load properly, (icon, background, font mapping partial)
- ❌ **AST:** Not yet implemented (CST available)
- ✅ **Testing:** 5,500+ tests across unit, integration, and doc tests
- ✅ **Benchmarking:** Criterion-based performance testing
- ✅ **Fuzz Testing:** Coverage-guided fuzzing with cargo-fuzz
- ✅ **Documentation:** Comprehensive API docs and examples

## Fuzz Testing

VB6Parse includes comprehensive fuzz testing using `cargo-fuzz` and libFuzzer to discover edge cases, crashes, and undefined behavior.

**Available Fuzz Targets:**
- `sourcefile_decode` - Tests Windows-1252 decoding with arbitrary bytes
- `sourcestream` - Tests low-level character stream operations
- `tokenize` - Tests tokenization with malformed VB6 code
- `cst_parse` - Tests Concrete Syntax Tree parsing with invalid syntax

**Quick Start:**

```bash
# Install cargo-fuzz (requires nightly)
cargo install cargo-fuzz

# Run a fuzzer for 60 seconds
cargo +nightly fuzz run sourcefile_decode -- -max_total_time=60

# List all fuzz targets
cargo +nightly fuzz list
```

**Learn More:** See [fuzz/README.md](fuzz/README.md) for detailed usage and [Fuzzing.md](Fuzzing.md) for the complete fuzzing strategy.

## Examples

All examples are located in the [examples/](examples/) directory:

| Example | Description |
|---------|-------------|
| [audiostation_parse.rs](examples/audiostation_parse.rs) | Parse a complete real-world VB6 project |
| [cst_parse.rs](examples/cst_parse.rs) | Parse tokens directly to CST |
| [debug_cst.rs](examples/debug_cst.rs) | Display CST debug representation |
| [debug_resource.rs](examples/debug_resource.rs) | Inspect FRX resource files |
| [parse_class.rs](examples/parse_class.rs) | Parse class files from bytes |
| [parse_control_only.rs](examples/parse_control_only.rs) | Parse individual form controls |
| [parse_form.rs](examples/parse_form.rs) | Parse VB6 forms |
| [parse_module.rs](examples/parse_module.rs) | Parse code modules |
| [parse_project.rs](examples/parse_project.rs) | Parse project files |
| [sourcestream.rs](examples/sourcestream.rs) | Work with character streams |
| [tokenstream.rs](examples/tokenstream.rs) | Tokenize VB6 code |

Run any example with:

```bash
cargo run --example parse_project
```

## Resources

- **Documentation:** [docs.rs/vb6parse](https://docs.rs/vb6parse)
- **Repository:** [github.com/scriptandcompile/vb6parse](https://github.com/scriptandcompile/vb6parse)
- **Crates.io:** [crates.io/crates/vb6parse](https://crates.io/crates/vb6parse)
- **License:** MIT

## Limitations

1. **Encoding:** Primarily designed for "predominantly English" source code with Windows-1252 encoding detection limitations
2. **AST:** Abstract Syntax Tree is not yet implemented (Concrete Syntax Tree is available)
3. **FRX Mapping:** Binary resources are loaded but not all are mapped to control properties
4. **Real-time Use:** While capable, not optimized for real-time highlighting or LSP (focus is on offline analysis)

## License

MIT License - See [LICENSE](LICENSE) file for details.

---

Built with ❤️ by [ScriptAndCompile](https://github.com/scriptandcompile)
