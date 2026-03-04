# VB6Parse Enhancement Opportunities

This document outlines potential enhancements for the vb6parse library, organized by category and priority. These suggestions are based on a comprehensive scan of the repository as of February 23, 2026.

## Table of Contents

- [Outstanding TODOs and Code Issues](#outstanding-todos-and-code-issues)
- [Testing and Quality Assurance](#testing-and-quality-assurance)
- [Documentation Enhancements](#documentation-enhancements)
- [Performance Optimizations](#performance-optimizations)
- [API and Functionality Enhancements](#api-and-functionality-enhancements)
- [Error Handling Improvements](#error-handling-improvements)
- [Tooling and DevOps](#tooling-and-devops)
- [Code Organization and Maintenance](#code-organization-and-maintenance)
- [Security Considerations](#security-considerations)

---

## Outstanding TODOs and Code Issues

### High Priority

#### 1. Form Resource (FRX) Property Parsing
**File:** Multiple control files in `src/language/controls/`

**Issue:** Several control types have incomplete property parsing from binary FRX data:
- `DragIcon` parsing (drivelistbox.rs, dirlistbox.rs, data.rs, label.rs, frame.rs, scrollbars.rs, textbox.rs, optionbutton.rs)
- `MouseIcon` parsing (label.rs, frame.rs, picturebox.rs, scrollbars.rs, textbox.rs, optionbutton.rs, form.rs)
- `Picture` parsing (picturebox.rs, optionbutton.rs, form.rs)
- `DisabledPicture` and `DownPicture` parsing (optionbutton.rs)
- `Font` parsing (form.rs)
- `Icon` parsing (form.rs)
- `Palette` parsing (form.rs)

**Suggested Fix:**
- Review `documents/FRX_format.md` for binary blob format specifications
- Implement parsers for icon/picture binary data structures
- Add comprehensive tests for each property type
- Consider creating a shared `ImageProperty` parser module

#### 2. String Parsing Issues
**Files:** 
- `src/syntax/library/statements/filesystem/setattr.rs:690`
- `src/syntax/library/statements/string_manipulation/rset.rs:563`

**Issues:**
```rust
// TODO: need to fix so it captures strings correctly. I completely forgot that
// TODO: This one is definitely incorrect. It looks like it's getting borked up with 'text' and the 'Text' Keyword.
```

**Suggested Fix:**
- Review tokenizer behavior with string literals vs Text keyword
- Add specific test cases for edge cases (strings containing keywords)
- Consider context-aware tokenization for distinguishing identifiers from keywords

#### 3. Array Function Parsing Issues
**File:** `src/syntax/library/functions/arrays/array.rs`

**Issues:**
```rust
// TODO: Fix failure to get NothingKeyword, instead, getting an Identifier ("Nothing")
// TODO: Looks like the CallStatement doesn't correctly have a CallExpression internally here.
// TODO: Inside the case clause we should be parsing a CallExpression for Array(...)
// TODO: It looks like the IdentifierExpression and the PeriodOperator don't correctly parse here.
```

**Suggested Fix:**
- Review keyword recognition in expression contexts
- Ensure `Nothing` keyword is properly tokenized in all contexts
- Fix CallStatement to properly contain CallExpression
- Add test cases for Array() in various contexts (case statements, member access)

#### 4. Object Statement Whitespace Handling
**File:** `src/parsers/cst/mod.rs:1990`

```rust
// TODO: Change this parsing to better handle leading whitespace on object statements.
```

**Suggested Fix:**
- Review whitespace handling in object statement parser
- Ensure consistent whitespace behavior across statement types
- Add regression tests for various whitespace patterns

### Medium Priority

#### 5. Label and GoTo Statement Improvements
**File:** `src/syntax/statements/control_flow/jump.rs:165-166`

```rust
// TODO: Consider adding a list of keywords that can be used as labels.
// TODO: Also consider modifying tokenizer to recognize when inside header to more easily identify Identifiers vs header only keywords.
```

**Suggested Fix:**
- Create a whitelist of keywords valid as labels
- Implement context-aware tokenization for procedure headers
- Document which keywords can/cannot be used as labels

#### 6. DefType Statement Validation
**File:** `src/parsers/cst/deftype_statements.rs:67,73`

```rust
// TODO: Validate that the keyword is one of the valid DefType keywords
// TODO: Validate letter ranges
```

**Suggested Fix:**
- Add validation for DefBool, DefByte, DefInt, DefLng, DefCur, DefSng, DefDbl, DefDec, DefDate, DefStr, DefObj, DefVar
- Validate letter ranges (A-Z, ensure start <= end)
- Return appropriate error for invalid ranges (e.g., "DefInt Z-A")

#### 7. Form Error Handling
**File:** `src/files/form/mod.rs:72`

```rust
// TODO: Handle errors from tokenization.
```

**Suggested Fix:**
- Properly propagate tokenization errors in form parsing
- Ensure partial form data is preserved when tokenization has warnings
- Add test cases for malformed form files

#### 8. ReferenceOrValue Property Handling
**File:** `src/language/controls/combobox.rs`

Multiple TODOs for handling ReferenceOrValue enum variants:
```rust
// TODO: Handle ReferenceOrValue for drag_icon
// TODO: Handle ReferenceOrValue for list
// TODO: Handle ReferenceOrValue for item_data
// TODO: Handle ReferenceOrValue for mouse_icon
```

**Suggested Fix:**
- Implement proper handling for both Reference and Value variants
- Add serialization support for referenced properties
- Create unified property handling system

#### 9. ListBox/ComboBox Data Serialization
**Files:** 
- `src/language/controls/listbox.rs:230,235`
- `src/language/controls/combobox.rs:243,246`

```rust
// TODO: Serialize item_data
// TODO: Serialize list
```

**Suggested Fix:**
- Implement ItemData array serialization
- Implement List array serialization
- Add format documentation for array serialization

#### 10. OLE Control Class Property
**File:** `src/language/controls/ole.rs:433,446`

```rust
// TODO: process Class property
// TODO: process DragIcon property
```

**Suggested Fix:**
- Implement Class property parsing for OLE controls
- Document expected Class property format
- Add test cases for various OLE control types

---

## Testing and Quality Assurance

### Test Coverage Expansion

#### 1. Unit Test Coverage
**Current State:** 26 test/example files identified

**Enhancements:**
- Increase test coverage for edge cases in tokenization
- Add more property parsing tests (especially for binary properties)
- Add negative test cases (malformed input, invalid syntax)
- Test memory limits and large file handling
- Add tests for all error paths

**Specific Areas Needing Tests:**
- String literal parsing with embedded quotes
- Keyword vs identifier disambiguation in all contexts
- Unicode characters in identifiers and strings
- Very long lines (>10,000 characters)
- Deeply nested control structures
- All DefType combinations and letter ranges

#### 2. Integration Tests
**Suggested Addition:**
- End-to-end tests parsing complete real-world VB6 projects
- Performance regression tests
- Cross-platform compatibility tests (Windows, Linux, macOS)
- Stress tests with pathological inputs

#### 3. Snapshot Testing
**Current:** Using `insta` for snapshot tests

**Enhancements:**
- Add snapshots for all VB6 statement types
- Add snapshots for all control types in forms
- Add snapshots for project file variations
- Add snapshots for error output formatting

#### 4. Fuzzing Improvements
**Current:** 9 fuzz targets covering all layers

**Enhancements:**
- Run fuzzing in CI with time limits
- Generate corpus from real-world VB6 projects
- Add structure-aware fuzzing for form files
- Create mutation dictionaries for VB6 keywords
- Document fuzzing findings and fixes
- Set up continuous fuzzing (e.g., OSS-Fuzz)

#### 5. Property-Based Testing
**Suggested Addition:**
- Use `proptest` or `quickcheck` for property-based testing
- Verify parser invariants (parse → serialize → parse = identity)
- Test tokenizer reversibility where possible
- Verify error recovery produces valid partial results

---

## Documentation Enhancements

### Code Documentation

#### 1. Module-Level Documentation
**Enhancement:** Expand documentation for complex modules:
- `parsers/cst/mod.rs` - Document CST construction algorithm
- `lexer/mod.rs` - Document tokenization state machine
- `files/resource/mod.rs` - Document FRX binary format in detail
- `language/controls/` - Add visual examples of each control

#### 2. Function Documentation
**Issues:**
- Some public functions lack examples
- Complex error types need usage examples
- Parser combinator functions need better documentation

**Suggested Improvements:**
- Add `# Examples` to all public functions
- Document panic conditions explicitly
- Add `# Errors` sections describing error cases
- Include performance characteristics for O(n²) or worse operations

#### 3. Architecture Documentation
**Suggested Addition:**
- Create `docs/ARCHITECTURE.md` explaining:
  - Red-green tree (rowan) usage rationale
  - Why CST vs AST approach
  - Memory management strategy
  - Error recovery philosophy
  - Performance characteristics of each layer

#### 4. Migration Guide
**Suggested Addition:**
- Create migration guide for VB6 → Rust
- Document common VB6 patterns and Rust equivalents
- Provide recipes for common analysis tasks
- Show how to traverse and query CST effectively

### API Documentation

#### 1. Tutorial Expansion
**Current:** Good getting-started guide exists

**Enhancements:**
- Add "Advanced Parsing" tutorial
- Add "Building Analysis Tools" tutorial
- Add "Custom Control Property Extraction" tutorial
- Add "Error Recovery Strategies" tutorial

#### 2. Cookbook/Recipes
**Suggested Addition:**
Create `docs/COOKBOOK.md` with recipes for:
- Finding all uses of a variable across a project
- Extracting form hierarchy
- Computing cyclomatic complexity
- Dead code detection
- Dependency graph construction
- Code style analysis

#### 3. FFI Documentation
**Current:** WASM bindings exist

**Enhancements:**
- Document WASM API thoroughly
- Add C FFI guide for embedding in other languages
- Add Python bindings (using PyO3)
- Add Node.js native module guide (using neon)

---

## Performance Optimizations

### Memory Optimization

#### 1. Clone Reduction
**Observation:** Some unnecessary clones found in hot paths

**Specific Cases:**
```rust
// src/parsers/cst/mod.rs:800
properties.insert(nested_group.name.clone(), Either::Right(nested_group));

// src/wasm.rs:166,169
let tokens = produce_tokens(token_stream.clone());
let cst = parsers::cst::parse(token_stream.clone());
```

**Suggested Fix:**
- Audit all `.clone()` calls in parser hot paths
- Use `Cow<str>` where appropriate
- Consider arena allocation for temporary strings
- Use reference counting (Rc/Arc) only where necessary

#### 2. String Allocation
**Suggested Enhancement:**
- Use `String` interning for commonly repeated strings (control names, property names)
- Consider using `smol_str` or `compact_str` for small strings (<23 bytes)
- Profile string allocation hotspots with `dhat` or `heaptrack`

#### 3. Token Stream Optimization
**Suggested Enhancement:**
- Consider implementing zero-copy token stream (tokens reference source)
- Benchmark current approach vs alternatives
- Add memory usage benchmarks to CI

### Parsing Performance

#### 1. Parallel Parsing
**Suggested Enhancement:**
- Parse multiple module files in parallel when loading projects
- Use rayon for parallel iteration over project files
- Preserve single-threaded option for simple cases

#### 2. Incremental Parsing
**Suggested Enhancement:**
- Investigate incremental reparsing for LSP use case
- Use rowan's green node caching effectively
- Document incremental parsing capabilities (if any)

#### 3. Lazy Parsing
**Suggested Enhancement:**
- Consider lazy parsing for large forms (parse headers first, bodies on demand)
- Add streaming API for very large projects
- Document memory/speed tradeoffs

---

## API and Functionality Enhancements

### Parser Features

#### 1. Source Map Generation
**Suggested Addition:**
- Generate source maps for CST nodes → original source locations
- Support for remapping after code transformations
- Useful for transpilers and error reporting

#### 2. Pretty Printer
**Suggested Addition:**
- Implement CST → formatted VB6 code
- Configurable formatting rules
- Preserve comments and whitespace when desired
- Support for code normalization

#### 3. Visitor Pattern API
**Current:** Manual CST traversal

**Enhancement:**
- Implement visitor pattern for CST traversal
- Support pre-order, post-order, and in-order traversal
- Add `accept()` method to CST nodes
- Provide example visitors (symbol table builder, dead code finder)

#### 4. Query API
**Suggested Addition:**
- CSS-selector-like query language for CST
- Examples: `"FunctionDeclaration[name='Main']"`, `"IfStatement > BlockStatement"`
- Makes analysis tools easier to write

#### 5. Symbol Resolution
**Suggested Addition:**
- Optional semantic analysis layer
- Build symbol tables
- Resolve identifier references
- Type inference (basic)
- Useful for advanced analysis and transformations

### File Format Support

#### 1. Additional File Types
**Suggested Addition:**
- `.vbw` (workspace) file parser
- `.ctl` (user control) file parser (if not already complete)
- `.dsr` (data report) file parser
- `.dob` (user document) file parser

#### 2. VB5 Compatibility
**Enhancement:**
- Document VB5 vs VB6 differences
- Add compatibility flags for VB5 projects
- Test with real VB5 projects

#### 3. VBA Support
**Suggested Addition:**
- Add VBA dialect support (similar to VB6 but with differences)
- Parse `.bas` modules from Excel/Word/Access
- Document VBA vs VB6 differences

---

## Error Handling Improvements

### Error Quality

#### 1. Error Messages
**Enhancement:**
- Review all error messages for clarity
- Add "did you mean?" suggestions for common typos
- Include context (what was expected vs what was found)
- Add error codes for programmatic error handling

#### 2. Error Recovery
**Current:** Partial recovery implemented

**Enhancement:**
- Document error recovery strategy
- Add more recovery points (e.g., after statement errors, continue parsing next statement)
- Ensure recovered CST is always valid (even if partial)
- Add tests specifically for error recovery

#### 3. Diagnostic Output
**Enhancement:**
- Add JSON error output format (for tool integration)
- Add IDE-friendly formats (VS Code, Language Server Protocol)
- Support multiple simultaneous error formats
- Add severity levels (error, warning, info, hint)

#### 4. Lint Warnings
**Suggested Addition:**
- Add optional linting warnings:
  - Unused variables/functions
  - Deprecated function usage (DoEvents, GoSub, etc.)
  - Naming convention violations
  - Complexity warnings (cyclomatic complexity, nesting depth)
  - Style violations (inconsistent indentation, line length)

---

## Tooling and DevOps

### CI/CD Improvements

#### 1. Additional CI Checks
**Current:** benchmarks.yml, coverage.yml, library.yml, wasm.yml

**Suggested Additions:**
- Clippy linting with `clippy::pedantic` and `clippy::cargo`
- `cargo-deny` for dependency auditing (already have deny.toml)
- `cargo-outdated` to check for outdated dependencies
- `cargo-audit` for security vulnerabilities
- Cross-platform testing (Windows, macOS, Linux)
- Minimum Supported Rust Version (MSRV) checking

#### 2. Automated Releases
**Suggested Addition:**
- Automated crate publishing on tag push
- Generate release notes from CHANGELOG.md
- Build and attach binary artifacts (if creating CLI tools)
- Update documentation automatically

#### 3. Dependency Management
**Current:** Using `deny.toml`

**Enhancement:**
- Enable Dependabot or Renovate for automated dependency updates
- Set up security advisory monitoring
- Document dependency update policy
- Consider feature flags for heavy dependencies

### Development Tools

#### 1. CLI Tool
**Suggested Addition:**
- Create standalone CLI tool for:
  - Parsing and validating VB6 projects
  - Converting to JSON/YAML
  - Generating statistics
  - Basic linting
  - Format checking

#### 2. Language Server Protocol (LSP)
**Suggested Addition:**
- Implement Language Server for VS Code/other editors
- Features:
  - Syntax highlighting
  - Error checking
  - Go to definition
  - Find references
  - Hover information
  - Code completion (basic)

#### 3. VS Code Extension
**Suggested Addition:**
- Package LSP as VS Code extension
- Add VB6 syntax themes
- Provide project scaffolding
- Integrate with debugger (if feasible)

#### 4. IDE Integration
**Suggested Addition:**
- IntelliJ IDEA plugin
- Sublime Text package
- Vim/Neovim plugin (using LSP)

---

## Code Organization and Maintenance

### Code Quality

#### 1. Clippy Warnings
**Action:** Run `cargo clippy --all-targets --all-features` and address:
- `clippy::missing_errors_doc`
- `clippy::missing_panics_doc`
- `clippy::must_use_candidate`
- `clippy::return_self_not_must_use`
- `clippy::similar_names`
- `clippy::too_many_lines`

#### 2. Code Duplication
**Suggested Enhancement:**
- Audit control property parsing for duplicated code
- Extract common patterns into shared utilities
- Consider procedural macros for reducing boilerplate

#### 3. Type Safety
**Observations:**
- Many `unwrap()` calls in tests (acceptable)
- Some `unwrap()` in library code (audit needed)
- `expect()` with good messages is better than `unwrap()`

**Action:**
- Audit all `unwrap()` calls outside of tests
- Replace with proper error handling or well-justified `expect()`
- Consider `#![deny(clippy::unwrap_used)]` for library code (not tests)

#### 4. Public API Auditing
**Suggested Review:**
- Ensure all public APIs are intentional
- Consider `#[doc(hidden)]` for internal-but-public items
- Review field visibility (many made private in v1.0.0, good!)
- Ensure consistent naming conventions

### Module Organization

#### 1. Large File Splitting
**Observation:** Some files are quite large

**Specific Cases:**
- `src/parsers/cst/mod.rs` - Consider splitting by statement type
- `src/language/controls/form.rs` - Large property list
- Large control property files

**Suggested Fix:**
- Split by logical groups
- Use submodules more extensively
- Keep public API at parent module level

#### 2. Feature Flags
**Suggested Addition:**
- Add feature flags for optional functionality:
  - `"serde"` - Already present, good!
  - `"wasm"` - Separate WASM bindings
  - `"image"` - Image processing (currently always on)
  - `"lint"` - Optional linting capabilities
  - `"symbols"` - Optional symbol resolution

---

## Security Considerations

### Input Validation

#### 1. Fuzz Testing (Already Implemented!)
**Status:** ✅ Excellent - 9 fuzz targets covering all layers

**Enhancement:**
- Run fuzzers longer (days/weeks) in CI environment
- Set up continuous fuzzing infrastructure
- Document known fuzzing-found issues and fixes

#### 2. Binary Format Parsing
**File:** `src/files/resource/mod.rs`

**Concerns:**
- Binary FRX parsing involves pointer arithmetic
- Must validate offsets and sizes before reading
- Check for integer overflow in size calculations

**Current State:** Some validation present

**Enhancement:**
- Audit all binary parsing for:
  - Bounds checking
  - Integer overflow in offset calculations
  - Infinite loop possibilities
  - Excessive memory allocation
- Add fuzzing specifically for malformed FRX files
- Consider using safe parsing libraries (e.g., `nom`, `binread`)

#### 3. Denial of Service
**Concerns:**
- Deeply nested structures (>1000 levels)
- Very long identifiers or strings
- Extremely large projects

**Suggested Mitigations:**
- Add depth limits for nested structures
- Add length limits for identifiers/strings
- Add timeout mechanisms for parsing
- Document resource limits
- Test with pathological inputs

#### 4. Unsafe Code Audit
**Action:** 
- Search for `unsafe` blocks (if any)
- Document safety invariants
- Consider eliminating unsafe code if possible
- Use `cargo-geiger` to audit unsafe usage in dependencies

### Dependency Security

#### 1. Supply Chain Security
**Current:** Using trusted crates

**Enhancement:**
- Regular security audits with `cargo-audit`
- Pin dependencies for reproducible builds
- Review transitive dependencies periodically
- Consider using `cargo-vet` for dependency auditing

#### 2. SBOM Generation
**Suggested Addition:**
- Generate Software Bill of Materials (SBOM)
- Include in releases
- Use `cargo-sbom` or similar tool

---

## Priority Recommendations

### Immediate (High Impact, Low Effort)

1. Fix string parsing issues in setattr.rs and rset.rs
2. Add validation to DefType statement parsing
3. Improve error messages with better context
4. Run and address clippy warnings
5. Add more unit tests for edge cases

### Short Term (High Impact, Medium Effort)

1. Complete FRX property parsing (icons, pictures, fonts)
2. Fix Array function parsing issues
3. Implement ReferenceOrValue handling for controls
4. Add CLI tool for basic operations
5. Expand test coverage to >90%

### Medium Term (Medium Impact, High Value)

1. Implement visitor pattern API
2. Add pretty printer (CST → formatted VB6)
3. Create comprehensive cookbook documentation
4. Implement symbol resolution layer
5. Add Language Server Protocol support

### Long Term (Strategic Enhancements)

1. Build VS Code extension with LSP
2. Add VBA dialect support
3. Implement incremental parsing
4. Create analysis toolkit ecosystem
5. Build web-based VB6 analysis tools

---

## Metrics and Goals

### Current State
- ✅ No compilation errors
- ✅ Good core documentation
- ✅ Comprehensive fuzzing setup
- ✅ Decent test coverage (26 test/example files)
- ✅ CI/CD pipelines established
- ⚠️ Several TODOs in codebase (~50 identified)
- ⚠️ Some incomplete feature implementations

### Suggested Goals for Next Version

**Version 1.1.0**
- ✅ Resolve all high-priority TODOs
- ✅ Add 50+ more test cases
- ✅ Complete FRX property parsing
- ✅ API stability guarantees established

**Version 1.2.0**
- ✅ Visitor pattern API
- ✅ Pretty printer
- ✅ CLI tool released
- ✅ >90% test coverage

**Version 2.0.0**
- ✅ Language Server Protocol
- ✅ Symbol resolution layer
- ✅ VBA support
- ✅ Breaking API improvements

---

## Contributing

This enhancement document should be used to:
1. Prioritize development efforts
2. Guide contributor onboarding
3. Track progress on improvements
4. Set roadmap for future versions

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines on implementing these enhancements.

---

*Document generated: February 23, 2026*  
*Based on commit: [current HEAD]*  
*Next review: Every major version release*
