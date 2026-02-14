# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project (tries!) to adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- Removed `winnow` dependency - no longer used in the codebase.
- Removed `borrow` dependency - no longer used in the codebase.
- `Getting Started` guide now loads code examples dynamically from GitHub CDN for always-current documentation
- Renamed `source_file` field to `file_name` in `TokenStream`
- Made `TokenStream` fields private with accessor methods (`file_name()`, `offset()`)
- Reorganized module structure: created IO module and Lexer module
- Moved source_file and source_stream into new `io/` module
- Moved tokenization components into new `lexer/` module
- Reorganized parser components into dedicated modules (expressions, control_flow, declarations, objects)
- Moved all library statements and functions into their correct categories
- Improved module organization for better separation of concerns

### Fixed
- Double quote handling in string parsing  
- Backslash flag handling in strings
- DateTime literal parsing no longer confused by single # characters
- Missing EndKeyword in various parsing contexts
- All C-GETTER violations - removed `get_` prefix from parameterless getters
- ProjectFile fields are now mostly private (except `project_type`, `other_properties`, `properties`)
- ParseResult fields are now private with accessor methods
- FormResourceFile naming consistency (previously had some FormResource references)
- VB6 Color BGR format handling (was incorrectly using RGB)
- Form/MDIForm removed from ControlKind enum (now part of FormRoot enum)
- Parser layer now distinguishes between top-level forms and child controls
- Examples updated to use FormRoot API instead of ControlKind::Form
- Token::DateLiteral renamed to Token::DateTimeLiteral for clarity

### Added
- "Contributing" section in README.md linking to CONTRIBUTING.md
- New CONTRIBUTING.md file with a list of low and medium difficulty enhancements for new contributors
- Entry-level tasks documented: documentation improvements, new examples, increased test coverage, expanding FRX property mapping
- Improved onboarding documentation for new contributors
- Documented ConcreteSyntaxTree and CstNode navigation methods in README.md, including examples of tree traversal and node queries
- Comprehensive fuzz testing support with 9 fuzz targets covering all parser layers
- Fuzz targets for: SourceFile decoding, SourceStream, tokenization, CST parsing, ProjectFile, ClassFile, ModuleFile, FormFile, and FormResourceFile
- Corpus and artifacts directories for fuzzing with cargo-fuzz and libFuzzer
- FormRoot type system for type-safe form parsing with Form and MDIForm variants
- InvalidTopLevelControl error variant for invalid form root elements  
- parse_properties_block_to_form_root() function for top-level form parsing
- build_form_root() helper function in parser layer
- Additional missing keywords to language tokens (18+ new keywords including EndKeyword)
- Improved DateTime literal parsing to correctly handle date/time literals
- More comprehensive date/time parsing tests
- Pre-commit hook that runs `cargo check --examples` to prevent example bit-rot
- Documentation examples in `examples/docs/` directory (hello_world, project_parsing, error_handling, tokenization, cst_navigation, form_parsing)
- Explicit example declarations in Cargo.toml for examples in subdirectories
- Support for RmDir statements
- Support for Resume statements
- Support for Randomize statements
- Support for RaiseEvent statements
- Support for Put statements
- Support for Public variable declarations including 'WithEvents' keyword
- Support for Private variable declarations including 'WithEvents' keyword
- Support for Print statements
- Support for parsing Option Private statements
- Support for parsing Option Compare statements with keywords as identifiers
- Support for parsing Option Base statements
- Support for Open statement
- Support for OnGoTo and OnGoSub statements
- Support for On Error statements
- Support for parsing Name statements
- Support for parsing MkDir statements
- Support for parsing MidB statements
- Support for parsing Mid statements
- Support for parsing LSet statements
- Support for parsing Unlock statements
- Support for parsing Lock statements
- Support for parsing Load statement
- Support for parsing Line Input statements
- Support for parsing Let statement
- Support for parsing Kill statement
- Support for parsing Input statements
- Support for parsing Implements statements
- Support for GoSub and Return syntax
- Support for parsing Get statements
- Support for FileCopy statement parsing
- Support for event statements
- Support for Error parsing
- Support for parsing Erase statements
- Support for enum parsing
- Support for Reset statement
- Support for DeleteSettings built-in statement
- Support for Deftype statements
- Support for parsing external Sub/Function declarations using Declare keyword
- Support for Date built-in
- Support for Close built-in statement
- Support for Property statement parsing
- Support for Exit statements
- Support for GoTo and inline if-then statements
- Support for Select Case statements
- Support for With statements
- Support for label statements
- Support for SetStatements
- Support for For Each loops
- Support for Do...Loop parsing
- Support for Call statements
- Added `consume_until_after` helper method

### Changed
- Updated to image 0.25.9, rowan 0.16.1, and insta 1.44.0
- Removed 'Number' as token type, added IntegerLiteral, LongLiteral, SingleLiteral, DoubleLiteral, DecimalLiteral, CurrencyLiteral, and DateLiteral
- Renamed `is_keyword` to `at_keyword` for consistency
- Created SerializableTree for snapshot testing
- Moved various parsing components into dedicated modules for better organization
- Improved line continuation support, especially in functions
- Centralized all statement parsing processing
- Improved conditional parsing and control flow parsing
- Broke up large parsing files into manageable modules
- Improved test organization with utility methods

### Fixed
- Fixed spelling in various documentation and enum names
- Fixed error message reporting
- Various Clippy warnings addressed

## [0.5.0] - 2024

### Added
- Full FRX (form resource) file support
- Resource file resolver for loading binary resources, images, lists, and text from .frx files
- Support for multiple FRX header formats (4-byte, 8-byte, 12-byte)
- Support for zero-size icon resources
- Support for list parsing and buffer loading from FRX files
- Comprehensive test coverage for FRX loading across multiple real VB6 projects

### Changed
- Improved resource offset parsing to handle offsets larger than 4 digits
- Better support for FRX property parsing including text, captions, and images
- Enhanced error checking for malformed resource files

## [0.4.1] - 2024

### Fixed
- Control type naming: Changed 'Ole' to 'OLE' in control creation structure
- Various Clippy warnings

## [0.4.0] - 2024

### Added
- Support for MDIForm controls
- Support for nested property groups (common in custom controls)
- Support for GUID in BeginProperty parsing
- Support for property groups
- Support for custom control property loading
- Parsing support for Shape, ScrollBar, PictureBox, OptionButton, OLE, ListBox, Line, Label, Image, FileListBox, DriveListBox, DirListBox, Data, ComboBox, CheckBox, Timer, and TextBox controls
- Menu control creation using parsed data
- Property parsing for Forms, Menus, and Frames
- Support for control tags, WindowState, StartUpPosition, ScaleMode, PaletteMode, OLEDropMode, MousePointer, LinkMode, ForeColor, FillStyle, FillColor, DrawStyle, DrawMode, ClipControls, BorderStyle, and Appearance properties

### Changed
- Switched from `str` to `BStr` for better non-English code support
- Improved parsing of properties to use generic builder functions
- Enhanced project property parsing with numerous enum types
- VB6FileAttributes now use enums instead of booleans
- Improved qualified name handling for non-English character support

### Fixed
- Threading model parsing
- True/false value parsing (-1 vs 1 vs 0)
- Various property default values
- CompatibleMode enum values
- PCode vs NativeCode compilation type values

## [0.3.0] - 2024

### Added
- Support for higher character set (128-255) in variable names for non-English source code
- Language detection with `is_english_code()` function
- Custom control hashing functions for property building

### Changed
- Updated winnow and clap versions
- Improved parsing of non-English VB6 source code

### Fixed
- Various Clippy warnings and code quality improvements

## [0.2.0] - 2024

### Added
- Comprehensive support for VB6 class file parsing
- Support for VB6 module file parsing
- Error reporting using Ariadne for beautiful error messages
- Miette integration for enhanced error context
- VB6Stream for simplified parsing without winnow
- SourceStream for lower-level character stream parsing
- TokenStream bundling tokens with source information
- Support for double-quote escaped strings
- Support for numerous VB6 tokens and keywords
- Extensive documentation for VB6Token enum variants

### Changed
- Switched from winnow to custom parsing approach using SourceStream
- Moved from miette to ariadne for error reporting
- Improved error reporting with line and column information
- Moved parsers to dedicated parsers module
- Better organization of language-specific elements into language module

### Fixed
- Quote/escape handling in strings
- Whitespace and tab handling in quoted values
- Variable name parsing with keywords
- Various parsing edge cases

## [0.1.0] - Initial Release

### Added
- Initial VB6 project (.vbp) file parsing
- Support for parsing project references
- Support for parsing modules, classes, and forms declarations
- Support for project properties and settings
- Basic form (.frm) structure parsing
- Windows-1252 encoding support
- Basic tokenization of VB6 code
- Integration tests with real VB6 projects
- Criterion benchmarking support
- Comprehensive error types for different VB6 file types

[Unreleased]: https://github.com/scriptandcompile/vb6parse/compare/v0.5.1...HEAD
[0.5.1]: https://github.com/scriptandcompile/vb6parse/releases/tag/v0.5.1
[0.5.0]: https://github.com/scriptandcompile/vb6parse/releases/tag/v0.5.0
[0.4.1]: https://github.com/scriptandcompile/vb6parse/releases/tag/v0.4.1
[0.4.0]: https://github.com/scriptandcompile/vb6parse/releases/tag/v0.4.0
[0.3.0]: https://github.com/scriptandcompile/vb6parse/releases/tag/v0.3.0
[0.2.0]: https://github.com/scriptandcompile/vb6parse/releases/tag/v0.2.0
[0.1.0]: https://github.com/scriptandcompile/vb6parse/releases/tag/v0.1.0
