# Recent Fuzzing Failures

This document tracks crashes and bugs discovered by fuzz testing.

## ProjectFile Fuzzer

- **Crash discovered**: UTF-8 char boundary panic when parsing Chinese filenames encoded with Windows-1252
- **Location**: `src/io/source_stream.rs:310` - string slicing on multi-byte character boundary  
- **Test Case**: `fuzz/artifacts/project_file/crash-4ff6b15016e5af289309a9b4787e284b530aa3fc`
- **Reproduction**: `cargo +nightly fuzz run project_file fuzz/artifacts/project_file/crash-4ff6b15016e5af289309a9b4787e284b530aa3fc`
- **Status**: Needs fix - SourceStream string slicing doesn't handle multi-byte UTF-8 characters correctly
