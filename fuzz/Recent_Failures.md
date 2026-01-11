# Recent Fuzzing Failures

This document tracks crashes and bugs discovered by fuzz testing.

## FormResourceFile Fuzzer (NEW!)

- **Crash discovered**: Integer underflow panic in resource file parsing
- **Location**: `src/files/resource/mod.rs:593` - attempt to subtract with overflow
- **Test Case**: `fuzz/artifacts/form_resource/crash-141eb5122a02efc395e538e2a7e54a6e38f5c8ad`
- **Trigger**: Malformed 12-byte header with "lt\0\0" signature and size values causing subtraction underflow: `[0, 0, 0, 0, 108, 116, 0, 0, 0, 0, 0, 0]`
- **Reproduction**: `cargo +nightly fuzz run form_resource fuzz/artifacts/form_resource/crash-141eb5122a02efc395e538e2a7e54a6e38f5c8ad`
- **Status**: Needs fix - Resource parser must validate size fields and use checked arithmetic

## Common Bug: UTF-8 Char Boundary Panic in SourceStream

All fuzzers discovered the same critical bug: UTF-8 character boundary panic when parsing Windows-1252 encoded files with non-ASCII characters.

**Root Cause**: `src/io/source_stream.rs:163` and `:310` - string slicing operations don't handle multi-byte UTF-8 characters correctly after Windows-1252 decoding.

**Status**: Needs fix - SourceStream string slicing must use character boundary checks or operate on byte slices.

### ProjectFile Fuzzer

- **Crash location**: `src/io/source_stream.rs:310`
- **Test Case**: `fuzz/artifacts/project_file/crash-4ff6b15016e5af289309a9b4787e284b530aa3fc`
- **Trigger**: Chinese filenames (TLB×é¼þ - bytes 215,233,188,254) in path references
- **Reproduction**: `cargo +nightly fuzz run project_file fuzz/artifacts/project_file/crash-4ff6b15016e5af289309a9b4787e284b530aa3fc`

### ClassFile Fuzzer

- **Crash location**: `src/io/source_stream.rs:163`
- **Test Case**: `fuzz/artifacts/class_file/crash-140a84fa3e8ae7277d60d64a6ae2db730383c928`
- **Trigger**: String literal with "Ô" character (bytes 212,170)
- **Reproduction**: `cargo +nightly fuzz run class_file fuzz/artifacts/class_file/crash-140a84fa3e8ae7277d60d64a6ae2db730383c928`

### ModuleFile Fuzzer

- **Crash location**: `src/io/source_stream.rs:163`
- **Test Case**: `fuzz/artifacts/module_file/crash-9ed078c9446aa6eca02c2e40debd7cc6dd839a1b`
- **Trigger**: Cyrillic characters in Debug.Print statements (bytes 209,229,233,247,224,241...)
- **Reproduction**: `cargo +nightly fuzz run module_file fuzz/artifacts/module_file/crash-9ed078c9446aa6eca02c2e40debd7cc6dd839a1b`

### FormFile Fuzzer

- **Crash location**: `src/io/source_stream.rs:163`
- **Test Case**: `fuzz/artifacts/form_file/crash-f22f053be128d3e1e4c7b29215598aafb638ee46`
- **Trigger**: Chinese characters in form caption and comments (bytes 179,230,189,117,169...)
- **Reproduction**: `cargo +nightly fuzz run form_file fuzz/artifacts/form_file/crash-f22f053be128d3e1e4c7b29215598aafb638ee46`

