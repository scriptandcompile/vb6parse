# Test Organization

## Structure

Tests are organized by file type and data repository:

```
tests/
├── form.rs              # Entry point for form tests
├── form/
│   ├── vb6_code.rs      # Form tests for vb6-code repository
│   └── audiostation.rs  # Form tests for audiostation repository
├── form_resource.rs     # Entry point for form resource tests
├── form_resource/
│   ├── vb6_code.rs      # Form resource tests for vb6-code repository
│   └── audiostation.rs  # Form resource tests for audiostation repository
├── module.rs            # Entry point for module tests
├── module/
│   ├── vb6_code.rs      # Module tests for vb6-code repository
│   └── audiostation.rs  # Module tests for audiostation repository
├── class.rs             # Entry point for class tests
├── class/
│   ├── vb6_code.rs      # Class tests for vb6-code repository
│   └── audiostation.rs  # Class tests for audiostation repository
├── project.rs           # Entry point for project tests
└── project/
    ├── vb6_code.rs      # Project tests for vb6-code repository
    └── audiostation.rs  # Project tests for audiostation repository
```

## Adding a New Repository

To add tests for a new data repository (e.g., `my-repo`):

1. **Create test files** in each type directory:
   - `tests/form/my_repo.rs`
   - `tests/form_resource/my_repo.rs`
   - `tests/module/my_repo.rs`
   - `tests/class/my_repo.rs`
   - `tests/project/my_repo.rs`

2. **Update entry point files** to include the new module:

   In `tests/form.rs`:
   ```rust
   #[path = "form/my_repo.rs"]
   mod my_repo;
   ```

   (Repeat for `form_resource.rs`, `module.rs`, `class.rs`, `project.rs`)

3. **Write tests** following the existing pattern. Key points:
   - Add required imports at the top (see existing files for examples)
   - Use `./tests/data/my-repo/...` for file paths with `std::fs::read()`
   - Use `../data/my-repo/...` for file paths with `include_bytes!()`
   - Use `../../snapshots/tests/{type}/my-repo/` for snapshot paths
   - Each test function needs `#[test]` attribute

## Running Tests

```bash
# Run all tests
cargo test

# Run tests for a specific file type
cargo test --test form
cargo test --test form_resource
cargo test --test module
cargo test --test class
cargo test --test project

# Run tests for a specific repository
cargo test vb6_code
cargo test audiostation

# Run a specific test
cargo test artificial_life_form_load
```

## Snapshot Management

New or changed tests will create `.snap.new` files. Review and accept them:

```bash
cargo insta review
```

Or accept all at once:

```bash
cargo insta test --review
```
