# Contributing to VB6Parse

First off, thank you for considering contributing to VB6Parse! It's people like you that make open source great.

## Where to Start

If you're new to the project, a great place to start is by looking at the issues labeled "good first issue" or "help wanted". These are issues that are relatively easy to get started with.

Here are some other ideas for contributions, categorized by difficulty:

### Low Difficulty

*   **Improve Documentation:**
    *   Add more examples to the `examples/` directory for parsing different VB6 constructs.
    *   Expand the documentation for the `language` and `syntax` modules. Many of the enums and structs could benefit from more detailed explanations and examples.
    *   Document the error types in the `errors` module more thoroughly.
    *   Add or improve documentation in `src/syntax/library/functions` and `src/syntax/library/statements` (documentation is automatically published to the website).

*   **Increase Test Coverage:**
    *   Add more unit tests for individual parsers. The existing tests in `tests/` can be used as a template.
    *   Add more snapshot tests for different VB6 files. The `tests/data` directory contains real VB6 projects that can be used for testing.

### Medium Difficulty

*   **Form Resource (`.frx`) Parsing:**
    *   The `documents/FRX_format.md` file provides information about the format.
    *   Implementing the parsing for more control properties from the binary blobs would be a valuable contribution.

*   **Implement More VB6 Statements:**
    *   The `src/syntax/statements` module has many `// TODO` comments for statements that are not yet implemented. Implementing one of these statements would be a good medium-difficulty task.

## Development Setup

1.  Fork the repository and clone it to your local machine.
2.  Install the Rust toolchain: `rustup-init.sh`
3.  Initialize the git submodules: `git submodule update --init --recursive`
4.  Run the tests to make sure everything is working: `cargo test`

## Documentation Generation

### Library Reference Documentation

When adding or updating VB6 library function/statement documentation in `src/syntax/library/`, regenerate the web documentation:

```bash
# Requires Python 3.6+ and markdown library
pip install markdown

# Generate library documentation
python3 scripts/generate-library-docs.py --clean
```

The generated HTML pages will be created in `docs/library/` and are automatically deployed to GitHub Pages.

### Coverage and Benchmarks

To update the coverage and benchmark data on the website:

```bash
# Generate coverage data (requires cargo-llvm-cov)
python3 scripts/generate-coverage.py

# Generate benchmark data (requires criterion benchmarks to be run first)
python3 scripts/generate-benchmarks.py
```

## Submitting a Pull Request

1.  Create a new branch for your changes.
2.  Make your changes and commit them with a descriptive commit message.
3.  Push your changes to your fork.
4.  Open a pull request on the main repository.

Thank you for your contributions!
