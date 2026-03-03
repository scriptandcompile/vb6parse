# Invalid Syntax Tests

This test suite validates the parser's behavior when encountering invalid VB6 syntax. These tests focus on **CST (Concrete Syntax Tree) level errors**, not semantic or type-checking errors.

## Purpose

These tests serve multiple purposes:

1. **Document current behavior**: Capture how the parser currently handles invalid syntax
2. **Verify resilient parsing**: Ensure the parser produces reasonable CST structures even with syntax errors
3. **Track error reporting**: Snapshot the failure messages to ensure they're meaningful
4. **Baseline for improvements**: Provide a foundation for enhancing error reporting

## Test Organization

Similar to other test suites (e.g., `form.rs`), this suite uses:
- `tests/invalid_syntax.rs` - Main test file that imports test modules
- `tests/invalid_syntax/` - Folder containing test modules organized by error category
- `snapshots/tests/invalid_syntax/` - Snapshot files for CST and failure verification

## Test Categories

### Missing End Statements (`missing_end.rs`)

Tests for missing closing keywords like:
- `End Sub` - Missing subroutine terminator
- `End Function` - Missing function terminator
- `End Property` - Missing property terminator
- `End If` - Missing If block terminator
- `End Type` - Missing Type definition terminator
- `End Select` - Missing Select Case terminator
- Nested missing ends - Multiple missing terminators

**Current Behavior**: The parser uses resilient parsing and does NOT report failures for missing End statements. It automatically closes constructs at EOF or when another construct begins. The CST structures are reasonable and complete.

### Missing Required Keywords (`missing_keywords.rs`)

Tests for missing required keywords in VB6 statements:
- Missing `Then` in If statement
- Missing `To` in For loop
- Missing `As` in Dim statement
- Missing `=` in Const declaration
- Missing `Case` in Select Case statement
- Missing `Loop` in Do statement
- Missing `Next` in For loop

**Current Behavior**: The parser uses resilient parsing and does NOT report failures for missing required keywords. It attempts to parse what it can and creates reasonable CST structures. In most cases, it treats subsequent tokens as part of the statement or as new statements.

### Mismatched Keywords (`mismatched_keywords.rs`)

Tests for mismatched opening and closing keywords in VB6 constructs:
- `Sub` with `End Function`
- `Function` with `End Sub`
- `Property Get` with `End Sub`
- `Property Let` with `End Function`
- `If` with `End Select`
- `Select Case` with `End If`
- `For` with `Wend` (should be `Next`)
- `Do While` with `Next` (should be `Loop`)
- `While` with `Loop` (should be `Wend`)
- `Type` with `End Enum` (should be `End Type`)

**Current Behavior**: The parser uses resilient parsing and does NOT report failures for mismatched keywords. It treats the mismatched end keyword as closing the construct, creating a valid CST structure despite the semantic mismatch.

### Invalid Literals (`invalid_literals.rs`)

Tests for malformed literal values in VB6:
- Unclosed string literal
- String with incomplete quote escape
- Multiple decimal points in numeric literal
- Invalid hexadecimal literal (non-hex digits)
- Invalid octal literal (invalid digits)
- Invalid date literal (bad month value)
- Unclosed date literal
- Invalid scientific notation (missing exponent)
- Invalid number suffix
- Number with leading zeros

**Current Behavior**: The parser's tokenizer handles most invalid literals. String and date literals may be parsed as unclosed tokens. Invalid numeric formats are typically tokenized as separate tokens or identifiers. All failure snapshots are empty arrays, indicating resilient tokenization.

### Invalid Declarations (`invalid_declarations.rs`)

Tests for malformed declaration statements in VB6:
- Missing identifier in `Dim` statement
- Missing type annotation after `As`
- Missing function return type after `As`
- Missing subroutine/function name
- Duplicate visibility modifiers (`Public Public`)
- Conflicting visibility modifiers (`Public Private`)
- Missing array bounds in declaration
- Missing constant value after `=`
- Missing member name in `Type` definition
- Missing parameter name in function signature
- Missing default value after `=` in Optional parameter
- Duplicate `Static` modifiers
- Missing enum member value after `=`

**Current Behavior**: The parser uses resilient parsing and does NOT report failures for invalid declarations. It attempts to parse what tokens are present and creates CST structures that may include incomplete or malformed nodes. Missing identifiers result in nodes without the expected identifier children. Duplicate modifiers are both included in the CST.

### Invalid Control Flow (`invalid_control_flow.rs`)

Tests for control flow statements used in invalid contexts in VB6:
- `Exit Sub` outside of a subroutine
- `Exit Function` outside of a function (e.g., in Sub)
- `Exit Property` outside of a property (e.g., in Function)
- `Exit For` outside of a For loop
- `Exit Do` outside of a Do loop (e.g., in While loop)
- `GoTo` with missing label
- `GoSub` with missing label
- `On Error` with missing destination
- `On Error GoTo` with missing target
- `Resume` without On Error context
- Nested `Exit` statements in wrong context (Exit Sub inside loops)
- `Return` statement in module (only valid in class event handlers)
- `Stop` statement with arguments (should have none)
- `On Error Resume` with invalid combination

**Current Behavior**: The parser uses resilient parsing and does NOT report failures for control flow statements in invalid contexts. Exit statements are parsed as valid statements regardless of their surrounding context. Missing labels/targets after GoTo/GoSub/On Error result in incomplete statement nodes. These contextual errors would need to be caught by semantic analysis.

### Invalid Parameter Lists (`invalid_parameter_list.rs`)

Tests for malformed parameter lists in subroutines, functions, and properties:
- Missing comma between parameters
- Trailing comma in parameter list
- Missing parameter after comma
- Duplicate `ByVal` modifier
- Conflicting `ByVal` and `ByRef` modifiers
- `Optional` parameter before required parameter
- `ParamArray` not as last parameter
- `ParamArray` with `ByVal` modifier (not allowed)
- Multiple consecutive commas
- Parameter missing `As` keyword
- `Optional` parameter with missing default value after `=`
- Duplicate `Optional` modifier
- `ParamArray` without array parentheses `()`
- Parameter with type character (`%`, `$`, etc.) and `As` clause (redundant)

**Current Behavior**: The parser uses resilient parsing and does NOT report failures for invalid parameter lists. Missing commas may cause parameters to be parsed as separate identifiers or expressions. Duplicate modifiers are typically both included in the CST. Missing elements result in incomplete parameter nodes. Parameter ordering rules (Optional after required, ParamArray last) and modifier conflicts would need to be validated by semantic analysis.

## Test Structure

Each test follows this pattern:

```rust
#[test]
fn test_name() {
    let source = r"
        <invalid VB6 code>
    ";

    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();

    // Log current behavior
    eprintln!("=== Failures ===");
    eprintln!("Number of failures: {}", failures.len());
    for failure in &failures {
        failure.eprint();
    }
    eprintln!("=== End Failures ===");

    // Verify CST is parseable
    let cst = cst_opt.expect("CST should be present even with syntax errors");
    let tree = cst.to_serializable();
    
    // Snapshot CST structure
    insta::assert_yaml_snapshot!("test_name_cst", tree);
    
    // Snapshot failure messages
    let failure_messages: Vec<String> = failures
        .iter()
        .map(|f| format!("{:?}", f))
        .collect();
    insta::assert_yaml_snapshot!("test_name_failures", failure_messages);
}
```

## Running the Tests

```bash
# Run all invalid syntax tests
cargo test --test invalid_syntax

# Run with output (see failure messages)
cargo test --test invalid_syntax -- --nocapture

# Update snapshots
cargo insta test --test invalid_syntax --review
```

## Adding New Test Categories

To add a new category of syntax errors:

1. Create a new test file in `tests/invalid_syntax/` (e.g., `invalid_operators.rs`)
2. Add tests following the pattern above
3. Import the module in `tests/invalid_syntax.rs`:
   ```rust
   #[path = "invalid_syntax/invalid_operators.rs"]
   mod invalid_operators;
   ```
4. Run tests and accept snapshots

## Future Improvements

These tests provide a baseline for future work on error reporting:

- [ ] Add error reporting for missing End statements
- [ ] Add error reporting for missing required keywords
- [ ] Add tests for mismatched keywords (e.g., `End Function` after `Sub`)
- [ ] Add tests for invalid expressions
- [ ] Add tests for invalid declarations
- [ ] Add tests for invalid control flow
- [ ] Add tests for duplicate declarations
- [ ] Add tests for invalid label usage
- [ ] Ensure error messages are helpful and point to the right location
