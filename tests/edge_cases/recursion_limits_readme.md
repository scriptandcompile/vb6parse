# Recursion Limits Tests

This test module contains tests for verifying the parser's behavior when handling deeply nested or complex recursive structures.

## Purpose

These tests are designed to:

1. **Document current behavior** - Establish baseline behavior for parsing deeply nested structures
2. **Prevent regressions** - Ensure parser changes don't unexpectedly break on legitimate nesting
3. **Prepare for limit enforcement** - When recursion depth limits are implemented, these tests will verify proper error handling

## Test Coverage


### Statement List Recursion
- `deeply_nested_if_statements` - Tests indirect recursion through If statement nesting (50 levels)
- `deeply_nested_for_loops` - Tests For loop nesting (30 levels)
- `deeply_nested_select_case` - Tests Select Case statement nesting (20 levels)
- `deeply_nested_with_blocks` - Tests With block nesting (25 levels)
- `mixed_nested_control_flow` - Tests various control flow structures nested together

### Expression Recursion  
- `deeply_nested_parentheses` - Tests parenthesized expression nesting (100 levels)
- `long_binary_operation_chain` - Tests long chains of binary operations (200 operations)
- `complex_nested_boolean_expression` - Tests complex boolean expressions with And/Or

### Combined Recursion
- `combined_expression_and_statement_nesting` - Tests interaction between expression and statement recursion

## Current Behavior (Before Limits)

All tests currently use **moderate depths** that should parse successfully without stack overflow on typical machines. The depths chosen are:

- If statements: 50 levels
- For loops: 30 levels  
- Select Case: 20 levels
- With blocks: 25 levels
- Parentheses: 100 levels
- Binary operations: 200 operations

These represent realistic upper bounds that might occur in real VB6 code, particularly in auto-generated forms or complex business logic.

## Future Updates (After Limit Implementation)

When recursion depth limits are implemented, these tests should be updated to:

1. **Test at limits** - Some tests should be modified to approach or exceed the configured limits:
   - MAX_EXPRESSION_DEPTH (proposed: 500)
   - MAX_CONTROL_DEPTH (proposed: 1000 for forms)
   - MAX_STATEMENT_DEPTH (proposed: 500)
   - MAX_PROPERTY_GROUP_DEPTH (proposed: 100)

2. **Verify error messages** - Tests should check that appropriate error messages are generated when limits are exceeded

3. **Test recovery** - Verify that the parser produces a reasonable CST even when depth limits are hit

## Running These Tests

```bash
# Run all recursion limit tests
cargo test --test edge_cases recursion_limits

# Run a specific test
cargo test --test edge_cases recursion_limits::deeply_nested_if_statements

# Run with output visible
cargo test --test edge_cases recursion_limits -- --nocapture
```

## Notes for Developers

- These tests generate relatively large snapshot files due to deeply nested structures
- Test execution time should remain under 1 second for all tests
- If tests start failing after parser changes, review whether:
  - The nesting depth changed behavior (may need snapshot update)
  - A recursion limit was added (tests may need updating per above)
  - An unintended regression occurred (parser change should be reviewed)
