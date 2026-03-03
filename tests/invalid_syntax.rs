// Invalid syntax tests - organized by error category
// These tests verify that the parser handles invalid syntax gracefully,
// producing reasonable CST structures and meaningful error messages

#[path = "invalid_syntax/missing_end.rs"]
mod missing_end;

#[path = "invalid_syntax/missing_keywords.rs"]
mod missing_keywords;

#[path = "invalid_syntax/mismatched_keywords.rs"]
mod mismatched_keywords;

#[path = "invalid_syntax/invalid_literals.rs"]
mod invalid_literals;

#[path = "invalid_syntax/invalid_declarations.rs"]
mod invalid_declarations;

#[path = "invalid_syntax/invalid_control_flow.rs"]
mod invalid_control_flow;
