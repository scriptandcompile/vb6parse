//! Tests for enum members with square bracket attributes.
//!
//! VB6 allows enum members to have square bracket attributes like `[Midi]`, `[sID]`, etc.
//! This test file verifies that these are correctly parsed as `LeftSquareBracket` and
//! `RightSquareBracket` tokens instead of Unknown tokens.

use vb6parse::*;

/// Test that enum members with square bracket attributes are parsed correctly.
///
/// Previously, the square brackets `[` and `]` were parsed as Unknown tokens.
/// Now they are correctly recognized as `LeftSquareBracket` and `RightSquareBracket`.
#[test]
fn test_enum_with_square_bracket_attributes() {
    let source = r"
Public Enum MyEnum
    [FirstValue]
    [SecondValue]
    [ThirdValue]
End Enum
";

    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should parse successfully");

    let tree = cst.to_serializable();
    let tree_str = format!("{:?}", tree);

    // Verify that there are no Unknown tokens
    assert!(
        !tree_str.contains("Unknown"),
        "CST should not contain any Unknown tokens for enum with square bracket attributes"
    );

    // Verify that square brackets are recognized
    assert!(
        tree_str.contains("LeftSquareBracket"),
        "CST should contain LeftSquareBracket tokens"
    );
    assert!(
        tree_str.contains("RightSquareBracket"),
        "CST should contain RightSquareBracket tokens"
    );
}

/// Test that keywords inside square brackets in enums are parsed correctly.
///
/// VB6 allows keywords to be used as identifiers when enclosed in square brackets.
#[test]
fn test_enum_with_keyword_in_brackets() {
    let source = r"
Public Enum MediaType
    [Beep]
    [Error]
    [End]
End Enum
";

    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should parse successfully");

    let tree = cst.to_serializable();
    let tree_str = format!("{:?}", tree);

    // Verify that there are no Unknown tokens
    assert!(
        !tree_str.contains("Unknown"),
        "CST should not contain any Unknown tokens for enum with keywords in brackets"
    );
}

/// Test enum with square brackets - ensures no Unknown tokens.
#[test]
fn test_enum_status_with_brackets() {
    let enum_source = r"
Public Enum Status
    [Active]
    [Inactive]
End Enum
";

    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", enum_source).unpack();
    let cst = cst_opt.expect("CST should parse successfully");

    let tree = cst.to_serializable();
    let tree_str = format!("{:?}", tree);

    assert!(
        !tree_str.contains("Unknown"),
        "Enum with square brackets should not have Unknown tokens"
    );
}
