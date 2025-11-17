//! Integration test for CST parsing functionality

use vb6parse::language::VB6Token;
use vb6parse::parsers::{parse, SourceStream, SyntaxKind, ConcreteSyntaxTree};
use vb6parse::tokenize::tokenize;

#[test]
fn cst_basic_parsing() {
    let code = "Sub Test()\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Now has 1 child: the SubStatement node (with EOF inside as Unknown)
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Sub Test()\n");
}

#[test]
fn cst_with_comments() {
    let code = "' This is a comment\nSub Main()\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // Now has 3 children: comment token, newline token, SubStatement (with EOF inside)
    assert_eq!(cst.child_count(), 3);
    assert!(cst.text().contains("' This is a comment"));
    assert!(cst.text().contains("Sub Main()"));
}

#[test]
fn cst_complex_code() {
    let code = r#"Private Sub Calculate(ByVal x As Integer)
    Dim result As Long
End Sub
"#;
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let text = cst.text();
    assert!(text.contains("Private Sub Calculate"));
    assert!(text.contains("ByVal x As Integer"));
    assert!(text.contains("Dim result As Long"));
    assert!(text.contains("End Sub"));
}

#[test]
fn cst_preserves_all_whitespace() {
    let code = "Sub  Test (  )\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // Verify that all whitespace is preserved exactly
    assert_eq!(cst.text(), "Sub  Test (  )\n");
}

#[test]
fn syntax_kind_conversions() {
    // Test keyword conversions
    assert_eq!(
        SyntaxKind::from(VB6Token::FunctionKeyword),
        SyntaxKind::FunctionKeyword
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::IfKeyword),
        SyntaxKind::IfKeyword
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::ForKeyword),
        SyntaxKind::ForKeyword
    );

    // Test operators
    assert_eq!(
        SyntaxKind::from(VB6Token::AdditionOperator),
        SyntaxKind::AdditionOperator
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::EqualityOperator),
        SyntaxKind::EqualityOperator
    );

    // Test literals
    assert_eq!(
        SyntaxKind::from(VB6Token::StringLiteral),
        SyntaxKind::StringLiteral
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::Number),
        SyntaxKind::Number
    );
}

#[test]
fn empty_token_stream() {
    let code = "";
    
    let mut source_stream = SourceStream::new("empty.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 0);
    assert_eq!(cst.text(), "");
}


#[test]
fn public_api_does_not_expose_rowan() {
    // This test verifies that we can use the CST API without ever
    // importing or referring to rowan types
    
    let code = "Sub Test()\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    
    // Parse using the public API
    let cst: ConcreteSyntaxTree = parse(token_stream);
    
    // Use all public methods
    let _kind: SyntaxKind = cst.root_kind();
    let _text: String = cst.text();
    let _count: usize = cst.child_count();
    let _debug: String = cst.debug_tree();
    
    // Verify the CST can be cloned (Clone is derived)
    let _cst_clone = cst.clone();
    
    // Verify equality checks work (PartialEq, Eq are derived)
    assert_eq!(cst, _cst_clone);
    
    // If this test compiles and runs without importing rowan,
    // then we've successfully hidden all rowan types!
}

#[test]
fn syntax_kind_public_interface() {
    // Verify SyntaxKind can be used without rowan
    
    let kind = SyntaxKind::from(VB6Token::SubKeyword);
    assert_eq!(kind, SyntaxKind::SubKeyword);
    
    // Verify Debug trait works
    let debug_str = format!("{:?}", kind);
    assert!(debug_str.contains("SubKeyword"));
    
    // Verify comparisons work
    assert_eq!(kind, SyntaxKind::SubKeyword);
    assert_ne!(kind, SyntaxKind::FunctionKeyword);
    
    // Verify Clone works
    let _cloned = kind.clone();
}

#[test]
fn parse_empty_stream() {
    let code = "";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 0);
}

#[test]
fn parse_simple_tokens() {
    let code = "Sub Main()\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 1 child: the SubStatement node
    assert_eq!(cst.child_count(), 1);
    assert!(cst.text().contains("Sub Main()"));
}

#[test]
fn syntax_kind_conversion() {
    assert_eq!(
        SyntaxKind::from(VB6Token::SubKeyword),
        SyntaxKind::SubKeyword
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::Identifier),
        SyntaxKind::Identifier
    );
    assert_eq!(
        SyntaxKind::from(VB6Token::LeftParentheses),
        SyntaxKind::LeftParentheses
    );
}

#[test]
fn parse_attribute_statement() {
    let code = "Attribute VB_Name = \"modTest\"\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1); // AttributeStatement
    assert_eq!(cst.text(), "Attribute VB_Name = \"modTest\"\n");

    // Use navigation methods
    assert!(cst.contains_kind(SyntaxKind::AttributeStatement));
    let attr_statements = cst.find_children_by_kind(SyntaxKind::AttributeStatement);
    assert_eq!(attr_statements.len(), 1);
    assert_eq!(attr_statements[0].text, "Attribute VB_Name = \"modTest\"\n");
}

#[test]
fn parse_option_explicit_on() {
    let code = "Option Explicit On\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Option statements are currently parsed as DimStatement (declarations)
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Option Explicit On\n");
}

#[test]
fn parse_option_explicit_off() {
    let code = "Option Explicit Off\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Option Explicit Off\n");
}

#[test]
fn parse_option_explicit() {
    let code = "Option Explicit\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Option Explicit\n");
}

#[test]
fn parse_single_quote_comment() {
    let code = "' This is a comment\nSub Main()\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 2 children: the comment and the SubStatement
    assert_eq!(cst.child_count(), 3); // 2 statements + EOF
    assert!(cst.text().contains("' This is a comment"));
    assert!(cst.text().contains("Sub Main()"));

    // Use navigation methods
    assert!(cst.contains_kind(SyntaxKind::EndOfLineComment));
    assert!(cst.contains_kind(SyntaxKind::SubStatement));
    
    let first = cst.first_child().unwrap();
    assert_eq!(first.kind, SyntaxKind::EndOfLineComment);
        assert!(first.is_token);
    }

#[test]
fn parse_rem_comment() {
    let code = "REM This is a REM comment\nSub Test()\nEnd Sub\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 2 children: the REM comment and the SubStatement
    assert_eq!(cst.child_count(), 3); // 2 statements + EOF
    assert!(cst.text().contains("REM This is a REM comment"));
    assert!(cst.text().contains("Sub Test()"));

    // Verify REM comment is preserved
    let debug = cst.debug_tree();
    assert!(debug.contains("RemComment"));
}

#[test]
fn parse_mixed_comments() {
    let code = "' Single quote comment\nREM REM comment\nSub Test()\nEnd Sub\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 5 children: EndOfLineComment, Newline, RemComment, Newline, SubStatement
    assert_eq!(cst.child_count(), 5);
    assert!(cst.text().contains("' Single quote comment"));
    assert!(cst.text().contains("REM REM comment"));

    // Use navigation methods
    let children = cst.children();
    assert_eq!(children[0].kind, SyntaxKind::EndOfLineComment);
    assert_eq!(children[1].kind, SyntaxKind::Newline);
    assert_eq!(children[2].kind, SyntaxKind::RemComment);
    assert_eq!(children[3].kind, SyntaxKind::Newline);
    assert_eq!(children[4].kind, SyntaxKind::SubStatement);
    
    assert!(cst.contains_kind(SyntaxKind::EndOfLineComment));
    assert!(cst.contains_kind(SyntaxKind::RemComment));
}

#[test]
fn parse_dim_declaration() {
    let code = "Dim x As Integer\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Dim x As Integer\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
}

#[test]
fn parse_private_declaration() {
    let code = "Private m_value As Long\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Private m_value As Long\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("PrivateKeyword"));
}

#[test]
fn parse_public_declaration() {
    let code = "Public g_config As String\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Public g_config As String\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("PublicKeyword"));
}

#[test]
fn parse_multiple_variable_declaration() {
    let code = "Dim x, y, z As Integer\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Dim x, y, z As Integer\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
}

#[test]
fn parse_const_declaration() {
    let code = "Const MAX_SIZE = 100\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Const MAX_SIZE = 100\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("ConstKeyword"));
}

#[test]
fn parse_private_const_declaration() {
    let code = "Private Const MODULE_NAME = \"MyModule\"\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Private Const MODULE_NAME = \"MyModule\"\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("PrivateKeyword"));
    assert!(debug.contains("ConstKeyword"));
}

#[test]
fn parse_static_declaration() {
    let code = "Static counter As Long\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Static counter As Long\n");

    let debug = cst.debug_tree();
    assert!(debug.contains("DimStatement"));
    assert!(debug.contains("StaticKeyword"));
}

// Navigation method tests

#[test]
fn navigation_children() {
    let code = "Attribute VB_Name\nSub Test()\nEnd Sub\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    let children = cst.children();
    
    assert_eq!(children.len(), 2); // AttributeStatement, SubStatement
    assert_eq!(children[0].kind, SyntaxKind::AttributeStatement);
    assert_eq!(children[1].kind, SyntaxKind::SubStatement);
    assert!(!children[0].is_token);
    assert!(!children[1].is_token);
}

#[test]
fn navigation_find_children_by_kind() {
    let code = "Dim x\nDim y\nSub Test()\nEnd Sub\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    // Find all DimStatements
    let dim_statements = cst.find_children_by_kind(SyntaxKind::DimStatement);
    assert_eq!(dim_statements.len(), 2);
    
    // Find all SubStatements
    let sub_statements = cst.find_children_by_kind(SyntaxKind::SubStatement);
    assert_eq!(sub_statements.len(), 1);
}

#[test]
fn navigation_contains_kind() {
    let code = "Sub Test()\nEnd Sub\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    assert!(cst.contains_kind(SyntaxKind::SubStatement));
    assert!(!cst.contains_kind(SyntaxKind::FunctionStatement));
    assert!(!cst.contains_kind(SyntaxKind::DimStatement));
}

#[test]
fn navigation_first_and_last_child() {
    let code = "Attribute VB_Name\nDim x\nSub Test()\nEnd Sub\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let first = cst.first_child().unwrap();
    assert_eq!(first.kind, SyntaxKind::AttributeStatement);
    assert_eq!(first.text, "Attribute VB_Name\n");
    
    let last = cst.last_child().unwrap();
    assert_eq!(last.kind, SyntaxKind::SubStatement);
}

#[test]
fn navigation_child_at() {
    let code = "Attribute VB_Name\nDim x\nSub Test()\nEnd Sub\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    let first = cst.child_at(0).unwrap();
    assert_eq!(first.kind, SyntaxKind::AttributeStatement);
    
    let second = cst.child_at(1).unwrap();
    assert_eq!(second.kind, SyntaxKind::DimStatement);
    
    let third = cst.child_at(2).unwrap();
    assert_eq!(third.kind, SyntaxKind::SubStatement);
    
    // Fourth is EOF, out of bounds after that
    assert!(cst.child_at(4).is_none());
}
#[test]
fn navigation_empty_tree() {
    let code = "";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    // Even empty code has no children now
    assert_eq!(cst.children().len(), 0);
    assert!(cst.first_child().is_none());
    assert!(cst.last_child().is_none());
    assert!(cst.child_at(0).is_none());
    assert!(!cst.contains_kind(SyntaxKind::SubStatement));
}
#[test]
fn navigation_with_comments_and_whitespace() {
    let code = "' Comment\n\nSub Test()\nEnd Sub\n";
    
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    let children = cst.children();
    
    // Should have 4 children: EndOfLineComment, newline, newline, SubStatement
    assert_eq!(children.len(), 4);
    
    // First is the comment
    assert_eq!(children[0].kind, SyntaxKind::EndOfLineComment);
    assert!(children[0].is_token);
    
    // Second is newline
    assert_eq!(children[1].kind, SyntaxKind::Newline);
    assert!(children[1].is_token);
    
    // Third is the second newline
    assert_eq!(children[2].kind, SyntaxKind::Newline);
    assert!(children[2].is_token);
    
    // Fourth is SubStatement
    assert_eq!(children[3].kind, SyntaxKind::SubStatement);
    assert!(!children[3].is_token);
}

#[test]
fn cst_function_with_modifiers() {
    // Test Public Function
    let code = "Public Function GetValue() As Integer\nEnd Function\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::FunctionStatement);
    }
    assert!(cst.text().contains("Public Function GetValue"));
}

#[test]
fn cst_private_static_function() {
    // Test Private Static Function
    let code = "Private Static Function Calculate(x As Long) As Long\nEnd Function\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::FunctionStatement);
    }
    assert!(cst.text().contains("Private Static Function Calculate"));
}

#[test]
fn cst_friend_function() {
    // Test Friend Function
    let code = "Friend Function ProcessData() As String\nEnd Function\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::FunctionStatement);
    }
    assert!(cst.text().contains("Friend Function ProcessData"));
}

#[test]
fn cst_public_static_sub() {
    // Test Public Static Sub
    let code = "Public Static Sub Initialize()\nEnd Sub\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    assert_eq!(cst.child_count(), 1);
    if let Some(child) = cst.child_at(0) {
        assert_eq!(child.kind, SyntaxKind::SubStatement);
    }
    assert!(cst.text().contains("Public Static Sub Initialize"));
}

#[test]
fn cst_distinguishes_declarations_from_functions() {
    // Test that Private declaration and Private Function are correctly distinguished
    let code = "Private myVar As Integer\nPrivate Function GetVar() As Integer\nEnd Function\n";
    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);
    
    assert_eq!(cst.child_count(), 2);
    
    // First child should be a DimStatement (declaration)
    if let Some(first_child) = cst.child_at(0) {
        assert_eq!(first_child.kind, SyntaxKind::DimStatement);
    }
    
    // Second child should be a FunctionStatement
    if let Some(second_child) = cst.child_at(1) {
        assert_eq!(second_child.kind, SyntaxKind::FunctionStatement);
    }
    
    assert!(cst.text().contains("Private myVar As Integer"));
    assert!(cst.text().contains("Private Function GetVar"));
}

#[test]
fn cst_all_function_modifier_combinations() {
    // Test all valid function/sub modifier combinations
    let test_cases = vec![
        ("Public Function Test() As Integer\nEnd Function\n", SyntaxKind::FunctionStatement),
        ("Private Function Test() As Integer\nEnd Function\n", SyntaxKind::FunctionStatement),
        ("Friend Function Test() As Integer\nEnd Function\n", SyntaxKind::FunctionStatement),
        ("Static Function Test() As Integer\nEnd Function\n", SyntaxKind::FunctionStatement),
        ("Public Static Function Test() As Integer\nEnd Function\n", SyntaxKind::FunctionStatement),
        ("Private Static Function Test() As Integer\nEnd Function\n", SyntaxKind::FunctionStatement),
        ("Friend Static Function Test() As Integer\nEnd Function\n", SyntaxKind::FunctionStatement),
        ("Public Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Private Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Friend Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Static Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Public Static Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Private Static Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
        ("Friend Static Sub Test()\nEnd Sub\n", SyntaxKind::SubStatement),
    ];

    for (code, expected_kind) in test_cases {
        let mut source_stream = SourceStream::new("test.bas", code);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);
        
        assert_eq!(cst.child_count(), 1, "Code: {}", code);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, expected_kind, "Code: {}", code);
        }
    }
}
