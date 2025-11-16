//! Integration test for CST parsing functionality

use vb6parse::language::VB6Token;
use vb6parse::parsers::{parse, SyntaxKind, ConcreteSyntaxTree};
use vb6parse::tokenstream::TokenStream;

#[test]
fn cst_basic_parsing() {
    let tokens = vec![
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("Test", VB6Token::Identifier),
        ("(", VB6Token::LeftParentheses),
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
    ];

    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Now has 1 child: the SubStatement node
    assert_eq!(cst.child_count(), 1);
    assert_eq!(cst.text(), "Sub Test()\n");
}

#[test]
fn cst_with_comments() {
    let tokens = vec![
        ("' This is a comment\n", VB6Token::EndOfLineComment),
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("Main", VB6Token::Identifier),
        ("(", VB6Token::LeftParentheses),
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
    ];

    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    let cst = parse(token_stream);

    // Now has 2 children: the comment (consumed at top level) and the SubStatement
    assert_eq!(cst.child_count(), 2);
    assert!(cst.text().contains("' This is a comment"));
    assert!(cst.text().contains("Sub Main()"));
}

#[test]
fn cst_complex_code() {
    let tokens = vec![
        ("Private", VB6Token::PrivateKeyword),
        (" ", VB6Token::Whitespace),
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("Calculate", VB6Token::Identifier),
        ("(", VB6Token::LeftParentheses),
        ("ByVal", VB6Token::ByValKeyword),
        (" ", VB6Token::Whitespace),
        ("x", VB6Token::Identifier),
        (" ", VB6Token::Whitespace),
        ("As", VB6Token::AsKeyword),
        (" ", VB6Token::Whitespace),
        ("Integer", VB6Token::IntegerKeyword),
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
        ("    ", VB6Token::Whitespace),
        ("Dim", VB6Token::DimKeyword),
        (" ", VB6Token::Whitespace),
        ("result", VB6Token::Identifier),
        (" ", VB6Token::Whitespace),
        ("As", VB6Token::AsKeyword),
        (" ", VB6Token::Whitespace),
        ("Long", VB6Token::LongKeyword),
        ("\n", VB6Token::Newline),
        ("End", VB6Token::EndKeyword),
        (" ", VB6Token::Whitespace),
        ("Sub", VB6Token::SubKeyword),
        ("\n", VB6Token::Newline),
    ];

    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    let cst = parse(token_stream);

    let text = cst.text();
    assert!(text.contains("Private Sub Calculate"));
    assert!(text.contains("ByVal x As Integer"));
    assert!(text.contains("Dim result As Long"));
    assert!(text.contains("End Sub"));
}

#[test]
fn cst_preserves_all_whitespace() {
    let tokens = vec![
        ("Sub", VB6Token::SubKeyword),
        ("  ", VB6Token::Whitespace), // Two spaces
        ("Test", VB6Token::Identifier),
        (" ", VB6Token::Whitespace),  // One space
        ("(", VB6Token::LeftParentheses),
        ("  ", VB6Token::Whitespace), // Two spaces
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
    ];

    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
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
    let tokens = vec![];
    let token_stream = TokenStream::new("empty.bas".to_string(), tokens);
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    assert_eq!(cst.child_count(), 0);
    assert_eq!(cst.text(), "");
}


#[test]
fn public_api_does_not_expose_rowan() {
    // This test verifies that we can use the CST API without ever
    // importing or referring to rowan types
    
    let tokens = vec![
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("Test", VB6Token::Identifier),
        ("(", VB6Token::LeftParentheses),
        (")", VB6Token::RightParentheses),
        ("\n", VB6Token::Newline),
    ];
    
    let token_stream = TokenStream::new("test.bas".to_string(), tokens);
    
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
        let tokens = TokenStream::new("test.bas".to_string(), vec![]);
        let cst = parse(tokens);
        
        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 0);
    }
    
    #[test]
    fn parse_simple_tokens() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Main", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
            ],
        );
        
        let cst = parse(tokens);
        
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
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Attribute", VB6Token::AttributeKeyword),
                (" ", VB6Token::Whitespace),
                ("VB_Name", VB6Token::Identifier),
                (" ", VB6Token::Whitespace),
                ("=", VB6Token::EqualityOperator),
                (" ", VB6Token::Whitespace),
                ("\"modTest\"", VB6Token::StringLiteral),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Attribute VB_Name = \"modTest\"\n");

        // Use navigation methods
        assert!(cst.contains_kind(SyntaxKind::AttributeStatement));
        let attr_statements = cst.find_children_by_kind(SyntaxKind::AttributeStatement);
        assert_eq!(attr_statements.len(), 1);
        assert_eq!(attr_statements[0].text, "Attribute VB_Name = \"modTest\"\n");
    }

    #[test]
    fn parse_option_explicit_on() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Option", VB6Token::OptionKeyword),
                (" ", VB6Token::Whitespace),
                ("Explicit", VB6Token::ExplicitKeyword),
                (" ", VB6Token::Whitespace),
                ("On", VB6Token::OnKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        // Option statements are currently parsed as DimStatement (declarations)
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit On\n");
    }

    #[test]
    fn parse_option_explicit_off() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Option", VB6Token::OptionKeyword),
                (" ", VB6Token::Whitespace),
                ("Explicit", VB6Token::ExplicitKeyword),
                (" ", VB6Token::Whitespace),
                ("Off", VB6Token::OffKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit Off\n");
    }

    #[test]
    fn parse_option_explicit() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Option", VB6Token::OptionKeyword),
                (" ", VB6Token::Whitespace),
                ("Explicit", VB6Token::ExplicitKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit\n");
    }

    #[test]
    fn parse_single_quote_comment() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("' This is a comment\n", VB6Token::EndOfLineComment),
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        // Should have 2 children: the comment and the SubStatement
        assert_eq!(cst.child_count(), 2);
        assert!(cst.text().contains("' This is a comment"));
        assert!(cst.text().contains("Sub Test()"));

        // Use navigation methods
        assert!(cst.contains_kind(SyntaxKind::Comment));
        assert!(cst.contains_kind(SyntaxKind::SubStatement));
        
        let first = cst.first_child().unwrap();
        assert_eq!(first.kind, SyntaxKind::Comment);
        assert!(first.is_token);
    }

    #[test]
    fn parse_rem_comment() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("REM This is a REM comment\n", VB6Token::RemComment),
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        // Should have 2 children: the REM comment and the SubStatement
        assert_eq!(cst.child_count(), 2);
        assert!(cst.text().contains("REM This is a REM comment"));
        assert!(cst.text().contains("Sub Test()"));

        // Verify REM comment is preserved
        let debug = cst.debug_tree();
        assert!(debug.contains("RemComment"));
    }

    #[test]
    fn parse_mixed_comments() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("' Single quote comment\n", VB6Token::EndOfLineComment),
                ("REM REM comment\n", VB6Token::RemComment),
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        // Should have 3 children: EndOfLineComment, RemComment, and SubStatement
        assert_eq!(cst.child_count(), 3);
        assert!(cst.text().contains("' Single quote comment"));
        assert!(cst.text().contains("REM REM comment"));

        // Use navigation methods
        let children = cst.children();
        assert_eq!(children[0].kind, SyntaxKind::Comment);
        assert_eq!(children[1].kind, SyntaxKind::RemComment);
        assert_eq!(children[2].kind, SyntaxKind::SubStatement);
        
        assert!(cst.contains_kind(SyntaxKind::Comment));
        assert!(cst.contains_kind(SyntaxKind::RemComment));
    }

    #[test]
    fn parse_dim_declaration() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Dim", VB6Token::DimKeyword),
                (" ", VB6Token::Whitespace),
                ("x", VB6Token::Identifier),
                (" ", VB6Token::Whitespace),
                ("As", VB6Token::AsKeyword),
                (" ", VB6Token::Whitespace),
                ("Integer", VB6Token::IntegerKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Dim x As Integer\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn parse_private_declaration() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Private", VB6Token::PrivateKeyword),
                (" ", VB6Token::Whitespace),
                ("m_value", VB6Token::Identifier),
                (" ", VB6Token::Whitespace),
                ("As", VB6Token::AsKeyword),
                (" ", VB6Token::Whitespace),
                ("Long", VB6Token::LongKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Private m_value As Long\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("PrivateKeyword"));
    }

    #[test]
    fn parse_public_declaration() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Public", VB6Token::PublicKeyword),
                (" ", VB6Token::Whitespace),
                ("g_config", VB6Token::Identifier),
                (" ", VB6Token::Whitespace),
                ("As", VB6Token::AsKeyword),
                (" ", VB6Token::Whitespace),
                ("String", VB6Token::StringKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Public g_config As String\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("PublicKeyword"));
    }

    #[test]
    fn parse_multiple_variable_declaration() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Dim", VB6Token::DimKeyword),
                (" ", VB6Token::Whitespace),
                ("x", VB6Token::Identifier),
                (",", VB6Token::Comma),
                (" ", VB6Token::Whitespace),
                ("y", VB6Token::Identifier),
                (",", VB6Token::Comma),
                (" ", VB6Token::Whitespace),
                ("z", VB6Token::Identifier),
                (" ", VB6Token::Whitespace),
                ("As", VB6Token::AsKeyword),
                (" ", VB6Token::Whitespace),
                ("Integer", VB6Token::IntegerKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Dim x, y, z As Integer\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn parse_const_declaration() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Const", VB6Token::ConstKeyword),
                (" ", VB6Token::Whitespace),
                ("MAX_SIZE", VB6Token::Identifier),
                (" ", VB6Token::Whitespace),
                ("=", VB6Token::EqualityOperator),
                (" ", VB6Token::Whitespace),
                ("100", VB6Token::Number),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Const MAX_SIZE = 100\n");

        let debug = cst.debug_tree();
        assert!(debug.contains("DimStatement"));
        assert!(debug.contains("ConstKeyword"));
    }

    #[test]
    fn parse_private_const_declaration() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Private", VB6Token::PrivateKeyword),
                (" ", VB6Token::Whitespace),
                ("Const", VB6Token::ConstKeyword),
                (" ", VB6Token::Whitespace),
                ("MODULE_NAME", VB6Token::Identifier),
                (" ", VB6Token::Whitespace),
                ("=", VB6Token::EqualityOperator),
                (" ", VB6Token::Whitespace),
                ("\"MyModule\"", VB6Token::StringLiteral),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

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
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Static", VB6Token::StaticKeyword),
                (" ", VB6Token::Whitespace),
                ("counter", VB6Token::Identifier),
                (" ", VB6Token::Whitespace),
                ("As", VB6Token::AsKeyword),
                (" ", VB6Token::Whitespace),
                ("Long", VB6Token::LongKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);

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
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Attribute", VB6Token::AttributeKeyword),
                (" ", VB6Token::Whitespace),
                ("VB_Name", VB6Token::Identifier),
                ("\n", VB6Token::Newline),
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);
        let children = cst.children();
        
        assert_eq!(children.len(), 2);
        assert_eq!(children[0].kind, SyntaxKind::AttributeStatement);
        assert_eq!(children[1].kind, SyntaxKind::SubStatement);
        assert!(!children[0].is_token);
        assert!(!children[1].is_token);
    }

    #[test]
    fn navigation_find_children_by_kind() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Dim", VB6Token::DimKeyword),
                (" ", VB6Token::Whitespace),
                ("x", VB6Token::Identifier),
                ("\n", VB6Token::Newline),
                ("Dim", VB6Token::DimKeyword),
                (" ", VB6Token::Whitespace),
                ("y", VB6Token::Identifier),
                ("\n", VB6Token::Newline),
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);
        
        // Find all DimStatements
        let dim_statements = cst.find_children_by_kind(SyntaxKind::DimStatement);
        assert_eq!(dim_statements.len(), 2);
        
        // Find all SubStatements
        let sub_statements = cst.find_children_by_kind(SyntaxKind::SubStatement);
        assert_eq!(sub_statements.len(), 1);
    }

    #[test]
    fn navigation_contains_kind() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);
        
        assert!(cst.contains_kind(SyntaxKind::SubStatement));
        assert!(!cst.contains_kind(SyntaxKind::FunctionStatement));
        assert!(!cst.contains_kind(SyntaxKind::DimStatement));
    }

    #[test]
    fn navigation_first_and_last_child() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Attribute", VB6Token::AttributeKeyword),
                (" ", VB6Token::Whitespace),
                ("VB_Name", VB6Token::Identifier),
                ("\n", VB6Token::Newline),
                ("Dim", VB6Token::DimKeyword),
                (" ", VB6Token::Whitespace),
                ("x", VB6Token::Identifier),
                ("\n", VB6Token::Newline),
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);
        
        let first = cst.first_child().unwrap();
        assert_eq!(first.kind, SyntaxKind::AttributeStatement);
        assert_eq!(first.text, "Attribute VB_Name\n");
        
        let last = cst.last_child().unwrap();
        assert_eq!(last.kind, SyntaxKind::SubStatement);
    }

    #[test]
    fn navigation_child_at() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("Attribute", VB6Token::AttributeKeyword),
                (" ", VB6Token::Whitespace),
                ("VB_Name", VB6Token::Identifier),
                ("\n", VB6Token::Newline),
                ("Dim", VB6Token::DimKeyword),
                (" ", VB6Token::Whitespace),
                ("x", VB6Token::Identifier),
                ("\n", VB6Token::Newline),
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);
        
        let first = cst.child_at(0).unwrap();
        assert_eq!(first.kind, SyntaxKind::AttributeStatement);
        
        let second = cst.child_at(1).unwrap();
        assert_eq!(second.kind, SyntaxKind::DimStatement);
        
        let third = cst.child_at(2).unwrap();
        assert_eq!(third.kind, SyntaxKind::SubStatement);
        
        // Out of bounds should return None
        assert!(cst.child_at(3).is_none());
    }

    #[test]
    fn navigation_empty_tree() {
        let tokens = TokenStream::new("test.bas".to_string(), vec![]);
        let cst = parse(tokens);
        
        assert_eq!(cst.children().len(), 0);
        assert!(cst.first_child().is_none());
        assert!(cst.last_child().is_none());
        assert!(cst.child_at(0).is_none());
        assert!(!cst.contains_kind(SyntaxKind::SubStatement));
    }

    #[test]
    fn navigation_with_comments_and_whitespace() {
        let tokens = TokenStream::new(
            "test.bas".to_string(),
            vec![
                ("' Comment\n", VB6Token::EndOfLineComment),
                ("\n", VB6Token::Newline),
                ("Sub", VB6Token::SubKeyword),
                (" ", VB6Token::Whitespace),
                ("Test", VB6Token::Identifier),
                ("(", VB6Token::LeftParentheses),
                (")", VB6Token::RightParentheses),
                ("\n", VB6Token::Newline),
                ("End", VB6Token::EndKeyword),
                (" ", VB6Token::Whitespace),
                ("Sub", VB6Token::SubKeyword),
                ("\n", VB6Token::Newline),
            ],
        );

        let cst = parse(tokens);
        let children = cst.children();
        
        // Should have 3 children: comment, newline, and SubStatement
        assert_eq!(children.len(), 3);
        
        // First is the comment
        assert_eq!(children[0].kind, SyntaxKind::Comment);
        assert!(children[0].is_token);
        
        // Second is newline
        assert_eq!(children[1].kind, SyntaxKind::Newline);
        assert!(children[1].is_token);
        
        // Third is SubStatement
        assert_eq!(children[2].kind, SyntaxKind::SubStatement);
        assert!(!children[2].is_token);
    }