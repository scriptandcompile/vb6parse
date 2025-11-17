use vb6parse::parsers::cst::parse;
use vb6parse::parsers::{SourceStream, SyntaxKind};
use vb6parse::tokenize::tokenize;

#[test]
fn binary_conditional() {
    let code = r#"Sub Test()
    If x = 5 Then
    End If
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // Navigate the tree structure
    let children = cst.children();

    // Find the SubStatement node
    let sub_statement = children
        .iter()
        .find(|child| child.kind == SyntaxKind::SubStatement)
        .expect("Should have a SubStatement node");

    // The SubStatement should contain a CodeBlock
    let code_block = sub_statement
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::CodeBlock)
        .expect("SubStatement should contain a CodeBlock");

    // The CodeBlock should contain an IfStatement
    let if_statement = code_block
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::IfStatement)
        .expect("CodeBlock should contain an IfStatement");

    // The IfStatement should contain a BinaryConditional
    let binary_conditional = if_statement
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::BinaryConditional)
        .expect("IfStatement should contain a BinaryConditional");

    // Verify the BinaryConditional structure
    assert_eq!(binary_conditional.kind, SyntaxKind::BinaryConditional);
    assert!(
        !binary_conditional.is_token,
        "BinaryConditional should be a node, not a token"
    );

    // Verify the BinaryConditional contains the expected elements:
    // whitespace, identifier "x", whitespace, "=", whitespace, number "5", whitespace
    assert!(binary_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::Identifier && c.text == "x"));
    assert!(binary_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::EqualityOperator));
    assert!(binary_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::Number && c.text == "5"));
}

#[test]
fn unary_conditional() {
    let code = r#"Sub Test()
    If Not isEmpty(x) Then
    End If
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // Navigate the tree structure
    let children = cst.children();

    // Find the SubStatement node
    let sub_statement = children
        .iter()
        .find(|child| child.kind == SyntaxKind::SubStatement)
        .expect("Should have a SubStatement node");

    // The SubStatement should contain a CodeBlock
    let code_block = sub_statement
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::CodeBlock)
        .expect("SubStatement should contain a CodeBlock");

    // The CodeBlock should contain an IfStatement
    let if_statement = code_block
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::IfStatement)
        .expect("CodeBlock should contain an IfStatement");

    // The IfStatement should contain a UnaryConditional
    let unary_conditional = if_statement
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::UnaryConditional)
        .expect("IfStatement should contain a UnaryConditional");

    // Verify the UnaryConditional structure
    assert_eq!(unary_conditional.kind, SyntaxKind::UnaryConditional);
    assert!(
        !unary_conditional.is_token,
        "UnaryConditional should be a node, not a token"
    );

    // Verify the UnaryConditional contains the expected elements:
    // whitespace, Not keyword, whitespace, identifier "isEmpty", parentheses, identifier "x", parentheses, whitespace
    assert!(unary_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::NotKeyword));
    assert!(unary_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::Identifier && c.text == "isEmpty"));
    assert!(unary_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::Identifier && c.text == "x"));
    assert!(unary_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::LeftParentheses));
    assert!(unary_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::RightParentheses));
}

#[test]
fn nested_if_elseif_else() {
    let code = r#"Sub Test()
    If x > 0 Then
        If y > 0 Then
        ElseIf y < 0 Then
        Else
        End If
    ElseIf x < 0 Then
    Else
    End If
End Sub
"#;

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    // Navigate the tree structure
    let children = cst.children();

    // Find the SubStatement node
    let sub_statement = children
        .iter()
        .find(|child| child.kind == SyntaxKind::SubStatement)
        .expect("Should have a SubStatement node");

    // The SubStatement should contain a CodeBlock
    let code_block = sub_statement
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::CodeBlock)
        .expect("SubStatement should contain a CodeBlock");

    // Find the outer IfStatement in the CodeBlock
    let outer_if = code_block
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::IfStatement)
        .expect("CodeBlock should contain an outer IfStatement");

    // Verify outer If has a BinaryConditional (x > 0)
    let outer_conditional = outer_if
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::BinaryConditional)
        .expect("Outer IfStatement should contain a BinaryConditional");
    assert!(outer_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::Identifier && c.text == "x"));
    assert!(outer_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::GreaterThanOperator));

    // Find the CodeBlock inside the outer If
    let outer_code_block = outer_if
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::CodeBlock)
        .expect("Outer IfStatement should contain a CodeBlock");

    // Find the inner IfStatement (nested within the outer If's CodeBlock)
    let inner_if = outer_code_block
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::IfStatement)
        .expect("Outer CodeBlock should contain a nested IfStatement");

    // Verify inner If has a BinaryConditional (y > 0)
    let inner_conditional = inner_if
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::BinaryConditional)
        .expect("Inner IfStatement should contain a BinaryConditional");
    assert!(inner_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::Identifier && c.text == "y"));
    assert!(inner_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::GreaterThanOperator));

    // Verify inner If has ElseIf clause
    let inner_elseif = inner_if
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::ElseIfClause)
        .expect("Inner IfStatement should contain an ElseIfClause");

    // Verify inner ElseIf has a BinaryConditional (y < 0)
    let inner_elseif_conditional = inner_elseif
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::BinaryConditional)
        .expect("Inner ElseIfClause should contain a BinaryConditional");
    assert!(inner_elseif_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::Identifier && c.text == "y"));
    assert!(inner_elseif_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::LessThanOperator));

    // Verify inner If has Else clause
    assert!(
        inner_if
            .children
            .iter()
            .any(|child| child.kind == SyntaxKind::ElseClause),
        "Inner IfStatement should contain an ElseClause"
    );

    // Verify outer If has ElseIf clause
    let outer_elseif = outer_if
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::ElseIfClause)
        .expect("Outer IfStatement should contain an ElseIfClause");

    // Verify outer ElseIf has a BinaryConditional (x < 0)
    let outer_elseif_conditional = outer_elseif
        .children
        .iter()
        .find(|child| child.kind == SyntaxKind::BinaryConditional)
        .expect("Outer ElseIfClause should contain a BinaryConditional");
    assert!(outer_elseif_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::Identifier && c.text == "x"));
    assert!(outer_elseif_conditional
        .children
        .iter()
        .any(|c| c.kind == SyntaxKind::LessThanOperator));

    // Verify outer If has Else clause
    assert!(
        outer_if
            .children
            .iter()
            .any(|child| child.kind == SyntaxKind::ElseClause),
        "Outer IfStatement should contain an ElseClause"
    );
}
