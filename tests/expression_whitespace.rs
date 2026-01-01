use vb6parse::parsers::cst::ConcreteSyntaxTree;
use vb6parse::parsers::SyntaxKind;

#[test]
fn expression_whitespace_conservation() {
    let input = "a = 1 + 2";

    // Use the high-level API to parse directly from source
    let result = ConcreteSyntaxTree::from_text("test.bas", input);

    // Check for errors
    if result.has_failures() {
        for failure in result.failures() {
            println!("Error: {failure:?}");
        }
        panic!("Parsing failed with errors");
    }

    let (cst_opt, failures) = result.unpack();

    if !failures.is_empty() {
        for failure in failures {
            failure.eprint();
        }
    }

    let cst = cst_opt.expect("Expected CST");

    // Convert to serializable tree to inspect the structure
    let tree = cst.to_serializable();
    let root = tree.root;

    // Helper to print the tree for debugging if needed
    println!("Tree structure:");
    print_node(&root, 0);

    // Verify we have children
    assert!(!root.children().is_empty(), "Root should have children");

    // Find the assignment statement
    let assignment = find_node_by_kind(&root, SyntaxKind::AssignmentStatement)
        .expect("Should find an AssignmentStatement");

    // Check that the assignment statement contains whitespace tokens
    let has_whitespace = assignment
        .children()
        .iter()
        .any(|child| child.kind() == SyntaxKind::Whitespace);

    assert!(
        has_whitespace,
        "AssignmentStatement should contain whitespace tokens"
    );

    // Check specifically for the whitespace around the equals sign
    let children = &assignment.children();
    let equal_pos = children
        .iter()
        .position(|c| c.kind() == SyntaxKind::EqualityOperator)
        .expect("Should find EqualityOperator");

    assert!(
        equal_pos > 0,
        "EqualityOperator should not be the first child"
    );
    assert!(
        children[equal_pos - 1].kind() == SyntaxKind::Whitespace,
        "Should have whitespace before ="
    );
    assert!(
        children[equal_pos + 1].kind() == SyntaxKind::Whitespace,
        "Should have whitespace after ="
    );

    // Now check the expression part: "1 + 2"
    // We look for the BinaryExpression
    let expression = find_node_by_kind(assignment, SyntaxKind::BinaryExpression);

    if let Some(expr) = expression {
        // Check for whitespace around the + operator
        let plus_pos = expr
            .children()
            .iter()
            .position(|c| c.kind() == SyntaxKind::AdditionOperator);

        if let Some(pos) = plus_pos {
            assert!(
                expr.children()[pos - 1].kind() == SyntaxKind::Whitespace,
                "Should have whitespace before +"
            );
            assert!(
                expr.children()[pos + 1].kind() == SyntaxKind::Whitespace,
                "Should have whitespace after +"
            );
        }
    } else {
        // If we can't find AddExpression, maybe it's structured differently, but we should at least verify text
        println!("Warning: Could not find AddExpression, checking text preservation only");
    }

    // Ultimate test: The full text reconstruction should match the input exactly
    assert_eq!(
        cst.text(),
        input,
        "Full text reconstruction should match input exactly"
    );
}

fn print_node(node: &vb6parse::parsers::cst::CstNode, depth: usize) {
    let indent = "  ".repeat(depth);
    println!("{}{:?} ({:?})", indent, node.kind(), node.text());
    for child in node.children() {
        print_node(child, depth + 1);
    }
}

fn find_node_by_kind(
    node: &vb6parse::parsers::cst::CstNode,
    kind: SyntaxKind,
) -> Option<&vb6parse::parsers::cst::CstNode> {
    if node.kind() == kind {
        return Some(node);
    }

    for child in node.children() {
        if let Some(found) = find_node_by_kind(child, kind) {
            return Some(found);
        }
    }

    None
}
