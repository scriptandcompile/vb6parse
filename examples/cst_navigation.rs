//! Example demonstrating CST navigation methods
//!
//! Shows how to use the various navigation APIs to traverse and query
//! the Concrete Syntax Tree.
//!

use vb6parse::parsers::cst::CstNode;
use vb6parse::parsers::SyntaxKind;
use vb6parse::ConcreteSyntaxTree;

fn main() {
    let source = r"
' Module for testing
Sub Test()
    Dim x As Integer
    Dim y As String
    x = 42
End Sub

Function Calculate(a As Integer, b As Integer) As Integer
    Calculate = a + b
End Function
";

    let result = ConcreteSyntaxTree::from_text("example.bas", source);
    let cst = result.unwrap_or_fail();
    let root = cst.to_serializable().root;

    println!("=== CST Navigation Examples ===\n");

    // Example 1: Count different statement types
    println!("1. Statement Counts:");
    let sub_count = root.find_all(SyntaxKind::SubStatement).len();
    let func_count = root.find_all(SyntaxKind::FunctionStatement).len();
    let dim_count = root.find_all(SyntaxKind::DimStatement).len();
    println!("   Sub statements: {sub_count}");
    println!("   Function statements: {func_count}");
    println!("   Dim statements: {dim_count}");

    // Example 2: Find all identifiers
    println!("\n2. All Identifiers:");
    let identifiers = root.find_all(SyntaxKind::Identifier);
    for (i, id) in identifiers.iter().enumerate() {
        if !id.text().trim().is_empty() {
            println!("   {}: {}", i + 1, id.text().trim());
        }
    }

    // Example 3: Find first Sub statement and navigate its children
    println!("\n3. First Sub Statement Structure:");
    if let Some(sub_stmt) = root.find(SyntaxKind::SubStatement) {
        println!("   Text: {}", sub_stmt.text().lines().next().unwrap_or(""));
        println!("   Direct children: {}", sub_stmt.child_count());
        println!(
            "   Significant children: {}",
            sub_stmt.significant_children().count()
        );

        // Find Dim statements inside the Sub
        let dims = sub_stmt.find_all(SyntaxKind::DimStatement);
        println!("   Dim statements inside: {}", dims.len());
    }

    // Example 4: Filter out trivia
    println!("\n4. Non-Trivia Structural Nodes:");
    let significant_nodes = root.find_all_if(|n| n.is_significant() && !n.is_token());
    println!("   Count: {}", significant_nodes.len());
    for node in significant_nodes.iter().take(5) {
        println!("   - {:?}", node.kind());
    }

    // Example 5: Depth-first traversal
    println!("\n5. Depth-First Traversal (first 15 non-trivia nodes):");
    for (i, node) in root
        .descendants()
        .filter(|n| n.is_significant())
        .take(15)
        .enumerate()
    {
        let prefix = if node.is_token() {
            "[Token]"
        } else {
            "[Node] "
        };
        let text_preview = node
            .text()
            .chars()
            .take(20)
            .collect::<String>()
            .replace('\n', "\\n");
        println!(
            "   {} {} {:?} -> {:?}",
            i + 1,
            prefix,
            node.kind(),
            text_preview
        );
    }

    // Example 6: Find nodes matching complex criteria
    println!("\n6. Find All Keywords:");
    let keywords = root.find_all_if(|n| n.is_token() && n.kind().to_string().ends_with("Keyword"));
    println!("   Found {} keywords:", keywords.len());
    for kw in keywords.iter().take(10) {
        println!("      - {} ({:?})", kw.text().trim(), kw.kind());
    }

    // Example 7: Token vs Non-Token children
    println!("\n7. Root-Level Children Analysis:");
    let all_children = root.child_count();
    let non_token_count = root.non_token_children().count();
    let token_count = root.token_children().count();
    println!("   Total children: {all_children}");
    println!("   Structural nodes: {non_token_count}");
    println!("   Tokens: {token_count}");

    // Example 8: Comments and trivia
    println!("\n8. Comments Found:");
    let comments = root.find_all_if(CstNode::is_comment);
    for comment in &comments {
        println!("   - {}", comment.text().trim());
    }

    // Example 9: Using predicates for complex queries
    println!("\n9. Complex Predicate Example:");
    println!("   Find all As keywords in type declarations:");
    let as_in_types = root.find_all_if(|n| {
        n.kind() == SyntaxKind::AsKeyword && n.text().trim().eq_ignore_ascii_case("as")
    });
    println!("   Found {} 'As' keywords", as_in_types.len());

    // Example 10: Iterating significant children only
    println!("\n10. First Significant Child of Root:");
    let first_sig_opt = root
        .significant_children()
        .next()
        .map(|n| (n.kind(), n.is_token(), n.child_count()));

    if let Some((kind, is_token, child_count)) = first_sig_opt {
        println!("   Kind: {kind:?}");
        println!("   Is token: {is_token}");
        println!("   Child count: {child_count}");
    }
}
