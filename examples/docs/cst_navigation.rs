use vb6parse::parsers::SyntaxKind;
use vb6parse::*;

fn main() {
    let code = r#"Sub Calculate()
    Dim result As Double
    result = 10 * 5 + 3
    MsgBox result
End Sub"#;

    let cst = ConcreteSyntaxTree::from_text("test.bas", code).unwrap();
    let root = cst.to_serializable().root;

    println!("Total nodes in tree: {}", root.descendants().count());

    // Find all Dim statements
    let dims = root.find_all(SyntaxKind::DimStatement);
    println!("Found {} Dim statements", dims.len());

    // Find all identifiers
    let identifiers = root.find_all_if(|n| n.kind() == SyntaxKind::Identifier);
    println!("Found {} identifiers", identifiers.len());
    for id in identifiers {
        println!("  - {}", id.text());
    }

    // Navigate to specific nodes
    if let Some(sub_stmt) = root.find(SyntaxKind::SubStatement) {
        println!("\nSubroutine found:");
        println!("  Text:\n'{}'", sub_stmt.text());
        println!("  Children: {}", sub_stmt.child_count());
    }
}
