//! Example demonstrating CST navigation methods
//!
//! Shows how to use the various navigation APIs to traverse and query
//! the Concrete Syntax Tree.
//!

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

    println!("=== CST Navigation - Depth-First Traversal Example ===\n");

    println!("Source code:\n{source}\n");

    println!("==============================================\n");

    println!("Depth-First Traversal (first 15 non-trivia nodes):");
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
            "   {} {prefix} {:?} -> {text_preview:?}",
            i + 1,
            node.kind()
        );
    }
}
