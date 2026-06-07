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

    println!("=== CST Navigation - Filter Trivia Example ===\n");

    println!("Source code:\n{source}\n");

    println!("==============================================\n");

    println!("Non-Trivia Structural Nodes of First 5 nodes:");
    let significant_nodes = root.find_all_if(|n| n.is_significant() && !n.is_token());
    println!("   Count: {}", significant_nodes.len());
    for node in significant_nodes.iter().take(5) {
        println!("   - {:?}", node.kind());
    }
}
