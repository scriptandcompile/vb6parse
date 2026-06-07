//! Example demonstrating CST navigation methods
//!
//! Shows how to use the various navigation APIs to traverse and query
//! the Concrete Syntax Tree.
//!

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

    println!("=== CST Navigation - First Sub Structure Example ===\n");

    println!("Source code:\n{source}\n");

    println!("==============================================\n");

    println!("First Sub Statement Structure:");
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
}
