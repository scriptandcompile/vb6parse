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

    println!("=== CST Navigation - First Significant Child Example ===\n");

    println!("Source code:\n{source}\n");

    println!("==============================================\n");

    println!("First Significant Child of Root:");
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
