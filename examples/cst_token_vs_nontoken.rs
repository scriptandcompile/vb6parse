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

    println!("=== CST Navigation - Token vs Non-Token Example ===\n");

    println!("Source code:\n{source}\n");

    println!("==============================================\n");

    println!("Root-Level Children Analysis:");
    let all_children = root.child_count();
    let non_token_count = root.non_token_children().count();
    let token_count = root.token_children().count();
    println!("   Total children: {all_children}");
    println!("   Structural nodes: {non_token_count}");
    println!("   Tokens: {token_count}");
}
