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

    println!("=== CST Navigation - Complex Find Example ===\n");

    println!("Source code:\n{source}\n");

    println!("==============================================\n");

    println!("Find All Keywords:");
    let keywords = root.find_all_if(|n| n.is_token() && n.kind().to_string().ends_with("Keyword"));
    println!("   Found {} keywords:", keywords.len());
    for kw in keywords.iter().take(10) {
        println!("      - {} ({:?})", kw.text().trim(), kw.kind());
    }
}
