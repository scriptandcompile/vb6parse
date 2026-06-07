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

    println!("=== CST Navigation - All Identifiers Example ===\n");

    println!("Source code:\n{source}\n");

    println!("==============================================\n");

    println!("All Identifiers:");
    let identifiers = root.find_all(SyntaxKind::Identifier);
    for (i, id) in identifiers.iter().enumerate() {
        if !id.text().trim().is_empty() {
            println!("   {}: {}", i + 1, id.text().trim());
        }
    }
}
