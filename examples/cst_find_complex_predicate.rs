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

    println!("=== CST Navigation - Complex Predicate Example ===\n");

    println!("Source code:\n{source}\n");

    println!("==============================================\n");

    println!("Complex Predicate Example:");
    println!("   Find all 'As' keywords in type declarations:");
    let as_in_types = root.find_all_if(|n| {
        n.kind() == SyntaxKind::AsKeyword && n.text().trim().eq_ignore_ascii_case("as")
    });
    println!("   Found {} 'As' keywords", as_in_types.len());
}
