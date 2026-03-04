/// Example demonstrating the visitor pattern for CST analysis
///
/// This example shows how to use the Visitor trait to traverse and analyze
/// a VB6 Concrete Syntax Tree without writing complex pattern matching code.
use vb6parse::parsers::cst::{CstNode, Visitor};
use vb6parse::parsers::SyntaxKind;
use vb6parse::ConcreteSyntaxTree;

/// A visitor that collects statistics about a VB6 module
struct StatisticsVisitor {
    sub_count: usize,
    function_count: usize,
    dim_count: usize,
    identifier_count: usize,
    sub_names: Vec<String>,
    function_names: Vec<String>,
}

impl StatisticsVisitor {
    fn new() -> Self {
        Self {
            sub_count: 0,
            function_count: 0,
            dim_count: 0,
            identifier_count: 0,
            sub_names: Vec::new(),
            function_names: Vec::new(),
        }
    }

    fn print_report(&self) {
        println!("=== VB6 Module Statistics ===");
        println!("Subroutines: {}", self.sub_count);
        if !self.sub_names.is_empty() {
            println!("  Names: {}", self.sub_names.join(", "));
        }
        println!("Functions: {}", self.function_count);
        if !self.function_names.is_empty() {
            println!("  Names: {}", self.function_names.join(", "));
        }
        println!("Dim statements: {}", self.dim_count);
        println!("Identifiers: {}", self.identifier_count);
        println!();
    }
}

impl Visitor for StatisticsVisitor {
    // Override specific node type handlers to collect statistics

    fn visit_sub_statement(&mut self, node: &CstNode) {
        self.sub_count += 1;

        // Extract the sub name (first identifier after Sub keyword)
        if let Some(id) = node.find(SyntaxKind::Identifier) {
            self.sub_names.push(id.text().to_string());
        }

        // Note: We don't need to call walk_node here - the default visit_node
        // implementation already handles recursion for us.
    }

    fn visit_function_statement(&mut self, node: &CstNode) {
        self.function_count += 1;

        // Extract the function name
        if let Some(id) = node.find(SyntaxKind::Identifier) {
            self.function_names.push(id.text().to_string());
        }
    }

    fn visit_dim_statement(&mut self, _node: &CstNode) {
        self.dim_count += 1;
    }

    fn visit_identifier(&mut self, _node: &CstNode) {
        self.identifier_count += 1;
    }
}

/// A visitor that prints the structure of a VB6 module
struct StructurePrinter {
    depth: usize,
}

impl StructurePrinter {
    fn new() -> Self {
        Self { depth: 0 }
    }

    fn indent(&self) -> String {
        "  ".repeat(self.depth)
    }
}

impl Visitor for StructurePrinter {
    // Override visit_node to control when we print and recurse
    fn visit_node(&mut self, node: &CstNode) {
        // Only print interesting nodes
        match node.kind() {
            SyntaxKind::SubStatement
            | SyntaxKind::FunctionStatement
            | SyntaxKind::DimStatement
            | SyntaxKind::IfStatement
            | SyntaxKind::ForStatement
            | SyntaxKind::WhileStatement
            | SyntaxKind::AssignmentStatement => {
                println!("{}{:?}", self.indent(), node.kind());
                self.depth += 1;

                // Manually traverse children to maintain depth tracking
                for child in node.children() {
                    self.visit_node(&child);
                }

                self.depth -= 1;
                return; // Don't call default walk_node
            }
            _ => {}
        }

        // For other nodes, use default traversal
        vb6parse::parsers::cst::walk_node(self, node);
    }
}

fn main() {
    // Example VB6 code
    let source = r#"
Option Explicit

Public Sub Initialize()
    Dim x As Integer
    Dim message As String
    x = 42
    message = "Hello"
End Sub

Public Function Calculate(value As Integer) As Integer
    Dim result As Integer
    result = value * 2
    Calculate = result
End Function

Private Sub Helper()
    Dim i As Integer
    For i = 1 To 10
        ' Do something
    Next i
End Sub
"#;

    // Parse the source code
    let (cst_opt, _errors) = ConcreteSyntaxTree::from_text("example.bas", source).unpack();

    if let Some(cst) = cst_opt {
        let root = cst.to_root_node();

        // Example 1: Collect statistics
        println!("Example 1: Collecting Statistics");
        println!("{}", "-".repeat(40));
        let mut stats = StatisticsVisitor::new();
        stats.visit_node(&root);
        stats.print_report();

        // Example 2: Print module structure
        println!("Example 2: Module Structure");
        println!("{}", "-".repeat(40));
        let mut printer = StructurePrinter::new();
        printer.visit_node(&root);
        println!();
    } else {
        eprintln!("Failed to parse source code");
    }
}
