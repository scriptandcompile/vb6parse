//! Visitor pattern for traversing and analyzing the Concrete Syntax Tree.
//!
//! This module provides visitor traits that enable type-safe, composable traversal
//! of the CST for complex analyses like semantic analysis, code generation, and linting.
//!
//! # Overview
//!
//! The visitor pattern separates tree traversal logic from analysis logic, making it easier
//! to implement complex multi-pass analyses without deeply nested match statements.
//!
//! Two visitor traits are provided:
//! - [`Visitor`] - Immutable visitor for read-only analysis
//! - [`VisitorMut`] - Mutable visitor for tree transformations
//!
//! # Examples
//!
//! ## Simple Analysis Visitor
//!
//! ```rust
//! use vb6parse::ConcreteSyntaxTree;
//! use vb6parse::parsers::cst::{CstNode, Visitor};
//! use vb6parse::parsers::SyntaxKind;
//!
//! struct IdentifierCollector {
//!     identifiers: Vec<String>,
//! }
//!
//! impl Visitor for IdentifierCollector {
//!     fn visit_identifier(&mut self, node: &CstNode) {
//!         self.identifiers.push(node.text().to_string());
//!     }
//! }
//!
//! let source = "Sub Test()\nDim x As Integer\nEnd Sub";
//! let (cst_opt, _) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
//! let cst = cst_opt.expect("Failed to parse");
//! let root = cst.to_root_node();
//!
//! let mut collector = IdentifierCollector { identifiers: Vec::new() };
//! collector.visit_node(&root);
//!
//! assert!(collector.identifiers.len() > 0);
//! ```
//!
//! ## Multi-Pass Analysis
//!
//! ```rust
//! use vb6parse::ConcreteSyntaxTree;
//! use vb6parse::parsers::cst::{CstNode, Visitor};
//! use vb6parse::parsers::SyntaxKind;
//!
//! // First pass: collect declarations
//! struct DeclarationFinder {
//!     declarations: Vec<String>,
//! }
//!
//! impl Visitor for DeclarationFinder {
//!     fn visit_dim_statement(&mut self, node: &CstNode) {
//!         self.declarations.push(node.text().to_string());
//!     }
//! }
//!
//! // Second pass: count references
//! struct ReferenceCounter {
//!     count: usize,
//! }
//!
//! impl Visitor for ReferenceCounter {
//!     fn visit_identifier(&mut self, node: &CstNode) {
//!         self.count += 1;
//!     }
//! }
//!
//! let source = "Sub Test()\nDim x As Integer\nx = 42\nEnd Sub";
//! let (cst_opt, _) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
//! let cst = cst_opt.expect("Failed to parse");
//! let root = cst.to_root_node();
//!
//! // Run multiple visitors
//! let mut decl_finder = DeclarationFinder { declarations: Vec::new() };
//! decl_finder.visit_node(&root);
//!
//! let mut ref_counter = ReferenceCounter { count: 0 };
//! ref_counter.visit_node(&root);
//! ```
//!
//! # Design Rationale
//!
//! - **Lightweight**: Visitors are optional; existing navigation APIs remain available
//! - **Type-safe**: Each node type gets its own visit method
//! - **Composable**: Multiple visitors can process the same tree
//! - **Flexible**: Override only the node types you care about
//! - **Default traversal**: `walk_node` provides standard depth-first traversal
//!
//! # Traversal Pattern and Common Pitfalls
//!
//! **IMPORTANT**: The `walk_node` function automatically visits children after calling
//! type-specific visitor methods. This means:
//!
//! ❌ **DON'T** call `walk_node(self, node)` inside specific visitor method overrides:
//! ```rust,ignore
//! impl Visitor for MyVisitor {
//!     fn visit_sub_statement(&mut self, node: &CstNode) {
//!         // Process the node...
//!         walk_node(self, node);  // ❌ CAUSES INFINITE RECURSION!
//!     }
//! }
//! ```
//!
//! ✅ **DO** let the default traversal handle children:
//! ```rust,ignore
//! impl Visitor for MyVisitor {
//!     fn visit_sub_statement(&mut self, node: &CstNode) {
//!         // Process the node...
//!         // Children are automatically visited after this returns
//!     }
//! }
//! ```
//!
//! ✅ **DO** override `visit_node` if you need custom traversal control:
//! ```rust,ignore
//! impl Visitor for MyVisitor {
//!     fn visit_node(&mut self, node: &CstNode) {
//!         // Pre-processing
//!         println!("Entering: {:?}", node.kind());
//!         
//!         // Standard traversal
//!         walk_node(self, node);
//!         
//!         // Post-processing
//!         println!("Leaving: {:?}", node.kind());
//!     }
//! }
//! ```

use super::CstNode;
use crate::parsers::SyntaxKind;

/// Immutable visitor for read-only CST traversal and analysis.
///
/// Implement this trait to perform analysis passes over the CST without modifying it.
/// Override only the `visit_*` methods for node types you're interested in.
///
/// The default implementation of `visit_node` calls `walk_node`, which performs
/// depth-first traversal and dispatches to type-specific visit methods.
///
/// # Examples
///
/// ```rust
/// use vb6parse::parsers::cst::{CstNode, Visitor};
///
/// struct SubCounter {
///     count: usize,
/// }
///
/// impl Visitor for SubCounter {
///     fn visit_sub_statement(&mut self, _node: &CstNode) {
///         self.count += 1;
///         // Note: Don't call walk_node here! The default visit_node
///         // implementation already handles child traversal after this returns.
///     }
/// }
/// # use vb6parse::ConcreteSyntaxTree;
/// # let source = "Sub A()\nEnd Sub\nSub B()\nEnd Sub";
/// # let (cst_opt, _) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
/// # let cst = cst_opt.expect("Failed to parse");
/// # let root = cst.to_root_node();
/// # let mut counter = SubCounter { count: 0 };
/// # counter.visit_node(&root);
/// # assert_eq!(counter.count, 2);
/// ```
pub trait Visitor: Sized {
    /// Visit a node in the CST.
    ///
    /// The default implementation calls `walk_node` to perform depth-first traversal.
    /// Override this to change the traversal strategy.
    fn visit_node(&mut self, node: &CstNode) {
        walk_node(self, node);
    }

    // Statement visitors

    /// Visit a `Dim` statement node.
    fn visit_dim_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Const` statement node.
    fn visit_const_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Sub` statement node.
    fn visit_sub_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Function` statement node.
    fn visit_function_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Property` statement node.
    fn visit_property_statement(&mut self, _node: &CstNode) {}
    /// Visit an `If` statement node.
    fn visit_if_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Select` Case statement node.
    fn visit_select_case_statement(&mut self, _node: &CstNode) {}
    /// Visit a `For` statement node.
    fn visit_for_statement(&mut self, _node: &CstNode) {}
    /// Visit a `For Each` statement node.
    fn visit_for_each_statement(&mut self, _node: &CstNode) {}
    /// Visit a `While` statement node.
    fn visit_while_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Do` statement node.
    fn visit_do_statement(&mut self, _node: &CstNode) {}
    /// Visit a `With` statement node.
    fn visit_with_statement(&mut self, _node: &CstNode) {}
    /// Visit an `Exit` statement node.
    fn visit_exit_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Goto` statement node.
    fn visit_goto_statement(&mut self, _node: &CstNode) {}
    /// Visit an `On Error` statement node.
    fn visit_on_error_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Resume` statement node.
    fn visit_resume_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Call` statement node.
    fn visit_call_statement(&mut self, _node: &CstNode) {}
    /// Visit an `Assignment` statement node.
    fn visit_assignment_statement(&mut self, _node: &CstNode) {}
    /// Visit a `ReDim` statement node.
    fn visit_redim_statement(&mut self, _node: &CstNode) {}
    /// Visit an `Erase` statement node.
    fn visit_erase_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Type` statement node.
    fn visit_type_statement(&mut self, _node: &CstNode) {}
    /// Visit an `Enum` statement node.
    fn visit_enum_statement(&mut self, _node: &CstNode) {}

    // Expression visitors

    /// Visit an `Identifier` node.
    fn visit_identifier(&mut self, _node: &CstNode) {}
    /// Visit a `Literal` expression node.
    fn visit_literal_expression(&mut self, _node: &CstNode) {}

    // Module-level visitors

    /// Visit an `Option` statement node.
    fn visit_option_statement(&mut self, _node: &CstNode) {}
    /// Visit an `Attribute` statement node.
    fn visit_attribute_statement(&mut self, _node: &CstNode) {}
    /// Visit an `Implements` statement node.
    fn visit_implements_statement(&mut self, _node: &CstNode) {}
    /// Visit an `Event` statement node.
    fn visit_event_statement(&mut self, _node: &CstNode) {}
    /// Visit a `Declare` statement node.
    fn visit_declare_statement(&mut self, _node: &CstNode) {}

    // Token visitors

    /// Visit a `Keyword` token node.
    fn visit_keyword(&mut self, _node: &CstNode) {}
    /// Visit an `Operator` token node.
    fn visit_operator(&mut self, _node: &CstNode) {}
    /// Visit a `Punctuation` token node.
    fn visit_punctuation(&mut self, _node: &CstNode) {}
    /// Visit an `EndOfLineComment` token node.
    fn visit_end_of_line_comment(&mut self, _node: &CstNode) {}
    /// Visit a `RemComment` token node.
    fn visit_rem_comment(&mut self, _node: &CstNode) {}
    /// Visit a `Whitespace` token node.
    fn visit_whitespace(&mut self, _node: &CstNode) {}
}

/// Mutable visitor for CST transformations.
///
/// Implement this trait to perform transformations on the CST.
/// Override only the `visit_*_mut` methods for node types you're interested in.
///
/// **Note**: Currently `CstNode` is designed to be immutable. This trait is provided
/// for future extensibility if mutable tree operations are needed.
pub trait VisitorMut: Sized {
    /// Visit a node mutably in the CST.
    ///
    /// The default implementation calls `walk_node_mut` to perform depth-first traversal.
    fn visit_node_mut(&mut self, node: &mut CstNode) {
        walk_node_mut(self, node);
    }

    // Statement visitors

    /// Visit a `Dim` statement node mutably.
    fn visit_dim_statement_mut(&mut self, _node: &mut CstNode) {}
    /// Visit a `Sub` statement node mutably.
    fn visit_sub_statement_mut(&mut self, _node: &mut CstNode) {}
    /// Visit a `Function` statement node mutably.
    fn visit_function_statement_mut(&mut self, _node: &mut CstNode) {}
    /// Visit an `If` statement node mutably.
    fn visit_if_statement_mut(&mut self, _node: &mut CstNode) {}
    /// Visit a `For` statement node mutably.
    fn visit_for_statement_mut(&mut self, _node: &mut CstNode) {}
    /// Visit a `While` statement node mutably.
    fn visit_while_statement_mut(&mut self, _node: &mut CstNode) {}
    /// Visit an `Assignment` statement node mutably.
    fn visit_assignment_statement_mut(&mut self, _node: &mut CstNode) {}

    // Expression visitors

    /// Visit an `Identifier` node mutably.
    fn visit_identifier_mut(&mut self, _node: &mut CstNode) {}
    /// Visit a `Literal` expression node mutably.
    fn visit_literal_expression_mut(&mut self, _node: &mut CstNode) {}
    /// Visit a `Binary` expression node mutably.
    fn visit_binary_expression_mut(&mut self, _node: &mut CstNode) {}
}

/// Perform depth-first traversal of a CST node, dispatching to type-specific visit methods.
///
/// This is the default traversal strategy used by [`Visitor::visit_node`].
/// It visits the current node first (pre-order), then recursively visits all children.
///
/// # Important: After calling a type-specific visitor method, `walk_node` automatically
/// visits all children. Therefore, **do NOT call `walk_node(self, node)` from within
/// type-specific visitor methods** (like `visit_sub_statement`) as this causes infinite
/// recursion.
///
/// # Examples
///
/// **Correct usage** - Override `visit_node` for pre/post processing:
/// ```rust
/// use vb6parse::parsers::cst::{CstNode, Visitor, walk_node};
///
/// struct MyVisitor;
///
/// impl Visitor for MyVisitor {
///     fn visit_node(&mut self, node: &CstNode) {
///         // Custom pre-processing
///         println!("Visiting: {:?}", node.kind());
///         
///         // Standard traversal (calls type-specific methods + visits children)
///         walk_node(self, node);
///         
///         // Custom post-processing
///         println!("Done with: {:?}", node.kind());
///     }
/// }
/// ```
///
/// **Incorrect usage** - Calling `walk_node` from type-specific methods:
/// ```rust,ignore
/// impl Visitor for MyVisitor {
///     fn visit_sub_statement(&mut self, node: &CstNode) {
///         println!("Found sub");
///         walk_node(self, node);  // ❌ INFINITE RECURSION!
///     }
/// }
/// ```
pub fn walk_node<V: Visitor>(visitor: &mut V, node: &CstNode) {
    // Dispatch to type-specific visit method based on node kind
    match node.kind() {
        // Statements
        SyntaxKind::DimStatement => visitor.visit_dim_statement(node),
        SyntaxKind::ConstStatement => visitor.visit_const_statement(node),
        SyntaxKind::SubStatement => visitor.visit_sub_statement(node),
        SyntaxKind::FunctionStatement => visitor.visit_function_statement(node),
        SyntaxKind::PropertyStatement => visitor.visit_property_statement(node),
        SyntaxKind::IfStatement => visitor.visit_if_statement(node),
        SyntaxKind::SelectCaseStatement => visitor.visit_select_case_statement(node),
        SyntaxKind::ForStatement => visitor.visit_for_statement(node),
        SyntaxKind::ForEachStatement => visitor.visit_for_each_statement(node),
        SyntaxKind::WhileStatement => visitor.visit_while_statement(node),
        SyntaxKind::DoStatement => visitor.visit_do_statement(node),
        SyntaxKind::WithStatement => visitor.visit_with_statement(node),
        SyntaxKind::ExitStatement => visitor.visit_exit_statement(node),
        SyntaxKind::GotoStatement => visitor.visit_goto_statement(node),
        SyntaxKind::OnErrorStatement => visitor.visit_on_error_statement(node),
        SyntaxKind::ResumeStatement => visitor.visit_resume_statement(node),
        SyntaxKind::CallStatement => visitor.visit_call_statement(node),
        SyntaxKind::AssignmentStatement => visitor.visit_assignment_statement(node),
        SyntaxKind::ReDimStatement => visitor.visit_redim_statement(node),
        SyntaxKind::EraseStatement => visitor.visit_erase_statement(node),
        SyntaxKind::TypeStatement => visitor.visit_type_statement(node),
        SyntaxKind::EnumStatement => visitor.visit_enum_statement(node),

        // Expressions
        SyntaxKind::Identifier => visitor.visit_identifier(node),
        SyntaxKind::LiteralExpression => visitor.visit_literal_expression(node),

        // Module-level
        SyntaxKind::OptionStatement => visitor.visit_option_statement(node),
        SyntaxKind::AttributeStatement => visitor.visit_attribute_statement(node),
        SyntaxKind::ImplementsStatement => visitor.visit_implements_statement(node),
        SyntaxKind::EventStatement => visitor.visit_event_statement(node),
        SyntaxKind::DeclareStatement => visitor.visit_declare_statement(node),

        // Tokens - only visit if they're significant
        kind if is_keyword(kind) => visitor.visit_keyword(node),
        kind if is_operator(kind) => visitor.visit_operator(node),
        kind if is_punctuation(kind) => visitor.visit_punctuation(node),
        SyntaxKind::EndOfLineComment => visitor.visit_end_of_line_comment(node),
        SyntaxKind::RemComment => visitor.visit_rem_comment(node),
        SyntaxKind::Whitespace | SyntaxKind::Newline => visitor.visit_whitespace(node),

        // For all other node types, we don't call a specific visitor method
        _ => {}
    }

    // Recursively visit children
    for child in node.children() {
        visitor.visit_node(child);
    }
}

/// Perform depth-first traversal of a mutable CST node.
///
/// This is the default traversal strategy used by [`VisitorMut::visit_node_mut`].
///
/// **Note**: Currently provided for API completeness. `CstNode` is designed to be
/// immutable, so this trait has limited practical use until mutable tree operations
/// are implemented. For now, it uses the immutable children iterator.
pub fn walk_node_mut<V: VisitorMut>(visitor: &mut V, node: &mut CstNode) {
    // Dispatch to type-specific visit method based on node kind
    match node.kind() {
        SyntaxKind::DimStatement => visitor.visit_dim_statement_mut(node),
        SyntaxKind::SubStatement => visitor.visit_sub_statement_mut(node),
        SyntaxKind::FunctionStatement => visitor.visit_function_statement_mut(node),
        SyntaxKind::IfStatement => visitor.visit_if_statement_mut(node),
        SyntaxKind::ForStatement => visitor.visit_for_statement_mut(node),
        SyntaxKind::WhileStatement => visitor.visit_while_statement_mut(node),
        SyntaxKind::AssignmentStatement => visitor.visit_assignment_statement_mut(node),
        SyntaxKind::Identifier => visitor.visit_identifier_mut(node),
        SyntaxKind::LiteralExpression => visitor.visit_literal_expression_mut(node),
        SyntaxKind::BinaryExpression => visitor.visit_binary_expression_mut(node),
        _ => {}
    }

    // Note: CstNode is currently immutable, so we use the immutable children() method
    // If mutable tree traversal is implemented in the future, this would change to
    // use children_mut() or similar
    let children: Vec<_> = node.children().to_vec();
    for mut child in children {
        visitor.visit_node_mut(&mut child);
    }
}

// Helper functions to categorize token types

fn is_keyword(kind: SyntaxKind) -> bool {
    matches!(
        kind,
        SyntaxKind::SubKeyword
            | SyntaxKind::FunctionKeyword
            | SyntaxKind::EndKeyword
            | SyntaxKind::IfKeyword
            | SyntaxKind::ThenKeyword
            | SyntaxKind::ElseKeyword
            | SyntaxKind::SelectKeyword
            | SyntaxKind::CaseKeyword
            | SyntaxKind::ForKeyword
            | SyntaxKind::ToKeyword
            | SyntaxKind::NextKeyword
            | SyntaxKind::WhileKeyword
            | SyntaxKind::WendKeyword
            | SyntaxKind::DoKeyword
            | SyntaxKind::LoopKeyword
            | SyntaxKind::DimKeyword
            | SyntaxKind::AsKeyword
            | SyntaxKind::ConstKeyword
            | SyntaxKind::PrivateKeyword
            | SyntaxKind::PublicKeyword
            | SyntaxKind::StaticKeyword
            | SyntaxKind::TypeKeyword
            | SyntaxKind::EnumKeyword
            | SyntaxKind::WithKeyword
            | SyntaxKind::NewKeyword
            | SyntaxKind::SetKeyword
            | SyntaxKind::GetKeyword
            | SyntaxKind::LetKeyword
            | SyntaxKind::PropertyKeyword
            | SyntaxKind::ByValKeyword
            | SyntaxKind::ByRefKeyword
            | SyntaxKind::OptionalKeyword
            | SyntaxKind::ParamArrayKeyword
            | SyntaxKind::CallKeyword
            | SyntaxKind::GotoKeyword
            | SyntaxKind::OnKeyword
            | SyntaxKind::ErrorKeyword
            | SyntaxKind::ResumeKeyword
            | SyntaxKind::ExitKeyword
            | SyntaxKind::ReDimKeyword
            | SyntaxKind::PreserveKeyword
            | SyntaxKind::EraseKeyword
            | SyntaxKind::OptionKeyword
            | SyntaxKind::ExplicitKeyword
            | SyntaxKind::CompareKeyword
            | SyntaxKind::AttributeKeyword
            | SyntaxKind::ImplementsKeyword
            | SyntaxKind::EventKeyword
            | SyntaxKind::RaiseEventKeyword
            | SyntaxKind::DeclareKeyword
            | SyntaxKind::LibKeyword
            | SyntaxKind::AliasKeyword
    )
}

fn is_operator(kind: SyntaxKind) -> bool {
    matches!(
        kind,
        SyntaxKind::AdditionOperator
            | SyntaxKind::SubtractionOperator
            | SyntaxKind::MultiplicationOperator
            | SyntaxKind::DivisionOperator
            | SyntaxKind::BackwardSlashOperator
            | SyntaxKind::ExponentiationOperator
            | SyntaxKind::Ampersand
            | SyntaxKind::EqualityOperator
            | SyntaxKind::LessThanOperator
            | SyntaxKind::GreaterThanOperator
            | SyntaxKind::LessThanOrEqualOperator
            | SyntaxKind::GreaterThanOrEqualOperator
            | SyntaxKind::InequalityOperator
            | SyntaxKind::AndKeyword
            | SyntaxKind::OrKeyword
            | SyntaxKind::NotKeyword
            | SyntaxKind::XorKeyword
            | SyntaxKind::EqvKeyword
            | SyntaxKind::ImpKeyword
            | SyntaxKind::ModKeyword
            | SyntaxKind::IsKeyword
            | SyntaxKind::LikeKeyword
    )
}

fn is_punctuation(kind: SyntaxKind) -> bool {
    matches!(
        kind,
        SyntaxKind::LeftParenthesis
            | SyntaxKind::RightParenthesis
            | SyntaxKind::Comma
            | SyntaxKind::PeriodOperator
            | SyntaxKind::ColonOperator
            | SyntaxKind::Semicolon
            | SyntaxKind::Octothorpe
            | SyntaxKind::ExclamationMark
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ConcreteSyntaxTree;

    #[test]
    fn visitor_counts_identifiers() {
        struct IdentifierCounter {
            count: usize,
        }

        impl Visitor for IdentifierCounter {
            fn visit_identifier(&mut self, _node: &CstNode) {
                self.count += 1;
            }
        }

        let source = "Sub Test()\nDim x As Integer\nDim y As String\nx = y\nEnd Sub";
        let (cst_opt, _) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("Failed to parse");
        let root = cst.to_root_node();

        let mut counter = IdentifierCounter { count: 0 };
        counter.visit_node(&root);

        // Should find identifiers: Test, x, y, x, y
        // Note: Integer and String are keywords, not identifiers
        assert!(counter.count >= 5);
    }

    #[test]
    fn visitor_collects_sub_names() {
        struct SubCollector {
            names: Vec<String>,
        }

        impl Visitor for SubCollector {
            fn visit_sub_statement(&mut self, node: &CstNode) {
                // Extract the sub name (first identifier in the sub statement)
                if let Some(id) = node.find(SyntaxKind::Identifier) {
                    self.names.push(id.text().to_string());
                }
                // Note: Don't call walk_node here - the default visit_node already traverses
            }
        }

        let source = "Sub Alpha()\nEnd Sub\nSub Beta()\nEnd Sub\nSub Gamma()\nEnd Sub";
        let (cst_opt, _) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("Failed to parse");
        let root = cst.to_root_node();

        let mut collector = SubCollector { names: Vec::new() };
        collector.visit_node(&root);

        assert_eq!(collector.names.len(), 3);
        assert!(collector.names.contains(&"Alpha".to_string()));
        assert!(collector.names.contains(&"Beta".to_string()));
        assert!(collector.names.contains(&"Gamma".to_string()));
    }

    #[test]
    fn visitor_finds_dim_statements() {
        struct DimCounter {
            count: usize,
        }

        impl Visitor for DimCounter {
            fn visit_dim_statement(&mut self, _node: &CstNode) {
                self.count += 1;
            }
        }

        let source = "Sub Test()\nDim x As Integer\nDim y, z As String\nEnd Sub";
        let (cst_opt, _) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("Failed to parse");
        let root = cst.to_root_node();

        let mut counter = DimCounter { count: 0 };
        counter.visit_node(&root);

        assert_eq!(counter.count, 2);
    }

    #[test]
    fn visitor_multi_pass_analysis() {
        struct DeclarationFinder {
            declarations: Vec<String>,
        }

        impl Visitor for DeclarationFinder {
            fn visit_dim_statement(&mut self, node: &CstNode) {
                self.declarations.push(node.text().to_string());
            }
        }

        struct StatementCounter {
            subs: usize,
            dims: usize,
        }

        impl Visitor for StatementCounter {
            fn visit_sub_statement(&mut self, _node: &CstNode) {
                self.subs += 1;
            }

            fn visit_dim_statement(&mut self, _node: &CstNode) {
                self.dims += 1;
            }
        }

        let source =
            "Sub Test()\nDim x As Integer\nEnd Sub\nSub Another()\nDim y As String\nEnd Sub";
        let (cst_opt, _) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("Failed to parse");
        let root = cst.to_root_node();

        // First pass
        let mut decl_finder = DeclarationFinder {
            declarations: Vec::new(),
        };
        decl_finder.visit_node(&root);
        assert_eq!(decl_finder.declarations.len(), 2);

        // Second pass
        let mut counter = StatementCounter { subs: 0, dims: 0 };
        counter.visit_node(&root);
        assert_eq!(counter.subs, 2);
        assert_eq!(counter.dims, 2);
    }

    #[test]
    fn visitor_handles_nested_structures() {
        struct NestedCounter {
            if_count: usize,
            for_count: usize,
        }

        impl Visitor for NestedCounter {
            fn visit_if_statement(&mut self, _node: &CstNode) {
                self.if_count += 1;
            }

            fn visit_for_statement(&mut self, _node: &CstNode) {
                self.for_count += 1;
            }
        }

        let source = r#"Sub Test()
If x > 0 Then
    For i = 1 To 10
        If i < 5 Then
        End If
    Next i
End If
End Sub"#;

        let (cst_opt, _) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("Failed to parse");
        let root = cst.to_root_node();

        let mut counter = NestedCounter {
            if_count: 0,
            for_count: 0,
        };
        counter.visit_node(&root);

        assert_eq!(counter.if_count, 2);
        assert_eq!(counter.for_count, 1);
    }
}
