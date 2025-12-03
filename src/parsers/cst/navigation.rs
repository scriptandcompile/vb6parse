//! Navigation methods for the Concrete Syntax Tree.
//!
//! This module provides methods for navigating and querying the CST structure,
//! including finding children, accessing specific nodes, and querying node properties.

use crate::parsers::SyntaxKind;

use super::{ConcreteSyntaxTree, VB6Language};

/// Represents a node in the Concrete Syntax Tree
///
/// This can be either a structural node (like `SubStatement`) or a token (like Identifier).
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
pub struct CstNode {
    /// The kind of syntax element this node represents
    pub kind: SyntaxKind,
    /// The text content of this node
    pub text: String,
    /// Whether this is a token (true) or a structural node (false)
    pub is_token: bool,
    /// The children of this node (empty for tokens)
    pub children: Vec<CstNode>,
}

impl ConcreteSyntaxTree {
    /// Get a textual representation of the tree structure (for debugging)
    #[must_use]
    pub fn debug_tree(&self) -> String {
        let syntax_node = rowan::SyntaxNode::<VB6Language>::new_root(self.root.clone());
        format!("{syntax_node:#?}")
    }

    /// Get the text content of the entire tree
    #[must_use]
    pub fn text(&self) -> String {
        let syntax_node = rowan::SyntaxNode::<VB6Language>::new_root(self.root.clone());
        syntax_node.text().to_string()
    }

    /// Get the number of children of the root node
    #[must_use]
    pub fn child_count(&self) -> usize {
        self.root.children().count()
    }

    /// Get the children of the root node
    ///
    /// Returns a vector of child nodes with their kind and text content.
    #[must_use]
    pub fn children(&self) -> Vec<CstNode> {
        let syntax_node = rowan::SyntaxNode::<VB6Language>::new_root(self.root.clone());
        syntax_node
            .children_with_tokens()
            .map(Self::build_cst_node)
            .collect()
    }

    /// Recursively build a `CstNode` from a rowan `NodeOrToken`
    fn build_cst_node(
        node_or_token: rowan::NodeOrToken<
            rowan::SyntaxNode<VB6Language>,
            rowan::SyntaxToken<VB6Language>,
        >,
    ) -> CstNode {
        match node_or_token {
            rowan::NodeOrToken::Node(node) => {
                let children = node
                    .children_with_tokens()
                    .map(Self::build_cst_node)
                    .collect();

                CstNode {
                    kind: node.kind(),
                    text: node.text().to_string(),
                    is_token: false,
                    children,
                }
            }
            rowan::NodeOrToken::Token(token) => CstNode {
                kind: token.kind(),
                text: token.text().to_string(),
                is_token: true,
                children: Vec::new(),
            },
        }
    }

    /// Find all child nodes of a specific kind
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// A vector of all child nodes matching the specified kind
    #[must_use]
    pub fn find_children_by_kind(&self, kind: SyntaxKind) -> Vec<CstNode> {
        self.children()
            .into_iter()
            .filter(|child| child.kind == kind)
            .collect()
    }

    /// Check if the tree contains any node of the specified kind
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// `true` if at least one node of the specified kind exists, `false` otherwise
    #[must_use]
    pub fn contains_kind(&self, kind: SyntaxKind) -> bool {
        self.children().iter().any(|child| child.kind == kind)
    }

    /// Get the first child node (including tokens)
    ///
    /// # Returns
    ///
    /// The first child node if it exists, `None` otherwise
    #[must_use]
    pub fn first_child(&self) -> Option<CstNode> {
        self.children().into_iter().next()
    }

    /// Get the last child node (including tokens)
    ///
    /// # Returns
    ///
    /// The last child node if it exists, `None` otherwise
    #[must_use]
    pub fn last_child(&self) -> Option<CstNode> {
        self.children().into_iter().last()
    }

    /// Get child at a specific index
    ///
    /// # Arguments
    ///
    /// * `index` - The index of the child to retrieve
    ///
    /// # Returns
    ///
    /// The child at the specified index if it exists, `None` otherwise
    #[must_use]
    pub fn child_at(&self, index: usize) -> Option<CstNode> {
        self.children().into_iter().nth(index)
    }
}

#[cfg(test)]
mod test {
    use crate::parsers::{ConcreteSyntaxTree, SyntaxKind};

    // Navigation method tests

    #[test]
    fn navigation_children() {
        let source = "Attribute VB_Name\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let children = cst.children();

        assert_eq!(children.len(), 2); // AttributeStatement, SubStatement
        assert_eq!(children[0].kind, SyntaxKind::AttributeStatement);
        assert_eq!(children[1].kind, SyntaxKind::SubStatement);
        assert!(!children[0].is_token);
        assert!(!children[1].is_token);
    }

    #[test]
    fn navigation_find_children_by_kind() {
        let source = "Dim x\nDim y\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        // Find all DimStatements
        let dim_statements = cst.find_children_by_kind(SyntaxKind::DimStatement);
        assert_eq!(dim_statements.len(), 2);

        // Find all SubStatements
        let sub_statements = cst.find_children_by_kind(SyntaxKind::SubStatement);
        assert_eq!(sub_statements.len(), 1);
    }

    #[test]
    fn navigation_contains_kind() {
        let source = "Sub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert!(cst.contains_kind(SyntaxKind::SubStatement));
        assert!(!cst.contains_kind(SyntaxKind::FunctionStatement));
        assert!(!cst.contains_kind(SyntaxKind::DimStatement));
    }

    #[test]
    fn navigation_first_and_last_child() {
        let source = "Attribute VB_Name\nDim x\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let first = cst.first_child().unwrap();
        assert_eq!(first.kind, SyntaxKind::AttributeStatement);
        assert_eq!(first.text, "Attribute VB_Name\n");

        let last = cst.last_child().unwrap();
        assert_eq!(last.kind, SyntaxKind::SubStatement);
    }

    #[test]
    fn navigation_child_at() {
        let source = "Attribute VB_Name\nDim x\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let first = cst.child_at(0).unwrap();
        assert_eq!(first.kind, SyntaxKind::AttributeStatement);

        let second = cst.child_at(1).unwrap();
        assert_eq!(second.kind, SyntaxKind::DimStatement);

        let third = cst.child_at(2).unwrap();
        assert_eq!(third.kind, SyntaxKind::SubStatement);

        // Fourth is EOF, out of bounds after that
        assert!(cst.child_at(4).is_none());
    }

    #[test]
    fn navigation_empty_tree() {
        let source = "";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        // Even empty code has no children now
        assert_eq!(cst.children().len(), 0);
        assert!(cst.first_child().is_none());
        assert!(cst.last_child().is_none());
        assert!(cst.child_at(0).is_none());
        assert!(!cst.contains_kind(SyntaxKind::SubStatement));
    }

    #[test]
    fn navigation_with_comments_and_whitespace() {
        let source = "' Comment\n\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let children = cst.children();

        // Should have 4 children: EndOfLineComment, newline, newline, SubStatement
        assert_eq!(children.len(), 4);

        // First is the comment
        assert_eq!(children[0].kind, SyntaxKind::EndOfLineComment);
        assert!(children[0].is_token);

        // Second is newline
        assert_eq!(children[1].kind, SyntaxKind::Newline);
        assert!(children[1].is_token);

        // Third is the second newline
        assert_eq!(children[2].kind, SyntaxKind::Newline);
        assert!(children[2].is_token);

        // Fourth is SubStatement
        assert_eq!(children[3].kind, SyntaxKind::SubStatement);
        assert!(!children[3].is_token);
    }
}
