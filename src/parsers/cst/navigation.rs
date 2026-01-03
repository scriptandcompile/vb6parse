//! Navigation methods for the Concrete Syntax Tree.
//!
//! This module provides comprehensive methods for navigating and querying the CST structure,
//! including finding children, accessing specific nodes, filtering by predicates, and
//! traversing the entire tree.
//!
//! # Overview
//!
//! The navigation API provides two main types:
//! - [`ConcreteSyntaxTree`] - Represents the root of the tree with navigation methods
//! - [`CstNode`] - Represents individual nodes with the same navigation capabilities
//!
//! Both types provide parallel APIs for consistency and ease of use.
//!
//! # Navigation Patterns
//!
//! ## Basic Child Access
//!
//! Access direct children of a node:
//!
//! ```rust
//! use vb6parse::*;
//!
//! let source = "Sub Test()\nEnd Sub\n";
//! let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//! let root = cst.to_serializable().root;
//!
//! // Count and access children
//! let count = root.child_count();
//! let first = root.first_child();
//! let last = root.last_child();
//! let third = root.child_at(2);
//! ```
//!
//! ## Filtering by Kind
//!
//! Find nodes of a specific [`SyntaxKind`]:
//!
//! ```rust
//! # use vb6parse::*;
//! # use vb6parse::parsers::SyntaxKind;
//! # let source = "Sub Test()\nDim x\nEnd Sub\n";
//! # let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//! # let root = cst.to_serializable().root;
//!
//! // Direct children only
//! let dims: Vec<_> = root.children_by_kind(SyntaxKind::DimStatement).collect();
//! let first_sub = root.first_child_by_kind(SyntaxKind::SubStatement);
//! let has_func = root.contains_kind(SyntaxKind::FunctionStatement);
//! ```
//!
//! ## Recursive Search
//!
//! Search the entire tree depth-first:
//!
//! ```rust
//! # use vb6parse::*;
//! # use vb6parse::parsers::SyntaxKind;
//! # let source = "Sub Test()\nDim x As Integer\nEnd Sub\n";
//! # let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//! # let root = cst.to_serializable().root;
//!
//! // Find first match (depth-first)
//! let dim = root.find(SyntaxKind::DimStatement);
//!
//! // Find all matches
//! let all_identifiers = root.find_all(SyntaxKind::Identifier);
//! println!("Found {} identifiers", all_identifiers.len());
//! ```
//!
//! ## Token Filtering
//!
//! Separate structural nodes from tokens:
//!
//! ```rust
//! # use vb6parse::*;
//! # let source = "Sub Test()\nEnd Sub\n";
//! # let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//! # let root = cst.to_serializable().root;
//!
//! // Get only structural nodes (not tokens)
//! let non_tokens: Vec<_> = root.non_token_children().collect();
//!
//! // Get only token children
//! let tokens: Vec<_> = root.token_children().collect();
//!
//! // Skip whitespace
//! let first_significant = root.first_non_whitespace_child();
//!
//! // Exclude whitespace and newlines
//! let significant: Vec<_> = root.significant_children().collect();
//! ```
//!
//! ## Predicate-Based Search
//!
//! Use custom logic for complex queries:
//!
//! ```rust
//! # use vb6parse::*;
//! # use vb6parse::parsers::SyntaxKind;
//! # let source = "Sub Test()\nDim x As Integer\nEnd Sub\n";
//! # let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//! # let root = cst.to_serializable().root;
//!
//! // Find first non-token node
//! let first_structural = root.find_if(|n| !n.is_token());
//!
//! // Find all keywords
//! let keywords = root.find_all_if(|n| {
//!     matches!(n.kind(),
//!         SyntaxKind::SubKeyword |
//!         SyntaxKind::DimKeyword |
//!         SyntaxKind::AsKeyword
//!     )
//! });
//!
//! // Complex queries
//! let complex_nodes = root.find_all_if(|n| {
//!     !n.is_token() && n.children().len() > 5
//! });
//! ```
//!
//! ## Tree Traversal
//!
//! Iterate over all nodes in depth-first order:
//!
//! ```rust
//! # use vb6parse::*;
//! # let source = "Sub Test()\nDim x\nEnd Sub\n";
//! # let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//! # let root = cst.to_serializable().root;
//!
//! // Iterate all descendants
//! for node in root.descendants() {
//!     if node.is_significant() {
//!         println!("{:?}: {}", node.kind(), node.text().trim());
//!     }
//! }
//!
//! // Count specific types
//! let identifier_count = root.descendants()
//!     .filter(|n| n.kind() == SyntaxKind::Identifier)
//!     .count();
//! ```
//!
//! ## Convenience Checkers
//!
//! Quickly check node properties:
//!
//! ```rust
//! # use vb6parse::*;
//! # let source = "' Comment\nSub Test()\nEnd Sub\n";
//! # let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//! # let root = cst.to_serializable().root;
//!
//! for node in root.descendants() {
//!     if node.is_comment() {
//!         println!("Comment: {}", node.text());
//!     }
//!     if node.is_trivia() {
//!         // Skip whitespace, newlines, and comments
//!         continue;
//!     }
//!     if node.is_significant() {
//!         // Process meaningful nodes
//!     }
//! }
//! ```
//!
//! # Performance Considerations
//!
//! - Most methods return iterators rather than collecting into `Vec` for efficiency
//! - Use `.collect()` only when you need ownership or random access
//! - `find()` stops at the first match, more efficient than `find_all()` for single results
//! - `descendants()` uses a stack-based iterator for memory efficiency
//! - Predicate methods use trait objects internally to avoid code bloat
//!
//! # See Also
//!
//! - [`ConcreteSyntaxTree`] - The main CST type
//! - [`CstNode`] - Individual node type
//! - [`SyntaxKind`] - Enum of all possible node/token types

use crate::parsers::SyntaxKind;

use super::{ConcreteSyntaxTree, VB6Language};

/// Represents a node in the Concrete Syntax Tree
///
/// This can be either a structural node (like `SubStatement`) or a token (like Identifier).
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, Hash)]
pub struct CstNode {
    /// The kind of syntax element this node represents
    kind: SyntaxKind,
    /// The text content of this node
    text: String,
    /// Whether this is a token (true) or a structural node (false)
    is_token: bool,
    /// The children of this node (empty for tokens)
    children: Vec<CstNode>,
}

impl CstNode {
    /// Create a new `CstNode` (internal use only)
    pub(crate) fn new(
        kind: SyntaxKind,
        text: String,
        is_token: bool,
        children: Vec<CstNode>,
    ) -> Self {
        Self {
            kind,
            text,
            is_token,
            children,
        }
    }
}

impl CstNode {
    /// Get the syntax kind of this node
    ///
    /// # Returns
    ///
    /// The `SyntaxKind` representing the type of this syntax element
    ///
    /// # Examples
    ///
    /// ```
    /// # use vb6parse::parsers::cst::ConcreteSyntaxTree;
    /// # use vb6parse::parsers::SyntaxKind;
    /// let (cst, _) = ConcreteSyntaxTree::from_text("test.bas", "Sub Test()\nEnd Sub").unpack();
    /// let cst = cst.unwrap();
    /// let root = cst.to_root_node();
    /// if let Some(child) = root.first_child() {
    ///     let kind = child.kind();
    ///     // kind will be SyntaxKind::Newline or SyntaxKind::SubStatement
    /// }
    /// ```
    #[inline]
    #[must_use]
    pub fn kind(&self) -> SyntaxKind {
        self.kind
    }

    /// Get the text content of this node
    ///
    /// Returns the complete text span covered by this node, including all child nodes
    /// and tokens. For tokens, this is the literal text. For structural nodes, this
    /// is the concatenation of all descendant text.
    ///
    /// # Returns
    ///
    /// A string slice containing the text content of this node
    ///
    /// # Examples
    ///
    /// ```
    /// # use vb6parse::parsers::cst::ConcreteSyntaxTree;
    /// let (cst, _) = ConcreteSyntaxTree::from_text("test.bas", "Dim x As Integer").unpack();
    /// let cst = cst.unwrap();
    /// let root = cst.to_root_node();
    /// if let Some(child) = root.first_child() {
    ///     println!("Text: {}", child.text());
    /// }
    /// ```
    #[inline]
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Check if this node is a token
    ///
    /// Tokens are leaf nodes in the CST that represent individual lexical elements
    /// like identifiers, keywords, operators, literals, etc. Structural nodes are
    /// non-leaf nodes that group tokens and other nodes into syntactic constructs.
    ///
    /// # Returns
    ///
    /// `true` if this is a token node, `false` if it's a structural node
    ///
    /// # Examples
    ///
    /// ```
    /// # use vb6parse::parsers::cst::ConcreteSyntaxTree;
    /// let (cst, _) = ConcreteSyntaxTree::from_text("test.bas", "Sub Test()\nEnd Sub").unpack();
    /// let cst = cst.unwrap();
    /// let root = cst.to_root_node();
    /// for child in root.descendants() {
    ///     if child.is_token() {
    ///         println!("Token: {:?} = '{}'", child.kind(), child.text());
    ///     }
    /// }
    /// ```
    #[inline]
    #[must_use]
    pub fn is_token(&self) -> bool {
        self.is_token
    }

    /// Get a slice of direct child nodes
    ///
    /// Returns a reference to the vector of child nodes. For token nodes, this
    /// will be an empty slice. For structural nodes, this contains all direct
    /// children including both tokens and other structural nodes.
    ///
    /// # Returns
    ///
    /// A slice of child `CstNode` references
    ///
    /// # Examples
    ///
    /// ```
    /// # use vb6parse::parsers::cst::ConcreteSyntaxTree;
    /// let (cst, _) = ConcreteSyntaxTree::from_text("test.bas", "Sub Test()\nEnd Sub").unpack();
    /// let cst = cst.unwrap();
    /// let root = cst.to_root_node();
    ///
    /// // Iterate over children
    /// for child in root.children() {
    ///     println!("Child: {:?}", child.kind());
    /// }
    ///
    /// // Check child count
    /// println!("Number of children: {}", root.children().len());
    ///
    /// // Access by index
    /// if let Some(first) = root.children().get(0) {
    ///     println!("First child: {:?}", first.kind());
    /// }
    /// ```
    #[inline]
    #[must_use]
    pub fn children(&self) -> &[CstNode] {
        &self.children
    }

    /// Get the number of children of this node
    #[must_use]
    pub fn child_count(&self) -> usize {
        self.children.len()
    }

    /// Get the first child node (including tokens)
    ///
    /// # Returns
    ///
    /// The first child node if it exists, `None` otherwise
    #[must_use]
    pub fn first_child(&self) -> Option<&CstNode> {
        self.children.first()
    }

    /// Get the last child node (including tokens)
    ///
    /// # Returns
    ///
    /// The last child node if it exists, `None` otherwise
    #[must_use]
    pub fn last_child(&self) -> Option<&CstNode> {
        self.children.last()
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
    pub fn child_at(&self, index: usize) -> Option<&CstNode> {
        self.children.get(index)
    }

    /// Get an iterator over direct children of a specific kind
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// An iterator over child nodes matching the specified kind
    pub fn children_by_kind(&self, kind: SyntaxKind) -> impl Iterator<Item = &CstNode> {
        self.children()
            .iter()
            .filter(move |child| child.kind() == kind)
    }

    /// Find the first direct child of a specific kind
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// The first child node matching the kind, or `None` if not found
    #[must_use]
    pub fn first_child_by_kind(&self, kind: SyntaxKind) -> Option<&CstNode> {
        self.children().iter().find(|child| child.kind() == kind)
    }

    /// Check if any direct child has the specified kind
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// `true` if at least one child matches the kind, `false` otherwise
    #[must_use]
    pub fn contains_kind(&self, kind: SyntaxKind) -> bool {
        self.children().iter().any(|child| child.kind() == kind)
    }

    /// Find the first descendant node of a specific kind (depth-first search)
    ///
    /// This searches recursively through the entire tree, unlike `first_child_by_kind()`
    /// which only searches direct children.
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// The first descendant node matching the kind, or `None` if not found
    #[must_use]
    pub fn find(&self, kind: SyntaxKind) -> Option<&CstNode> {
        // Check self first
        if self.kind() == kind {
            return Some(self);
        }

        // Then search children recursively
        for child in self.children() {
            if let Some(found) = child.find(kind) {
                return Some(found);
            }
        }

        None
    }

    /// Find all descendant nodes of a specific kind (depth-first search)
    ///
    /// This searches recursively through the entire tree, unlike `children_by_kind()`
    /// which only searches direct children.
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// A vector of all descendant nodes matching the kind
    #[must_use]
    pub fn find_all(&self, kind: SyntaxKind) -> Vec<&CstNode> {
        let mut results = Vec::new();
        self.find_all_recursive(kind, &mut results);
        results
    }

    /// Helper for recursive collection
    fn find_all_recursive<'a>(&'a self, kind: SyntaxKind, results: &mut Vec<&'a CstNode>) {
        if self.kind() == kind {
            results.push(self);
        }

        for child in self.children() {
            child.find_all_recursive(kind, results);
        }
    }

    /// Get an iterator over non-token children (structural nodes only)
    pub fn non_token_children(&self) -> impl Iterator<Item = &CstNode> {
        self.children().iter().filter(|child| !child.is_token())
    }

    /// Get an iterator over token children only
    pub fn token_children(&self) -> impl Iterator<Item = &CstNode> {
        self.children().iter().filter(|child| child.is_token())
    }

    /// Get the first non-whitespace child
    #[must_use]
    pub fn first_non_whitespace_child(&self) -> Option<&CstNode> {
        self.children()
            .iter()
            .find(|child| child.kind() != SyntaxKind::Whitespace)
    }

    /// Get an iterator over significant children (excluding whitespace and newlines)
    pub fn significant_children(&self) -> impl Iterator<Item = &CstNode> {
        self.children().iter().filter(|child| {
            child.kind() != SyntaxKind::Whitespace && child.kind() != SyntaxKind::Newline
        })
    }

    /// Find the first descendant node matching a predicate (depth-first search)
    ///
    /// This allows flexible searching with custom logic beyond just matching kinds.
    ///
    /// # Arguments
    ///
    /// * `predicate` - A closure that takes a `&CstNode` and returns `bool`
    ///
    /// # Returns
    ///
    /// The first descendant node for which the predicate returns `true`, or `None` if not found
    #[must_use]
    pub fn find_if<F>(&self, predicate: F) -> Option<&CstNode>
    where
        F: Fn(&CstNode) -> bool,
    {
        self.find_if_internal(&predicate)
    }

    /// Internal helper for `find_if` that avoids recursion limit issues
    fn find_if_internal(&self, predicate: &dyn Fn(&CstNode) -> bool) -> Option<&CstNode> {
        if predicate(self) {
            return Some(self);
        }

        for child in self.children() {
            if let Some(found) = child.find_if_internal(predicate) {
                return Some(found);
            }
        }

        None
    }

    /// Find all descendant nodes matching a predicate (depth-first search)
    ///
    /// This allows flexible searching with custom logic beyond just matching kinds.
    ///
    /// # Arguments
    ///
    /// * `predicate` - A closure that takes a `&CstNode` and returns `bool`
    ///
    /// # Returns
    ///
    /// A vector of all descendant nodes for which the predicate returns `true`
    #[must_use]
    pub fn find_all_if<F>(&self, predicate: F) -> Vec<&CstNode>
    where
        F: Fn(&CstNode) -> bool,
    {
        let mut results = Vec::new();
        self.find_all_if_internal(&predicate, &mut results);
        results
    }

    /// Internal helper for `find_all_if`
    fn find_all_if_internal<'a>(
        &'a self,
        predicate: &dyn Fn(&CstNode) -> bool,
        results: &mut Vec<&'a CstNode>,
    ) {
        if predicate(self) {
            results.push(self);
        }

        for child in self.children() {
            child.find_all_if_internal(predicate, results);
        }
    }

    /// Get an iterator over all descendants (depth-first, pre-order)
    ///
    /// This visits every node in the tree starting from this node.
    ///
    /// # Returns
    ///
    /// An iterator that yields references to all descendants in depth-first order
    #[must_use]
    pub fn descendants(&self) -> DepthFirstIter<'_> {
        DepthFirstIter { stack: vec![self] }
    }

    /// Get a depth-first iterator over this node and its descendants
    ///
    /// This is an alias for `descendants()` for compatibility with common tree APIs.
    ///
    /// # Returns
    ///
    /// An iterator that yields references to all descendants in depth-first order
    #[must_use]
    pub fn depth_first_iter(&self) -> DepthFirstIter<'_> {
        self.descendants()
    }

    /// Check if this node is whitespace
    #[must_use]
    pub fn is_whitespace(&self) -> bool {
        self.kind() == SyntaxKind::Whitespace
    }

    /// Check if this node is a newline
    #[must_use]
    pub fn is_newline(&self) -> bool {
        self.kind() == SyntaxKind::Newline
    }

    /// Check if this node is a comment
    #[must_use]
    pub fn is_comment(&self) -> bool {
        matches!(
            self.kind(),
            SyntaxKind::EndOfLineComment | SyntaxKind::RemComment
        )
    }

    /// Check if this node is significant (not whitespace, newline, or comment)
    ///
    /// Significant nodes are structural or token nodes that carry semantic meaning,
    /// as opposed to trivia like whitespace and comments.
    #[must_use]
    pub fn is_significant(&self) -> bool {
        !self.is_trivia()
    }

    /// Check if this node is trivia (whitespace, newline, or comment)
    ///
    /// Trivia nodes are formatting elements that don't affect program semantics.
    #[must_use]
    pub fn is_trivia(&self) -> bool {
        self.is_whitespace() || self.is_newline() || self.is_comment()
    }
}

/// Iterator for depth-first traversal of a CST
pub struct DepthFirstIter<'a> {
    stack: Vec<&'a CstNode>,
}

impl<'a> Iterator for DepthFirstIter<'a> {
    type Item = &'a CstNode;

    fn next(&mut self) -> Option<Self::Item> {
        let node = self.stack.pop()?;

        // Push children in reverse order so they're visited left-to-right
        for child in node.children().iter().rev() {
            self.stack.push(child);
        }

        Some(node)
    }
}

/// Iterator for depth-first traversal of a `ConcreteSyntaxTree` (owning version)
pub struct DepthFirstIterOwned {
    stack: Vec<CstNode>,
}

impl Iterator for DepthFirstIterOwned {
    type Item = CstNode;

    fn next(&mut self) -> Option<Self::Item> {
        let node = self.stack.pop()?;

        // Push children in reverse order so they're visited left-to-right
        for child in node.children().iter().rev() {
            self.stack.push(child.clone());
        }

        Some(node)
    }
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

    /// Get an iterator over direct children of a specific kind
    ///
    /// This method returns an iterator for better performance and composability.
    /// If you need a `Vec`, call `.collect()` on the result.
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// An iterator over child nodes matching the specified kind
    ///
    /// # Example
    ///
    /// ```rust
    /// # use vb6parse::ConcreteSyntaxTree;
    /// # use vb6parse::parsers::SyntaxKind;
    /// # let source = "Dim x\nDim y\n";
    /// # let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
    /// // Use iterator directly
    /// for dim_stmt in cst.children_by_kind(SyntaxKind::DimStatement) {
    ///     println!("Found: {}", dim_stmt.text());
    /// }
    ///
    /// // Or collect into a Vec
    /// let dim_stmts: Vec<_> = cst.children_by_kind(SyntaxKind::DimStatement).collect();
    /// ```
    pub fn children_by_kind(&self, kind: SyntaxKind) -> impl Iterator<Item = CstNode> {
        self.children()
            .into_iter()
            .filter(move |child| child.kind() == kind)
    }

    /// Find the first direct child of a specific kind
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// The first child node matching the kind, or `None` if not found
    #[must_use]
    pub fn first_child_by_kind(&self, kind: SyntaxKind) -> Option<CstNode> {
        self.children()
            .into_iter()
            .find(|child| child.kind() == kind)
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
        self.children().iter().any(|child| child.kind() == kind)
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

    /// Find the first descendant node of a specific kind (depth-first search)
    ///
    /// This searches recursively through the entire tree, unlike `first_child_by_kind()`
    /// which only searches direct children.
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// The first descendant node matching the kind, or `None` if not found
    #[must_use]
    pub fn find(&self, kind: SyntaxKind) -> Option<CstNode> {
        let root_node = self.to_root_node();
        root_node.find(kind).cloned()
    }

    /// Find all descendant nodes of a specific kind (depth-first search)
    ///
    /// This searches recursively through the entire tree, unlike `children_by_kind()`
    /// which only searches direct children.
    ///
    /// # Arguments
    ///
    /// * `kind` - The `SyntaxKind` to search for
    ///
    /// # Returns
    ///
    /// A vector of all descendant nodes matching the kind
    #[must_use]
    pub fn find_all(&self, kind: SyntaxKind) -> Vec<CstNode> {
        let root_node = self.to_root_node();
        root_node.find_all(kind).into_iter().cloned().collect()
    }

    /// Get an iterator over non-token children (structural nodes only)
    pub fn non_token_children(&self) -> impl Iterator<Item = CstNode> {
        self.children()
            .into_iter()
            .filter(|child| !child.is_token())
    }

    /// Get an iterator over token children only
    pub fn token_children(&self) -> impl Iterator<Item = CstNode> {
        self.children().into_iter().filter(CstNode::is_token)
    }

    /// Get the first non-whitespace child
    #[must_use]
    pub fn first_non_whitespace_child(&self) -> Option<CstNode> {
        self.children()
            .into_iter()
            .find(|child| child.kind() != SyntaxKind::Whitespace)
    }

    /// Get an iterator over significant children (excluding whitespace and newlines)
    pub fn significant_children(&self) -> impl Iterator<Item = CstNode> {
        self.children().into_iter().filter(|child| {
            child.kind() != SyntaxKind::Whitespace && child.kind() != SyntaxKind::Newline
        })
    }

    /// Find the first descendant node matching a predicate (depth-first search)
    ///
    /// This allows flexible searching with custom logic beyond just matching kinds.
    ///
    /// # Arguments
    ///
    /// * `predicate` - A closure that takes a `&CstNode` and returns `bool`
    ///
    /// # Returns
    ///
    /// The first descendant node for which the predicate returns `true`, or `None` if not found
    #[must_use]
    pub fn find_if<F>(&self, predicate: F) -> Option<CstNode>
    where
        F: Fn(&CstNode) -> bool,
    {
        let root_node = self.to_root_node();
        root_node.find_if(predicate).cloned()
    }

    /// Find all descendant nodes matching a predicate (depth-first search)
    ///
    /// This allows flexible searching with custom logic beyond just matching kinds.
    ///
    /// # Arguments
    ///
    /// * `predicate` - A closure that takes a `&CstNode` and returns `bool`
    ///
    /// # Returns
    ///
    /// A vector of all descendant nodes for which the predicate returns `true`
    #[must_use]
    pub fn find_all_if<F>(&self, predicate: F) -> Vec<CstNode>
    where
        F: Fn(&CstNode) -> bool,
    {
        let root_node = self.to_root_node();
        root_node
            .find_all_if(predicate)
            .into_iter()
            .cloned()
            .collect()
    }

    /// Get an iterator over all descendants (depth-first, pre-order)
    ///
    /// This visits every node in the tree.
    ///
    /// # Returns
    ///
    /// An iterator that yields owned copies of all descendants in depth-first order
    #[must_use]
    pub fn descendants(&self) -> DepthFirstIterOwned {
        let root = self.to_root_node();
        DepthFirstIterOwned { stack: vec![root] }
    }

    /// Get a depth-first iterator over the tree
    ///
    /// This is an alias for `descendants()` for compatibility with common tree APIs.
    ///
    /// # Returns
    ///
    /// An iterator that yields owned copies of all descendants in depth-first order
    #[must_use]
    pub fn depth_first_iter(&self) -> DepthFirstIterOwned {
        self.descendants()
    }
}

#[cfg(test)]
mod tests {
    use crate::parsers::{ConcreteSyntaxTree, SyntaxKind};

    // Navigation method tests

    #[test]
    fn navigation_children() {
        let source = "Attribute VB_Name\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let children = cst.children();

        assert_eq!(children.len(), 2); // AttributeStatement, SubStatement
        assert_eq!(children[0].kind(), SyntaxKind::AttributeStatement);
        assert_eq!(children[1].kind(), SyntaxKind::SubStatement);
        assert!(!children[0].is_token());
        assert!(!children[1].is_token());
    }

    #[test]
    fn navigation_children_by_kind() {
        let source = "Dim x\nDim y\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Find all DimStatements using iterator
        let dim_statements: Vec<_> = cst.children_by_kind(SyntaxKind::DimStatement).collect();
        assert_eq!(dim_statements.len(), 2);

        // Find all SubStatements
        let sub_statements: Vec<_> = cst.children_by_kind(SyntaxKind::SubStatement).collect();
        assert_eq!(sub_statements.len(), 1);

        // Test first_child_by_kind
        assert!(cst.first_child_by_kind(SyntaxKind::DimStatement).is_some());
        assert!(cst
            .first_child_by_kind(SyntaxKind::FunctionStatement)
            .is_none());
    }

    #[test]
    fn navigation_contains_kind() {
        let source = "Sub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert!(cst.contains_kind(SyntaxKind::SubStatement));
        assert!(!cst.contains_kind(SyntaxKind::FunctionStatement));
        assert!(!cst.contains_kind(SyntaxKind::DimStatement));
    }

    #[test]
    fn navigation_first_and_last_child() {
        let source = "Attribute VB_Name\nDim x\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let first = cst.first_child().unwrap();
        assert_eq!(first.kind(), SyntaxKind::AttributeStatement);
        assert_eq!(first.text(), "Attribute VB_Name\n");

        let last = cst.last_child().unwrap();
        assert_eq!(last.kind, SyntaxKind::SubStatement);
    }

    #[test]
    fn navigation_child_at() {
        let source = "Attribute VB_Name\nDim x\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let first = cst.child_at(0).unwrap();
        assert_eq!(first.kind(), SyntaxKind::AttributeStatement);

        let second = cst.child_at(1).unwrap();
        assert_eq!(second.kind(), SyntaxKind::DimStatement);

        let third = cst.child_at(2).unwrap();
        assert_eq!(third.kind, SyntaxKind::SubStatement);

        // Fourth is EOF, out of bounds after that
        assert!(cst.child_at(4).is_none());
    }

    #[test]
    fn navigation_empty_tree() {
        let source = "";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let children = cst.children();

        // Should have 4 children: EndOfLineComment, newline, newline, SubStatement
        assert_eq!(children.len(), 4);

        // First is the comment
        assert_eq!(children[0].kind(), SyntaxKind::EndOfLineComment);
        assert!(children[0].is_token());

        // Second is newline
        assert_eq!(children[1].kind(), SyntaxKind::Newline);
        assert!(children[1].is_token());

        // Third is the second newline
        assert_eq!(children[2].kind(), SyntaxKind::Newline);
        assert!(children[2].is_token());

        // Fourth is SubStatement
        assert_eq!(children[3].kind(), SyntaxKind::SubStatement);
        assert!(!children[3].is_token());
    }

    // CstNode navigation tests

    #[test]
    fn cst_node_basic_navigation() {
        let source = "Sub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let root = cst.to_serializable().root;

        assert_eq!(root.child_count(), 1);
        assert!(root.first_child().is_some());
        assert!(root.last_child().is_some());
        assert!(root.child_at(0).is_some());
        assert!(root.child_at(10).is_none());

        let first = root.first_child().unwrap();
        assert_eq!(first.kind(), SyntaxKind::SubStatement);
    }

    #[test]
    fn cst_node_filter_by_kind() {
        let source = "Dim x\nDim y\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let root = cst.to_serializable().root;

        let dim_stmts: Vec<_> = root.children_by_kind(SyntaxKind::DimStatement).collect();
        assert_eq!(dim_stmts.len(), 2);

        assert!(root.first_child_by_kind(SyntaxKind::DimStatement).is_some());
        assert!(root.contains_kind(SyntaxKind::SubStatement));
        assert!(!root.contains_kind(SyntaxKind::FunctionStatement));
    }

    #[test]
    fn cst_node_recursive_find() {
        let source = "Sub Test()\nDim x As Integer\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let root = cst.to_serializable().root;

        // Find nested DimStatement inside SubStatement
        let dim = root.find(SyntaxKind::DimStatement);
        assert!(dim.is_some());
        assert_eq!(dim.unwrap().kind, SyntaxKind::DimStatement);

        // Find all identifiers (multiple)
        let identifiers = root.find_all(SyntaxKind::Identifier);
        assert!(identifiers.len() >= 2); // "Test" and "x"
    }

    #[test]
    fn cst_node_token_filtering() {
        let source = "Sub Test()\n    Dim x\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let root = cst.to_serializable().root;

        let non_tokens: Vec<_> = root.non_token_children().collect();
        let _tokens: Vec<_> = root.token_children().collect();

        assert!(!non_tokens.is_empty());
        // Should include SubStatement but not Whitespace/Newline tokens

        let first_non_ws = root.first_non_whitespace_child();
        assert!(first_non_ws.is_some());
        assert_ne!(first_non_ws.unwrap().kind, SyntaxKind::Whitespace);

        let significant: Vec<_> = root.significant_children().collect();
        assert!(significant
            .iter()
            .all(|n| { n.kind != SyntaxKind::Whitespace && n.kind != SyntaxKind::Newline }));
    }

    #[test]
    fn concrete_syntax_tree_recursive_find() {
        let source = "Sub Test()\nDim x As Integer\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Find nested DimStatement (not a direct child)
        let dim = cst.find(SyntaxKind::DimStatement);
        assert!(dim.is_some());
        assert_eq!(dim.unwrap().kind, SyntaxKind::DimStatement);

        // Find all identifiers
        let identifiers = cst.find_all(SyntaxKind::Identifier);
        assert!(identifiers.len() >= 2); // "Test" and "x"

        // Compare with non-recursive method (should find nothing for nested nodes)
        let dim_direct: Vec<_> = cst.children_by_kind(SyntaxKind::DimStatement).collect();
        assert_eq!(dim_direct.len(), 0); // DimStatement is inside SubStatement
    }

    #[test]
    fn concrete_syntax_tree_token_filtering() {
        let source = "Sub Test()\n    Dim x\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // The root typically has structural nodes as direct children
        let non_tokens: Vec<_> = cst.non_token_children().collect();
        assert!(!non_tokens.is_empty());

        // Verify non_tokens are indeed structural nodes
        assert!(non_tokens.iter().all(|n| !n.is_token));

        // first_non_whitespace_child should work for roots that start with whitespace
        let source_with_leading_ws = "  \nSub Test()\nEnd Sub\n";
        let cst2 = ConcreteSyntaxTree::from_text("test.bas", source_with_leading_ws).unwrap();
        let first_non_ws = cst2.first_non_whitespace_child();

        if let Some(node) = first_non_ws {
            assert_ne!(node.kind(), SyntaxKind::Whitespace);
        }

        // significant_children should exclude whitespace/newlines
        let significant: Vec<_> = cst.significant_children().collect();
        assert!(significant
            .iter()
            .all(|n| { n.kind != SyntaxKind::Whitespace && n.kind != SyntaxKind::Newline }));
    }

    #[test]
    fn cst_node_predicate_search() {
        let source = "Sub Test()\nDim x As Integer\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let root = cst.to_serializable().root;

        // Find first non-token node
        let first_non_token = root.find_if(|n| !n.is_token);
        assert!(first_non_token.is_some());
        assert!(!first_non_token.unwrap().is_token);

        // Find all keywords
        let keywords = root.find_all_if(|n| {
            matches!(
                n.kind,
                SyntaxKind::SubKeyword | SyntaxKind::DimKeyword | SyntaxKind::AsKeyword
            )
        });
        assert!(keywords.len() >= 3); // Sub, Dim, As
    }

    #[test]
    fn concrete_syntax_tree_predicate_search() {
        let source = "Sub Test()\nDim x As Integer\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Find first non-token node
        let first_non_token = cst.find_if(|n| !n.is_token);
        assert!(first_non_token.is_some());
        assert!(!first_non_token.unwrap().is_token);

        // Find all keywords
        let keywords = cst.find_all_if(|n| {
            matches!(
                n.kind,
                SyntaxKind::SubKeyword | SyntaxKind::DimKeyword | SyntaxKind::AsKeyword
            )
        });
        assert!(keywords.len() >= 3);

        // Complex predicate: find all structural nodes with more than 2 children
        let complex_nodes = cst.find_all_if(|n| !n.is_token && n.children.len() > 2);
        assert!(!complex_nodes.is_empty());
    }

    #[test]
    fn cst_node_convenience_checkers() {
        let source = "' Comment\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let root = cst.to_serializable().root;

        // Find a comment
        let comment = root.find(SyntaxKind::EndOfLineComment);
        assert!(comment.is_some());
        let comment = comment.unwrap();
        assert!(comment.is_comment());
        assert!(comment.is_trivia());
        assert!(!comment.is_significant());

        // Find a structural node
        let sub_stmt = root.find(SyntaxKind::SubStatement);
        assert!(sub_stmt.is_some());
        let sub_stmt = sub_stmt.unwrap();
        assert!(sub_stmt.is_significant());
        assert!(!sub_stmt.is_trivia());
        assert!(!sub_stmt.is_whitespace());
        assert!(!sub_stmt.is_newline());
        assert!(!sub_stmt.is_comment());
    }

    #[test]
    fn cst_node_iterator_traversal() {
        let source = "Sub Test()\nDim x\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let root = cst.to_serializable().root;

        let all_nodes: Vec<_> = root.descendants().collect();
        assert!(!all_nodes.is_empty());
        assert_eq!(all_nodes[0].kind(), SyntaxKind::Root);

        // Count specific node types
        let identifier_count = root
            .descendants()
            .filter(|n| n.kind == SyntaxKind::Identifier)
            .count();
        assert!(identifier_count >= 2); // "Test" and "x"

        // Test depth_first_iter alias
        let count_via_dfs = root.depth_first_iter().count();
        assert_eq!(count_via_dfs, all_nodes.len());
    }

    #[test]
    fn concrete_syntax_tree_iterator_traversal() {
        use crate::parsers::{ConcreteSyntaxTree, CstNode, SyntaxKind};

        let source = "Sub Test()\nDim x\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let all_nodes: Vec<_> = cst.descendants().collect();
        assert!(!all_nodes.is_empty());

        // Count specific node types
        let identifier_count = cst
            .descendants()
            .filter(|n| n.kind == SyntaxKind::Identifier)
            .count();
        assert!(identifier_count >= 2);

        // Test depth_first_iter alias
        let count_via_dfs = cst.depth_first_iter().count();
        assert_eq!(count_via_dfs, all_nodes.len());

        // Test combining with other iterators
        let non_trivia_count = cst.descendants().filter(CstNode::is_significant).count();
        assert!(non_trivia_count > 0);
    }
}
