//! Concrete Syntax Tree (CST) implementation for VB6.
//!
//! This module provides a CST that wraps the rowan library internally while
//! providing a public API that doesn't expose rowan types directly.
//!
//! # Overview
//!
//! The CST (Concrete Syntax Tree) represents the complete structure of VB6 source code,
//! including all tokens such as whitespace, comments, and keywords. Unlike an AST
//! (Abstract Syntax Tree), a CST preserves all the original formatting and structure
//! of the source code, making it ideal for tools like formatters, linters, and
//! source-to-source transformations.
//!
//! # Architecture
//!
//! This implementation uses the [`rowan`](https://docs.rs/rowan/) library internally
//! for efficient CST representation, but all rowan types are kept private to the module.
//! The public API only exposes:
//!
//! - [`ConcreteSyntaxTree`] - The main CST struct
//! - [`SyntaxKind`] - An enum representing all possible node and token types
//! - [`parse`] - A function to parse a [`TokenStream`] into a CST
//!
//! # Example Usage
//!
//! ```rust
//! use vb6parse::language::VB6Token;
//! use vb6parse::parsers::cst::parse;
//! use vb6parse::tokenstream::TokenStream;
//!
//! // Create a token stream
//! let tokens = vec![
//!     ("Sub", VB6Token::SubKeyword),
//!     (" ", VB6Token::Whitespace),
//!     ("Main", VB6Token::Identifier),
//!     ("(", VB6Token::LeftParenthesis),
//!     (")", VB6Token::RightParenthesis),
//!     ("\n", VB6Token::Newline),
//! ];
//!
//! let token_stream = TokenStream::new("test.bas".to_string(), tokens);
//!
//! // Parse into a CST
//! let cst = parse(token_stream);
//!
//! // Use the CST
//! println!("Text: {}", cst.text());
//! println!("Children: {}", cst.child_count());
//! ```
//!
//! # Design Principles
//!
//! 1. **No rowan types exposed**: All public APIs use custom types that don't expose rowan.
//! 2. **Complete representation**: The CST includes all tokens, including whitespace and comments.
//! 3. **Efficient**: Uses rowan's red-green tree architecture for memory efficiency.
//! 4. **Type-safe**: All syntax kinds are represented as a Rust enum for compile-time safety.

use std::num::NonZeroUsize;

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;
use crate::tokenize::tokenize;
use crate::tokenstream::TokenStream;
use crate::ParseResult;
use crate::SourceStream;
use crate::VB6CodeErrorKind;
use rowan::{GreenNode, GreenNodeBuilder, Language};

// Submodules for organized CST parsing
mod array_statements;
mod assignment;
mod built_in_statements;
mod conditionals;
mod controlflow;
mod declarations;
mod for_statements;
mod helpers;
mod if_statements;
mod loop_statements;
mod object_statements;
mod select_statements;

/// The language type for VB6 syntax trees.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum VB6Language {}

impl Language for VB6Language {
    type Kind = SyntaxKind;

    fn kind_from_raw(raw: rowan::SyntaxKind) -> Self::Kind {
        SyntaxKind::from_raw(raw)
    }

    fn kind_to_raw(kind: Self::Kind) -> rowan::SyntaxKind {
        kind.to_raw()
    }
}

/// A Concrete Syntax Tree for VB6 code.
///
/// This structure wraps the rowan library's GreenNode internally but provides
/// a public API that doesn't expose rowan types.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConcreteSyntaxTree {
    /// The root green node (internal implementation detail)
    root: GreenNode,
}

impl ConcreteSyntaxTree {
    /// Create a new CST from a GreenNode (internal use only)
    fn new(root: GreenNode) -> Self {
        Self { root }
    }

    pub fn from_source<'a, S>(
        file_name: S,
        contents: &'a str,
    ) -> ParseResult<'a, Self, VB6CodeErrorKind>
    where
        S: Into<String>,
    {
        let mut source_stream = SourceStream::new(file_name.into(), contents);
        let token_stream = tokenize(&mut source_stream);

        if token_stream.result.is_none() {
            return ParseResult {
                result: None,
                failures: token_stream.failures,
            };
        }

        ParseResult {
            result: Some(parse(token_stream.result.unwrap())),
            failures: token_stream.failures,
        }
    }

    /// Get the kind of the root node
    pub fn root_kind(&self) -> SyntaxKind {
        SyntaxKind::from_raw(self.root.kind())
    }

    /// Get a textual representation of the tree structure (for debugging)
    pub fn debug_tree(&self) -> String {
        let syntax_node = rowan::SyntaxNode::<VB6Language>::new_root(self.root.clone());
        format!("{:#?}", syntax_node)
    }

    /// Get the text content of the entire tree
    pub fn text(&self) -> String {
        let syntax_node = rowan::SyntaxNode::<VB6Language>::new_root(self.root.clone());
        syntax_node.text().to_string()
    }

    /// Get the number of children of the root node
    pub fn child_count(&self) -> usize {
        self.root.children().count()
    }

    /// Get the children of the root node
    ///
    /// Returns a vector of child nodes with their kind and text content.
    pub fn children(&self) -> Vec<CstNode> {
        let syntax_node = rowan::SyntaxNode::<VB6Language>::new_root(self.root.clone());
        syntax_node
            .children_with_tokens()
            .map(|child| Self::build_cst_node(child))
            .collect()
    }

    /// Recursively build a CstNode from a rowan NodeOrToken
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
                    .map(|child| Self::build_cst_node(child))
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
    /// * `kind` - The SyntaxKind to search for
    ///
    /// # Returns
    ///
    /// A vector of all child nodes matching the specified kind
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
    /// * `kind` - The SyntaxKind to search for
    ///
    /// # Returns
    ///
    /// `true` if at least one node of the specified kind exists, `false` otherwise
    pub fn contains_kind(&self, kind: SyntaxKind) -> bool {
        self.children().iter().any(|child| child.kind == kind)
    }

    /// Get the first child node (including tokens)
    ///
    /// # Returns
    ///
    /// The first child node if it exists, `None` otherwise
    pub fn first_child(&self) -> Option<CstNode> {
        self.children().into_iter().next()
    }

    /// Get the last child node (including tokens)
    ///
    /// # Returns
    ///
    /// The last child node if it exists, `None` otherwise
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
    pub fn child_at(&self, index: usize) -> Option<CstNode> {
        self.children().into_iter().nth(index)
    }
}

/// Represents a node in the Concrete Syntax Tree
///
/// This can be either a structural node (like SubStatement) or a token (like Identifier).
#[derive(Debug, Clone, PartialEq, Eq)]
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

/// Parse a TokenStream into a Concrete Syntax Tree.
///
/// This function takes a TokenStream and constructs a CST that represents
/// the structure of the VB6 code.
///
/// # Arguments
///
/// * `tokens` - The token stream to parse
///
/// # Returns
///
/// A `ConcreteSyntaxTree` representing the parsed code.
///
/// # Example
///
/// ```rust
/// use vb6parse::tokenstream::TokenStream;
/// use vb6parse::parsers::cst::parse;
///
/// let tokens = TokenStream::new("example.bas".to_string(), vec![]);
/// let cst = parse(tokens);
/// ```
pub fn parse(tokens: TokenStream) -> ConcreteSyntaxTree {
    let parser = Parser::new(tokens);
    parser.parse_module()
}

/// Internal parser state for building the CST
struct Parser<'a> {
    tokens: Vec<(&'a str, VB6Token)>,
    pos: usize,
    builder: GreenNodeBuilder<'static>,
    parsing_header: bool,
}

impl<'a> Parser<'a> {
    fn new(token_stream: TokenStream<'a>) -> Self {
        Parser {
            tokens: token_stream.tokens,
            pos: 0,
            builder: GreenNodeBuilder::new(),
            parsing_header: true,
        }
    }

    /// Parse a complete module (the top-level structure)
    ///
    /// This function loops through all tokens and identifies what kind of
    /// VB6 construct to parse based on the current token. As more VB6 syntax
    /// is supported, additional branches can be added to this loop.
    fn parse_module(mut self) -> ConcreteSyntaxTree {
        self.builder.start_node(SyntaxKind::Root.to_raw());

        while !self.is_at_end() {
            // For a CST, we need to consume ALL tokens, including whitespace and comments
            // We look ahead to determine structure, but still consume everything

            // Check what kind of statement or declaration we're looking at
            match self.current_token() {
                // Attribute statement: Attribute VB_Name = "..."
                Some(VB6Token::AttributeKeyword) => {
                    self.parse_attribute_statement();
                }
                Some(VB6Token::OptionKeyword) => {
                    self.parse_option_statement();
                }
                // Sub procedure: Sub Name(...)
                Some(VB6Token::SubKeyword) => {
                    self.parse_sub_statement();
                }
                // Function Procedure Syntax:
                //
                // [Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]
                //
                Some(VB6Token::FunctionKeyword) => {
                    self.parse_function_statement();
                }
                // Property Procedure Syntax:
                //
                // [Public | Private | Friend] [ Static ] Property Get|Let|Set name [ ( arglist ) ] [ As type ]
                //
                Some(VB6Token::PropertyKeyword) => {
                    self.parse_property_statement();
                }
                // Variable declarations: Dim/Const
                // For Public/Private/Friend/Static, we need to look ahead to see if it's a
                // function/sub declaration or a variable declaration
                Some(VB6Token::DimKeyword) | Some(VB6Token::ConstKeyword) => {
                    self.parse_dim();
                }
                // Public/Private/Friend/Static - could be function/sub/property or declaration
                Some(VB6Token::PrivateKeyword)
                | Some(VB6Token::PublicKeyword)
                | Some(VB6Token::FriendKeyword)
                | Some(VB6Token::StaticKeyword) => {
                    // Look ahead to see if this is a function/sub/property declaration
                    // Peek at the next 2 keywords to handle cases like "Public Static Function"
                    let next_keywords: Vec<_> = self
                        .peek_next_count_keywords(NonZeroUsize::new(2).unwrap())
                        .collect();

                    let procedure_type = match next_keywords.as_slice() {
                        // Direct: Public/Private/Friend Function, Sub, or Property
                        [VB6Token::FunctionKeyword, ..] => Some(0), // Function
                        [VB6Token::SubKeyword, ..] => Some(1),      // Sub
                        [VB6Token::PropertyKeyword, ..] => Some(2), // Property
                        // With Static: Public/Private/Friend Static Function, Sub, or Property
                        [VB6Token::StaticKeyword, VB6Token::FunctionKeyword] => Some(0),
                        [VB6Token::StaticKeyword, VB6Token::SubKeyword] => Some(1),
                        [VB6Token::StaticKeyword, VB6Token::PropertyKeyword] => Some(2),
                        // Anything else is a declaration
                        _ => None,
                    };

                    match procedure_type {
                        Some(0) => self.parse_function_statement(), // Function
                        Some(1) => self.parse_sub_statement(),      // Sub
                        Some(2) => self.parse_property_statement(), // Property
                        _ => self.parse_dim(),                      // Declaration
                    }
                }
                // Whitespace and newlines - consume directly
                Some(VB6Token::Whitespace)
                | Some(VB6Token::Newline)
                | Some(VB6Token::EndOfLineComment)
                | Some(VB6Token::RemComment) => {
                    self.consume_token();
                }
                // Anything else - check if it's a statement, label, assignment, or unknown
                _ => {
                    // Try built-in statements
                    if self.is_builtin_statement_keyword() {
                        self.parse_builtin_statement();
                    // Try array statements
                    } else if self.is_array_statement_keyword() {
                        self.parse_array_statement();
                    // Try to parse common statements using centralized dispatcher
                    } else if self.is_statement_keyword() {
                        self.parse_statement();
                    // Check if this is a label (identifier followed by colon)
                    } else if self.is_at_label() {
                        self.parse_label_statement();
                    // Check if this looks like an assignment statement (identifier = expression)
                    } else if self.is_at_assignment() {
                        self.parse_assignment_statement();
                    } else if self.is_identifier() {
                        self.consume_token();
                    } else if self.is_keyword() {
                        self.consume_token();
                    } else {
                        // This is purely being done this way to make it easier during development.
                        // In a full implementation, we would have specific parsing functions
                        // for all VB6 constructs with anything unrecognized being treated as an error node.
                        self.consume_token_as_unknown();
                    }
                }
            }
        }

        self.builder.finish_node(); // Root

        let root = self.builder.finish();
        ConcreteSyntaxTree::new(root)
    }

    /// Parse an Attribute statement: Attribute VB_Name = "value"
    fn parse_attribute_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::AttributeStatement.to_raw());

        // Consume "Attribute" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // AttributeStatement
    }

    /// Parse an Option statement: Option Explicit On/Off
    fn parse_option_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::OptionStatement.to_raw());

        // Consume "Option" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // OptionStatement
    }

    /// Check if the current token is a control flow keyword.
    fn is_control_flow_keyword(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::IfKeyword)
                | Some(VB6Token::SelectKeyword)
                | Some(VB6Token::ForKeyword)
                | Some(VB6Token::DoKeyword)
                | Some(VB6Token::GotoKeyword)
                | Some(VB6Token::ExitKeyword)
        )
    }

    /// Dispatch control flow statement parsing to the appropriate parser.
    fn parse_control_flow_statement(&mut self) {
        match self.current_token() {
            Some(VB6Token::IfKeyword) => {
                self.parse_if_statement();
            }
            Some(VB6Token::SelectKeyword) => {
                self.parse_select_case_statement();
            }
            Some(VB6Token::ForKeyword) => {
                // Peek ahead to see if next keyword is "Each"
                if let Some(VB6Token::EachKeyword) = self.peek_next_keyword() {
                    self.parse_for_each_statement();
                } else {
                    self.parse_for_statement();
                }
            }
            Some(VB6Token::DoKeyword) => {
                self.parse_do_statement();
            }
            Some(VB6Token::GotoKeyword) => {
                self.parse_goto_statement();
            }
            Some(VB6Token::ExitKeyword) => {
                self.parse_exit_statement();
            }
            _ => {}
        }
    }

    /// Check if the current token is an array statement keyword.
    fn is_array_statement_keyword(&self) -> bool {
        matches!(self.current_token(), Some(VB6Token::ReDimKeyword))
    }

    /// Dispatch array statement parsing to the appropriate parser.
    fn parse_array_statement(&mut self) {
        match self.current_token() {
            Some(VB6Token::ReDimKeyword) => {
                self.parse_redim_statement();
            }
            _ => {}
        }
    }

    /// Parse a code block, consuming tokens until a termination condition is met.
    ///
    /// This is a generic code block parser that can handle different termination conditions:
    /// - End Sub, End Function, End If, etc.
    /// - ElseIf or Else (for If statements)
    ///
    /// # Arguments
    /// * `stop_conditions` - A closure that returns true when the block should stop parsing
    fn parse_code_block<F>(&mut self, stop_conditions: F)
    where
        F: Fn(&Parser) -> bool,
    {
        // Code blocks can appear in both header and body, so we do not modify parsing_header here.

        // Start a CodeBlock node
        self.builder.start_node(SyntaxKind::CodeBlock.to_raw());

        while !self.is_at_end() {
            if stop_conditions(self) {
                break;
            }

            // Try control flow statements first
            if self.is_control_flow_keyword() {
                self.parse_control_flow_statement();
                continue;
            }

            // Try built-in statements
            if self.is_builtin_statement_keyword() {
                self.parse_builtin_statement();
                continue;
            }

            // Try array statements
            if self.is_array_statement_keyword() {
                self.parse_array_statement();
                continue;
            }

            // Try to parse a statement using the centralized dispatcher
            if self.is_statement_keyword() {
                self.parse_statement();
                continue;
            }

            // Handle other constructs that aren't in parse_statement
            match self.current_token() {
                // Variable declarations: Dim/Private/Public/Const/Static
                Some(VB6Token::DimKeyword)
                | Some(VB6Token::PrivateKeyword)
                | Some(VB6Token::PublicKeyword)
                | Some(VB6Token::ConstKeyword)
                | Some(VB6Token::StaticKeyword) => {
                    self.parse_dim();
                }
                // Whitespace and newlines - consume directly
                Some(VB6Token::Whitespace)
                | Some(VB6Token::Newline)
                | Some(VB6Token::EndOfLineComment)
                | Some(VB6Token::RemComment) => {
                    self.consume_token();
                }
                // Anything else - check if it's a label, assignment, or unknown
                _ => {
                    // Check if this is a label (identifier followed by colon)
                    if self.is_at_label() {
                        self.parse_label_statement();
                    // Check if this looks like an assignment statement (identifier = expression)
                    } else if self.is_at_assignment() {
                        self.parse_assignment_statement();
                    } else {
                        self.consume_token_as_unknown();
                    }
                }
            }
        }
        self.builder.finish_node(); // CodeBlock
    }
}

#[test]
fn parse_single_quote_comment() {
    let code = "' This is a comment\nSub Main()\n";

    let mut source_stream = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    assert_eq!(cst.root_kind(), SyntaxKind::Root);
    // Should have 2 children: the comment and the SubStatement
    assert_eq!(cst.child_count(), 3); // 2 statements + EOF
    assert!(cst.text().contains("' This is a comment"));
    assert!(cst.text().contains("Sub Main()"));

    // Use navigation methods
    assert!(cst.contains_kind(SyntaxKind::EndOfLineComment));
    assert!(cst.contains_kind(SyntaxKind::SubStatement));

    let first = cst.first_child().unwrap();
    assert_eq!(first.kind, SyntaxKind::EndOfLineComment);
    assert!(first.is_token);
}

#[cfg(test)]
mod test {
    use crate::*;
    #[test]
    fn parse_rem_comment() {
        let source = "REM This is a REM comment\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        // Should have 2 children: the REM comment and the SubStatement
        assert_eq!(cst.child_count(), 3); // 2 statements + EOF
        assert!(cst.text().contains("REM This is a REM comment"));
        assert!(cst.text().contains("Sub Test()"));

        // Verify REM comment is preserved
        let debug = cst.debug_tree();
        assert!(debug.contains("RemComment"));
    }

    #[test]
    fn parse_mixed_comments() {
        let source = "' Single quote comment\nREM REM comment\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        // Should have 5 children: EndOfLineComment, Newline, RemComment, Newline, SubStatement
        assert_eq!(cst.child_count(), 5);
        assert!(cst.text().contains("' Single quote comment"));
        assert!(cst.text().contains("REM REM comment"));

        // Use navigation methods
        let children = cst.children();
        assert_eq!(children[0].kind, SyntaxKind::EndOfLineComment);
        assert_eq!(children[1].kind, SyntaxKind::Newline);
        assert_eq!(children[2].kind, SyntaxKind::RemComment);
        assert_eq!(children[3].kind, SyntaxKind::Newline);
        assert_eq!(children[4].kind, SyntaxKind::SubStatement);

        assert!(cst.contains_kind(SyntaxKind::EndOfLineComment));
        assert!(cst.contains_kind(SyntaxKind::RemComment));
    }

    #[test]
    fn cst_with_comments() {
        let source = "' This is a comment\nSub Main()\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        // Now has 3 children: comment token, newline token, SubStatement
        assert_eq!(cst.child_count(), 3);
        assert!(cst.text().contains("' This is a comment"));
        assert!(cst.text().contains("Sub Main()"));
    }
}
