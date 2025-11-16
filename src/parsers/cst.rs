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
//!     ("(", VB6Token::LeftParentheses),
//!     (")", VB6Token::RightParentheses),
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

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;
use crate::tokenstream::TokenStream;
use rowan::{GreenNode, GreenNodeBuilder, Language};

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
            .map(|child| match child {
                rowan::NodeOrToken::Node(node) => CstNode {
                    kind: node.kind(),
                    text: node.text().to_string(),
                    is_token: false,
                },
                rowan::NodeOrToken::Token(token) => CstNode {
                    kind: token.kind(),
                    text: token.text().to_string(),
                    is_token: true,
                },
            })
            .collect()
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
}

impl<'a> Parser<'a> {
    fn new(token_stream: TokenStream<'a>) -> Self {
        Parser {
            tokens: token_stream.tokens,
            pos: 0,
            builder: GreenNodeBuilder::new(),
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
                // Function procedure: Function Name(...) As Type
                Some(VB6Token::FunctionKeyword) => {
                    self.parse_function_statement();
                }
                // Variable declarations: Dim/Private/Public/Const/Static
                Some(VB6Token::DimKeyword) 
                | Some(VB6Token::PrivateKeyword)
                | Some(VB6Token::PublicKeyword)
                | Some(VB6Token::ConstKeyword)
                | Some(VB6Token::StaticKeyword) => {
                    self.parse_declaration();
                }
                // Whitespace and newlines - consume directly
                Some(VB6Token::Whitespace) 
                | Some(VB6Token::Newline)
                | Some(VB6Token::EndOfLineComment)
                | Some(VB6Token::RemComment) => {
                    self.consume_token();
                }
                // Anything else - consume as unknown for now
                _ => {
                    // This is purely being done this way to make it easier during development.
                    // In a full implementation, we would have specific parsing functions
                    // for all VB6 constructs with anything unrecognized being treated as an error node.
                    self.consume_token_as_unknown();
                }
            }
        }
        
        self.builder.finish_node(); // Root
        
        let root = self.builder.finish();
        ConcreteSyntaxTree::new(root)
    }
    
    /// Parse an Attribute statement: Attribute VB_Name = "value"
    fn parse_attribute_statement(&mut self) {
        self.builder.start_node(SyntaxKind::AttributeStatement.to_raw());
        
        // Consume "Attribute" keyword
        self.consume_token();
        
        // Consume everything until newline (preserving all tokens)
        while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        self.builder.finish_node(); // AttributeStatement
    }

    /// Parse an Option statement: Option Explicit On/Off
    fn parse_option_statement(&mut self) {
        self.builder.start_node(SyntaxKind::OptionStatement.to_raw());
        
        // Consume "Option" keyword
        self.consume_token();
        
        // Consume everything until newline (preserving all tokens)
        while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        self.builder.finish_node(); // OptionStatement
    }
    
    /// Parse a Sub procedure: Sub Name(...) ... End Sub
    fn parse_sub_statement(&mut self) {
        self.builder.start_node(SyntaxKind::SubStatement.to_raw());
        
        // Consume "Sub" keyword
        self.consume_token();
        
        // Consume any whitespace after "Sub"
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
        }
        
        // Consume procedure name
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        }
        
        // Consume any whitespace before parameter list
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
        }
        
        // Parse parameter list if present
        if self.at_token(VB6Token::LeftParentheses) {
            self.parse_parameter_list();
        }
        
        // Consume everything until newline (preserving all tokens)
        while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Parse body until "End Sub"
        while !self.is_at_end() {
            if self.at_keyword(VB6Token::EndKeyword) {
                // Look ahead to see if it's "End Sub"
                if self.peek_next_keyword() == Some(VB6Token::SubKeyword) {
                    // Consume "End"
                    self.consume_token();
                    
                    // Consume any whitespace between "End" and "Sub"
                    while self.at_token(VB6Token::Whitespace) {
                        self.consume_token();
                    }
                    
                    // Consume "Sub"
                    self.consume_token();
                    
                    // Consume until newline (including it)
                    while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }
                    if self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }
                    break;
                }
            }
            
            self.consume_token();
        }
        
        self.builder.finish_node(); // SubStatement
    }
    
    /// Parse a Function procedure: Function Name(...) As Type ... End Function
    fn parse_function_statement(&mut self) {
        self.builder.start_node(SyntaxKind::FunctionStatement.to_raw());
        
        // Consume "Function" keyword
        self.consume_token();
        
        // Consume any whitespace after "Function"
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
        }
        
        // Consume function name
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        }
        
        // Consume any whitespace before parameter list
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
        }
        
        // Parse parameter list if present
        if self.at_token(VB6Token::LeftParentheses) {
            self.parse_parameter_list();
        }
        
        // Consume everything until newline (includes "As Type" if present)
        while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Parse body until "End Function"
        while !self.is_at_end() {
            if self.at_keyword(VB6Token::EndKeyword) {
                // Look ahead to see if it's "End Function"
                if self.peek_next_keyword() == Some(VB6Token::FunctionKeyword) {
                    // Consume "End"
                    self.consume_token();
                    
                    // Consume any whitespace between "End" and "Function"
                    while self.at_token(VB6Token::Whitespace) {
                        self.consume_token();
                    }
                    
                    // Consume "Function"
                    self.consume_token();
                    
                    // Consume until newline (including it)
                    while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }
                    if self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }
                    break;
                }
            }
            
            self.consume_token();
        }
        
        self.builder.finish_node(); // FunctionStatement
    }
    
    /// Parse a parameter list: (param1 As Type, param2 As Type)
    fn parse_parameter_list(&mut self) {
        self.builder.start_node(SyntaxKind::ParameterList.to_raw());
        
        // Consume "("
        self.consume_token();
        
        // Consume everything until ")"
        let mut depth = 1;
        while !self.is_at_end() && depth > 0 {
            if self.at_token(VB6Token::LeftParentheses) {
                depth += 1;
            } else if self.at_token(VB6Token::RightParentheses) {
                depth -= 1;
            }
            
            self.consume_token();
            
            if depth == 0 {
                break;
            }
        }
        
        self.builder.finish_node(); // ParameterList
    }
    
    /// Parse a declaration: Dim/Private/Public x As Type
    fn parse_declaration(&mut self) {
        self.builder.start_node(SyntaxKind::DimStatement.to_raw());
        
        // Consume the keyword (Dim, Private, Public, etc.)
        self.consume_token();
        
        // Consume everything until newline (preserving all tokens)
        while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        self.builder.finish_node(); // DimStatement
    }
    
    // Helper methods
    
    fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len()
    }
    
    fn current_token(&self) -> Option<&VB6Token> {
        self.tokens.get(self.pos).map(|(_, token)| token)
    }
    
    fn at_token(&self, token: VB6Token) -> bool {
        self.current_token() == Some(&token)
    }
    
    fn at_keyword(&self, keyword: VB6Token) -> bool {
        self.at_token(keyword)
    }
    
    fn peek_next_keyword(&self) -> Option<VB6Token> {
        let mut i = self.pos + 1;
        while i < self.tokens.len() {
            let (_, token) = &self.tokens[i];
            if *token != VB6Token::Whitespace {
                return Some(*token);
            }
            i += 1;
        }
        None
    }
    
    fn consume_token(&mut self) {
        if let Some((text, token)) = self.tokens.get(self.pos) {
            let kind = SyntaxKind::from(*token);
            self.builder.token(kind.to_raw(), text);
            self.pos += 1;
        }
    }

    fn consume_token_as_unknown(&mut self) {
        if let Some((text, _)) = self.tokens.get(self.pos) {
            self.builder.token(SyntaxKind::Unknown.to_raw(), text);
            self.pos += 1;
        }
    }

}
