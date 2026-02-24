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
//! - [`CstNode`] - A structure for navigating and querying the CST
//!
//! # Example Usage
//!
//! ```rust
//! use vb6parse::language::Token;
//! use vb6parse::parsers::cst::parse;
//! use vb6parse::lexer::TokenStream;
//!
//! // Create a token stream
//! let tokens = vec![
//!     ("Sub", Token::SubKeyword),
//!     (" ", Token::Whitespace),
//!     ("Main", Token::Identifier),
//!     ("(", Token::LeftParenthesis),
//!     (")", Token::RightParenthesis),
//!     ("\n", Token::Newline),
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
//! # Navigating the CST
//!
//! The CST provides rich navigation capabilities for traversing and querying the tree.
//! Both [`ConcreteSyntaxTree`] and [`CstNode`] provide parallel navigation APIs:
//!
//! ## Root-Level Navigation
//!
//! ```rust
//! use vb6parse::ConcreteSyntaxTree;
//! use vb6parse::parsers::SyntaxKind;
//!
//! let source = "Sub Test()\nEnd Sub\n";
//! let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//!
//! // Access root-level children
//! println!("Child count: {}", cst.child_count());
//! let first = cst.first_child();
//!
//! // Search root-level children
//! let subs: Vec<_> = cst.children_by_kind(SyntaxKind::SubStatement).collect();
//! ```
//!
//! ## Node-Level Navigation
//!
//! Once you have a [`CstNode`], you can navigate its structure:
//!
//! ```rust
//! # use vb6parse::ConcreteSyntaxTree;
//! # use vb6parse::parsers::SyntaxKind;
//! # let source = "Sub Test()\nDim x\nEnd Sub\n";
//! # let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
//! let root = cst.to_serializable().root;
//!
//! // Direct children
//! println!("Child count: {}", root.child_count());
//! let first = root.first_child();
//!
//! // Filter by kind
//! let statements: Vec<_> = root.children_by_kind(SyntaxKind::DimStatement).collect();
//!
//! // Recursive search
//! let dim_stmt = root.find(SyntaxKind::DimStatement);
//! let all_identifiers = root.find_all(SyntaxKind::Identifier);
//!
//! // Filter tokens
//! let non_tokens: Vec<_> = root.non_token_children().collect();
//! let significant: Vec<_> = root.significant_children().collect();
//!
//! // Custom predicates
//! let keywords = root.find_all_if(|n| {
//!     matches!(n.kind(), SyntaxKind::SubKeyword | SyntaxKind::DimKeyword)
//! });
//!
//! // Iterate all nodes
//! for node in root.descendants() {
//!     if node.is_significant() {
//!         println!("{:?}: {}", node.kind(), node.text());
//!     }
//! }
//! ```
//!
//! ## Navigation Methods
//!
//! Available on both [`ConcreteSyntaxTree`] and [`CstNode`]:
//!
//! **Basic Access:**
//! - `child_count()` - Number of direct children
//! - `first_child()`, `last_child()`, `child_at(index)` - Access specific children
//!
//! **By Kind:**
//! - `children_by_kind(kind)` - Iterator over children of a specific kind
//! - `first_child_by_kind(kind)` - First child of a specific kind
//! - `contains_kind(kind)` - Check if a kind exists in children
//!
//! **Recursive Search:**
//! - `find(kind)` - Find first descendant of a specific kind
//! - `find_all(kind)` - Find all descendants of a specific kind
//!
//! **Token Filtering:**
//! - `non_token_children()` - Structural nodes only
//! - `token_children()` - Tokens only
//! - `first_non_whitespace_child()` - Skip leading whitespace
//! - `significant_children()` - Exclude whitespace and newlines
//!
//! **Predicate-Based:**
//! - `find_if(predicate)` - Find first node matching a custom condition
//! - `find_all_if(predicate)` - Find all nodes matching a custom condition
//!
//! **Tree Traversal:**
//! - `descendants()` - Depth-first iterator over all nodes
//! - `depth_first_iter()` - Alias for `descendants()`
//!
//! **Convenience Checkers** (`CstNode` only):
//! - `is_whitespace()` - Check if node is whitespace.
//! - `is_newline()` - Check if node is newline.
//! - `is_comment()` - Check if node is an end-of-Line or REM comment.
//! - `is_trivia()` - Whitespace, newline, end-of-Line comment, or REM comment.
//! - `is_significant()` - Not trivia.
//!
//! For more details, see the documentation for [`ConcreteSyntaxTree`] and [`CstNode`].
//!
//! # Design Principles
//!
//! 1. **No rowan types exposed**: All public APIs use custom types that don't expose rowan.
//! 2. **Complete representation**: The CST includes all tokens, including whitespace and comments.
//! 3. **Efficient**: Uses rowan's red-green tree architecture for memory efficiency.
//! 4. **Type-safe**: All syntax kinds are represented as a Rust enum for compile-time safety.

use std::collections::HashMap;
use std::num::NonZeroUsize;

use crate::errors::{ErrorKind, FormError};
use crate::files::common::{
    Creatable, Exposed, FileAttributes, FileFormatVersion, NameSpace, ObjectReference,
    PreDeclaredID, Properties,
};
use crate::io::{SourceFile, SourceStream};
use crate::language::{
    CheckBoxProperties, ComboBoxProperties, CommandButtonProperties, Control, ControlKind,
    DataProperties, DirListBoxProperties, DriveListBoxProperties, FileListBoxProperties, Font,
    Form, FormProperties, FormRoot, FrameProperties, LabelProperties, ListBoxProperties, MDIForm,
    MDIFormProperties, MenuControl, MenuProperties, OptionButtonProperties, PictureBoxProperties,
    PropertyGroup, TextBoxProperties, Token,
};
use crate::lexer::{tokenize, TokenStream};
use crate::parsers::SyntaxKind;
use crate::ParseResult;

use either::Either;
use rowan::{GreenNode, GreenNodeBuilder, Language};
use serde::Serialize;

// Submodules for organized CST parsing
mod assignment;
mod attribute_statements;
mod declarations;
mod deftype_statements;
mod enum_statements;
mod for_statements;
mod function_statements;
mod helpers;
mod if_statements;
mod loop_statements;
mod navigation;
mod option_statements;
mod parameters;
mod properties;
mod property_statements;
mod select_statements;
mod sub_statements;
mod type_statements;

// Re-export navigation types
pub use navigation::CstNode;

/// A serializable representation of the CST for snapshot testing.
///
/// This struct wraps the tree structure in a way that can be serialized
/// with serde, making it suitable for use with snapshot testing tools like insta.
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize, Hash)]
pub struct SerializableTree {
    /// The root node of the tree
    pub root: CstNode,
}

/// Helper function to serialize `ConcreteSyntaxTree` as `SerializableTree`
pub(crate) fn serialize_cst<S>(cst: &ConcreteSyntaxTree, serializer: S) -> Result<S::Ok, S::Error>
where
    S: serde::Serializer,
{
    cst.to_serializable().serialize(serializer)
}

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

/// Extract typed property groups from a Vec<PropertyGroup>
fn extract_property_groups(groups: &[PropertyGroup]) -> ExtractedGroups {
    let mut font = None;

    for group in groups {
        if group.name.eq_ignore_ascii_case("Font") {
            if let Ok(f) = Font::try_from(group) {
                font = Some(f);
            }
        }
        // Future: handle other property group types (Images, etc.)
    }

    ExtractedGroups { font }
}

/// Struct to hold extracted property groups for a control
struct ExtractedGroups {
    font: Option<Font>,
    // Future: add other property group types
}

/// A Concrete Syntax Tree for VB6 code.
///
/// This structure wraps the rowan library's `GreenNode` internally but provides
/// a public API that doesn't expose rowan types.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ConcreteSyntaxTree {
    /// The root green node (internal implementation detail)
    root: GreenNode,
}

impl ConcreteSyntaxTree {
    /// Create a new CST from a `GreenNode` (internal use only)
    fn new(root: GreenNode) -> Self {
        Self { root }
    }

    /// Parse a CST from a `SourceFile`.
    ///
    /// # Arguments
    ///
    /// * `source_file` - The source file to parse.
    ///
    /// # Returns
    ///
    /// A result containing the parsed CST or an error.
    #[must_use]
    pub fn from_source(source_file: &SourceFile) -> ParseResult<'_, Self> {
        Self::from_text(
            source_file.file_name().to_string(),
            source_file.source_stream().contents,
        )
    }

    /// Parse a CST from source code.
    ///
    /// # Arguments
    ///
    /// * `file_name` - The name of the source file.
    /// * `contents` - The contents of the source file.
    ///
    /// # Returns
    ///
    /// A result containing the parsed CST or an error.
    pub fn from_text<S>(file_name: S, contents: &str) -> ParseResult<'_, Self>
    where
        S: Into<String>,
    {
        let mut source_stream = SourceStream::new(file_name.into(), contents);
        let token_stream_result = tokenize(&mut source_stream);
        let (token_stream_opt, failures) = token_stream_result.unpack();

        let Some(token_stream) = token_stream_opt else {
            return ParseResult::new(None, failures);
        };

        let cst = parse(token_stream);

        ParseResult::new(Some(cst), failures)
    }

    /// Get the kind of the root node
    #[must_use]
    pub fn root_kind(&self) -> SyntaxKind {
        SyntaxKind::from_raw(self.root.kind())
    }

    /// Convert the CST to a serializable representation.
    ///
    /// This method creates a `SerializableTree` that can be used with
    /// snapshot testing tools like `insta`. The serializable tree contains
    /// the complete tree structure as a hierarchy of `CstNode` instances.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::ConcreteSyntaxTree;
    ///
    /// let source = "Sub Test()\nEnd Sub\n";
    /// let result = ConcreteSyntaxTree::from_text("test.bas", source);
    ///
    /// let (cst_opt, failures) = result.unpack();
    ///
    /// let cst = cst_opt.expect("Failed to parse source");
    ///
    /// if !failures.is_empty() {
    ///     for failure in failures.iter() {
    ///         failure.print();
    ///     }
    ///     panic!("Failed to parse source with {} errors.", failures.len());
    /// };
    ///
    /// let serializable = cst.to_serializable();
    ///
    /// // Can now be used with insta::assert_yaml_snapshot!
    /// ```
    #[must_use]
    pub fn to_serializable(&self) -> SerializableTree {
        SerializableTree {
            root: self.to_root_node(),
        }
    }

    /// Convert the internal rowan tree to a root `CstNode`.
    ///
    /// # Returns
    ///
    /// The root `CstNode` representing the entire CST.
    #[must_use]
    pub fn to_root_node(&self) -> CstNode {
        CstNode::new(SyntaxKind::Root, self.text(), false, self.children())
    }

    /// Create a new CST with specified node kinds removed from the root level.
    ///
    /// This method filters out direct children of the root node that match any of the
    /// specified kinds. This is useful for removing nodes that have already been parsed
    /// into structured data (like version statements, attributes, etc.) to avoid duplication.
    ///
    /// # Arguments
    ///
    /// * `kinds_to_remove` - A slice of `SyntaxKind` values to filter out
    ///
    /// # Returns
    ///
    /// A new `ConcreteSyntaxTree` with the specified kinds removed from the root level.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::ConcreteSyntaxTree;
    /// use vb6parse::parsers::SyntaxKind;
    ///
    /// let source = "VERSION 5.00\nSub Test()\nEnd Sub\n";
    /// let result = ConcreteSyntaxTree::from_text("test.bas", source);
    /// let (cst_opt, failures) = result.unpack();
    /// let cst = cst_opt.expect("Failed to parse source");
    ///
    /// // Remove version statement since it's already parsed
    /// let filtered = cst.without_kinds(&[SyntaxKind::VersionStatement]);
    ///
    /// assert!(!filtered.contains_kind(SyntaxKind::VersionStatement));
    /// ```
    #[must_use]
    pub fn without_kinds(&self, kinds_to_remove: &[SyntaxKind]) -> Self {
        let syntax_node = rowan::SyntaxNode::<VB6Language>::new_root(self.root.clone());
        let mut builder = GreenNodeBuilder::new();

        builder.start_node(SyntaxKind::Root.to_raw());

        // Iterate through children and only add those not in the filter list
        for child in syntax_node.children_with_tokens() {
            let child_kind = match &child {
                rowan::NodeOrToken::Node(node) => node.kind(),
                rowan::NodeOrToken::Token(token) => token.kind(),
            };

            // Skip if this kind should be removed
            if kinds_to_remove.contains(&child_kind) {
                continue;
            }

            // Add the child to the new tree
            Self::clone_node_or_token(&mut builder, child);
        }

        builder.finish_node();
        let new_root = builder.finish();

        Self::new(new_root)
    }

    /// Recursively clone a node or token into a builder
    fn clone_node_or_token(
        builder: &mut GreenNodeBuilder<'static>,
        node_or_token: rowan::NodeOrToken<
            rowan::SyntaxNode<VB6Language>,
            rowan::SyntaxToken<VB6Language>,
        >,
    ) {
        match node_or_token {
            rowan::NodeOrToken::Node(node) => {
                builder.start_node(node.kind().to_raw());
                for child in node.children_with_tokens() {
                    Self::clone_node_or_token(builder, child);
                }
                builder.finish_node();
            }
            rowan::NodeOrToken::Token(token) => {
                builder.token(token.kind().to_raw(), token.text());
            }
        }
    }
}

/// Parse a `TokenStream` into a Concrete Syntax Tree.
///
/// This function takes a `TokenStream` and constructs a CST that represents
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
/// use vb6parse::lexer::TokenStream;
/// use vb6parse::parsers::cst::parse;
///
/// let tokens = TokenStream::new("example.bas".to_string(), vec![]);
/// let cst = parse(tokens);
/// ```
#[must_use]
pub fn parse(tokens: TokenStream) -> ConcreteSyntaxTree {
    let parser = Parser::new(tokens);
    parser.parse_root()
}

/// Internal parser state for building the CST
pub(crate) struct Parser<'a> {
    pub(crate) tokens: Vec<(&'a str, Token)>,
    pub(crate) pos: usize,
    pub(crate) builder: GreenNodeBuilder<'static>,
    pub(crate) parsing_header: bool,
}

impl<'a> Parser<'a> {
    fn new(token_stream: TokenStream<'a>) -> Self {
        Parser {
            tokens: token_stream.into_tokens(),
            pos: 0,
            builder: GreenNodeBuilder::new(),
            parsing_header: true,
        }
    }

    /// Create parser for direct extraction mode (control-only parsing)
    pub(crate) fn new_direct_extraction(tokens: Vec<(&'a str, Token)>, pos: usize) -> Self {
        Parser {
            tokens,
            pos,
            builder: GreenNodeBuilder::new(),
            parsing_header: true,
        }
    }

    // Create parser for hybrid mode (`FormFile` optimization)
    // ==================== Direct Extraction Helpers ====================
    // These methods support direct extraction without CST building

    /// Consume the parser and return the remaining tokens
    /// Used to get tokens after direct extraction for CST building
    pub(crate) fn into_tokens(self) -> Vec<(&'a str, Token)> {
        // Return tokens from current position onwards
        self.tokens[self.pos..].to_vec()
    }

    /// Skip whitespace tokens without consuming them into the CST
    pub(crate) fn skip_whitespace(&mut self) {
        while self.at_token(Token::Whitespace) {
            self.pos += 1;
        }
    }

    /// Skip whitespace and newline tokens without consuming them into the CST
    pub(crate) fn skip_whitespace_and_newlines(&mut self) {
        while self.at_token(Token::Whitespace) || self.at_token(Token::Newline) {
            self.pos += 1;
        }
    }

    /// Consume and advance past the current token without adding to CST
    /// Returns the consumed token for inspection
    pub(crate) fn consume_advance(&mut self) -> Option<(&'a str, Token)> {
        if self.pos < self.tokens.len() {
            let token = self.tokens[self.pos];
            self.pos += 1;
            Some(token)
        } else {
            None
        }
    }

    // ==================== Direct Extraction Methods ====================

    /// Parse VERSION statement directly without building CST
    ///
    /// Extracts the file format version (e.g., \"VERSION 5.00\") by directly
    /// parsing tokens without CST construction overhead.
    ///
    /// # Returns
    ///
    /// A `ParseResult` containing:
    /// - `result`: `Some(FileFormatVersion)` if found and valid, `None` if not present or invalid
    /// - `failures`: Empty vec (no errors generated for missing VERSION)
    pub(crate) fn parse_version_direct(&mut self) -> ParseResult<'a, FileFormatVersion> {
        self.skip_whitespace();

        // Check if VERSION keyword is present
        if !self.at_token(Token::VersionKeyword) {
            return ParseResult::new(None, Vec::new());
        }

        self.consume_advance(); // VERSION keyword
        self.skip_whitespace();

        // Parse version number (e.g., \"5.00\" or \"1.0\")
        let version_result = if let Some((text, token)) = self.tokens.get(self.pos) {
            match token {
                Token::SingleLiteral | Token::DoubleLiteral | Token::IntegerLiteral => {
                    let version_str = text.trim();
                    self.consume_advance();

                    // Parse \"major.minor\" format
                    let parts: Vec<&str> = version_str.split('.').collect();
                    if parts.len() == 2 {
                        if let (Ok(major), Ok(minor)) =
                            (parts[0].parse::<u8>(), parts[1].parse::<u8>())
                        {
                            Some(FileFormatVersion { major, minor })
                        } else {
                            None
                        }
                    } else {
                        None
                    }
                }
                _ => None,
            }
        } else {
            None
        };

        // Skip optional CLASS keyword and trailing whitespace
        self.skip_whitespace();
        if self.at_token(Token::ClassKeyword) {
            self.consume_advance();
        }
        self.skip_whitespace_and_newlines();

        ParseResult::new(version_result, Vec::new())
    }

    // ==================== Core Control Extraction Methods ====================

    /// Check if current token is `BeginProperty` identifier
    fn is_begin_property(&self) -> bool {
        if let Some((text, token)) = self.tokens.get(self.pos) {
            *token == Token::Identifier && text.eq_ignore_ascii_case("BeginProperty")
        } else {
            false
        }
    }

    /// Check if current token is an identifier matching the target text (case-insensitive)
    fn is_identifier_text(&self, target: &str) -> bool {
        if let Some((text, token)) = self.tokens.get(self.pos) {
            *token == Token::Identifier && text.eq_ignore_ascii_case(target)
        } else {
            false
        }
    }

    /// Convert a Menu-typed Control into `MenuControl`
    fn control_to_menu(control: Control) -> MenuControl {
        let (name, tag, index, kind) = control.into_parts();

        if let ControlKind::Menu {
            properties,
            sub_menus,
        } = kind
        {
            MenuControl::new(name, tag, index, properties, sub_menus)
        } else {
            // Fallback: create empty menu control
            MenuControl::new(name, tag, index, MenuProperties::default(), Vec::new())
        }
    }

    /// Parse control type directly from tokens (e.g., "VB.Form", "VB.CommandButton")
    fn parse_control_type_direct(&mut self) -> String {
        let mut parts = Vec::new();

        // Parse identifier or keyword
        if self.is_identifier() || self.at_keyword() {
            if let Some((text, _)) = self.tokens.get(self.pos) {
                parts.push(text.to_string());
                self.consume_advance();
            }
        }

        // Parse dot-separated parts (e.g., "VB.Form")
        while self.at_token(Token::PeriodOperator) {
            self.consume_advance(); // dot
            if self.is_identifier() || self.at_keyword() {
                if let Some((text, _)) = self.tokens.get(self.pos) {
                    parts.push(".".to_string());
                    parts.push(text.to_string());
                    self.consume_advance();
                }
            }
        }

        parts.join("")
    }

    /// Parse control name directly from tokens
    fn parse_control_name_direct(&mut self) -> String {
        if self.is_identifier() || self.at_keyword() {
            if let Some((text, _)) = self.tokens.get(self.pos) {
                let name = text.to_string();
                self.consume_advance();
                return name;
            }
        }
        String::new()
    }

    /// Parse a property assignment (Key = Value) directly from tokens
    /// Returns (key, value) tuple
    fn parse_property_direct(&mut self) -> Option<(String, String)> {
        // Parse property key
        let key = if self.is_identifier() || self.at_keyword() {
            if let Some((text, _)) = self.tokens.get(self.pos) {
                let k = text.to_string();
                self.consume_advance();
                k
            } else {
                return None;
            }
        } else {
            return None;
        };

        self.skip_whitespace();

        // Parse = sign
        if !self.at_token(Token::EqualityOperator) {
            return None;
        }
        self.consume_advance();
        self.skip_whitespace();

        // Parse value (everything until newline/colon)
        // Special case: Resource references like "file.frx":0000 or $"file.frx":0000
        // should include the colon and offset. The quotes should be preserved.
        let mut value_parts = Vec::new();
        let mut in_resource_reference = false;

        while !self.is_at_end() && !self.at_token(Token::Newline) {
            if let Some((text, token)) = self.tokens.get(self.pos) {
                let text_copy = *text;
                let token_copy = *token;

                // Check if we see a dollar sign
                if token_copy == Token::DollarSign {
                    value_parts.push(text_copy);
                    self.consume_advance();
                }
                // If we see a string literal (with or without $), check if resource reference follows
                else if token_copy == Token::StringLiteral {
                    value_parts.push(text_copy);
                    self.consume_advance();

                    // Peek ahead - if next token is colon, this is a resource reference
                    if let Some((_, next_token)) = self.tokens.get(self.pos) {
                        if *next_token == Token::ColonOperator {
                            in_resource_reference = true;
                        }
                    }
                }
                // If in resource reference, capture colon
                else if in_resource_reference && token_copy == Token::ColonOperator {
                    value_parts.push(text_copy);
                    self.consume_advance();
                }
                // If in resource reference and we see the offset number, capture it and stop
                else if in_resource_reference
                    && (token_copy == Token::IntegerLiteral || token_copy == Token::LongLiteral)
                {
                    value_parts.push(text_copy);
                    self.consume_advance();
                    break; // Done with resource reference
                }
                // If we hit a colon and not in resource reference, stop
                else if token_copy == Token::ColonOperator {
                    break;
                }
                // Otherwise, capture the token
                else {
                    value_parts.push(text_copy);
                    self.consume_advance();
                }
            } else {
                break;
            }
        }

        // Skip newline
        self.skip_whitespace_and_newlines();

        // Join tokens directly without intermediate conversion
        let value = value_parts.concat().trim().to_string();
        Some((key, value))
    }

    /// Parse property group directly (BeginProperty...EndProperty)
    fn parse_property_group_direct(&mut self) -> Option<PropertyGroup> {
        // Expect BeginProperty identifier
        if !self.is_identifier_text("BeginProperty") {
            return None;
        }

        self.consume_advance(); // BeginProperty
        self.skip_whitespace();

        // Parse property group name and optional GUID
        let (name, guid) = self.parse_property_group_name_direct();
        self.skip_whitespace_and_newlines();

        // Parse nested properties and property groups
        let mut properties = HashMap::new();

        while !self.is_at_end() && !self.is_identifier_text("EndProperty") {
            self.skip_whitespace();

            if self.is_identifier_text("EndProperty") {
                break;
            }

            if self.is_identifier_text("BeginProperty") {
                // Nested property group
                if let Some(nested_group) = self.parse_property_group_direct() {
                    properties.insert(nested_group.name.clone(), Either::Right(nested_group));
                }
            } else if self.is_identifier() || self.at_keyword() {
                // Regular property
                if let Some((key, value)) = self.parse_property_direct() {
                    properties.insert(key, Either::Left(value));
                }
            } else {
                self.consume_advance();
            }
        }

        // Parse EndProperty
        if self.is_identifier_text("EndProperty") {
            self.consume_advance();
            self.skip_whitespace_and_newlines();
        }

        Some(PropertyGroup {
            name,
            guid,
            properties,
        })
    }

    /// Parse property group name and extract optional GUID
    /// Format: "Name {GUID}" or just "Name"
    fn parse_property_group_name_direct(&mut self) -> (String, Option<uuid::Uuid>) {
        let mut name_parts: Vec<&str> = Vec::new();
        let mut guid_parts: Vec<&str> = Vec::new();
        let mut in_guid = false;

        // Collect tokens until newline
        while !self.is_at_end() && !self.at_token(Token::Newline) {
            if let Some((text, token)) = self.tokens.get(self.pos) {
                if *token == Token::LeftCurlyBrace {
                    // Start of GUID
                    in_guid = true;
                } else if *token == Token::RightCurlyBrace {
                    // End of GUID
                    in_guid = false;
                } else if *token != Token::Whitespace && *token != Token::EndOfLineComment {
                    // Collect non-whitespace tokens
                    if in_guid {
                        guid_parts.push(*text);
                    } else {
                        name_parts.push(*text);
                    }
                }
            }
            self.consume_advance();
        }

        let name = name_parts.concat();
        let guid = if guid_parts.is_empty() {
            None
        } else {
            let guid_str = guid_parts.concat();
            uuid::Uuid::parse_str(&guid_str).ok()
        };

        (name, guid)
    }

    /// Parse Object statements directly (without CST)
    /// Phase 5: Direct extraction of Object references
    pub(crate) fn parse_objects_direct(&mut self) -> Vec<ObjectReference> {
        let mut objects = Vec::new();

        self.skip_whitespace_and_newlines();

        // Continue parsing Object statements until we hit something else
        while self.at_token(Token::ObjectKeyword) {
            if let Some(obj_ref) = self.parse_single_object_direct() {
                objects.push(obj_ref);
            }
            self.skip_whitespace_and_newlines();
        }

        objects
    }

    /// Parse a single Object statement line
    /// Format: Object = "{UUID}#version#flags"; "filename"
    /// Or:     Object = *\G{UUID}#version#flags; "filename"
    fn parse_single_object_direct(&mut self) -> Option<ObjectReference> {
        // Expect "Object" keyword
        if !self.at_token(Token::ObjectKeyword) {
            return None;
        }
        self.consume_advance(); // Object
        self.skip_whitespace();

        // Expect "="
        if !self.at_token(Token::EqualityOperator) {
            return None;
        }
        self.consume_advance(); // =
        self.skip_whitespace();

        // Check for optional "*\G" prefix (embedded object)
        let mut _is_embedded = false;
        if self.at_token(Token::MultiplicationOperator) {
            self.consume_advance(); // *
                                    // Expect \G (backslash followed by identifier "G")
            if let Some((_text, token)) = self.tokens.get(self.pos) {
                if *token == Token::BackwardSlashOperator {
                    self.consume_advance(); // \
                    if let Some((text2, token2)) = self.tokens.get(self.pos) {
                        if *token2 == Token::Identifier && text2.eq_ignore_ascii_case("G") {
                            self.consume_advance(); // G
                            _is_embedded = true;
                        }
                    }
                }
            }
        }
        self.skip_whitespace();

        // Parse first string literal or GUID tokens: "{UUID}#version#flags"
        // The GUID may be tokenized as:
        //   - A StringLiteral: "{UUID}#version#flags"
        //   - Individual tokens: { UUID-parts } # version # flags
        let uuid_part = if let Some((text, token)) = self.tokens.get(self.pos) {
            if *token == Token::StringLiteral {
                // String literal format
                let s = text.trim_matches('"').to_string();
                self.consume_advance();
                s
            } else if *token == Token::LeftCurlyBrace {
                // Token format: { guid-parts } #version# flags
                // Collect all tokens until semicolon
                let mut parts: Vec<&str> = Vec::new();
                while !self.is_at_end() && !self.at_token(Token::Semicolon) {
                    if let Some((text, token)) = self.tokens.get(self.pos) {
                        // Skip whitespace but collect everything else
                        if *token != Token::Whitespace {
                            parts.push(text);
                        }
                        self.consume_advance();
                    } else {
                        break;
                    }
                }
                // Reconstruct the UUID part string
                // Need to convert: { ... } #version# flags -> {UUID}#version#flags
                parts.concat()
            } else {
                return None;
            }
        } else {
            return None;
        };

        self.skip_whitespace();

        // Expect semicolon
        if !self.at_token(Token::Semicolon) {
            return None;
        }
        self.consume_advance(); // ;
        self.skip_whitespace();

        // Parse second string literal: filename
        let file_name = if let Some((text, token)) = self.tokens.get(self.pos) {
            if *token == Token::StringLiteral {
                let s = text.trim_matches('"').to_string();
                self.consume_advance();
                s
            } else {
                return None;
            }
        } else {
            return None;
        };

        // Parse UUID part: {UUID}#version#flags or UUID#version#flags
        let parts: Vec<&str> = uuid_part.split('#').collect();
        if parts.len() >= 3 {
            // Extract UUID (remove braces if present)
            let uuid_str = parts[0].trim_matches(|c| c == '{' || c == '}');

            if let Ok(uuid) = uuid::Uuid::parse_str(uuid_str) {
                let version = parts[1].to_string();
                let unknown1 = parts[2].to_string();

                return Some(ObjectReference::Compiled {
                    uuid,
                    version,
                    unknown1,
                    file_name,
                });
            }
        }

        None
    }

    /// Parse Attribute statements directly (without CST)
    /// Phase 6: Direct extraction of file attributes
    pub(crate) fn parse_attributes_direct(&mut self) -> FileAttributes {
        let mut name = String::new();
        let mut global_name_space = NameSpace::Local;
        let mut creatable = Creatable::True;
        let mut predeclared_id = PreDeclaredID::False;
        let mut exposed = Exposed::False;
        let mut description: Option<String> = None;
        let mut ext_key: HashMap<String, String> = HashMap::new();

        self.skip_whitespace_and_newlines();

        // Continue parsing Attribute statements until we hit something else
        while self.at_token(Token::AttributeKeyword) {
            if let Some((key, value)) = self.parse_single_attribute_direct() {
                // Process the extracted key-value pair
                match key.as_str() {
                    "VB_Name" => {
                        name = value;
                    }
                    "VB_GlobalNameSpace" => {
                        global_name_space = if value == "True" || value == "-1" {
                            NameSpace::Global
                        } else {
                            NameSpace::Local
                        };
                    }
                    "VB_Creatable" => {
                        creatable = if value == "True" || value == "-1" {
                            Creatable::True
                        } else {
                            Creatable::False
                        };
                    }
                    "VB_PredeclaredId" => {
                        predeclared_id = if value == "True" || value == "-1" {
                            PreDeclaredID::True
                        } else {
                            PreDeclaredID::False
                        };
                    }
                    "VB_Exposed" => {
                        exposed = if value == "True" || value == "-1" {
                            Exposed::True
                        } else {
                            Exposed::False
                        };
                    }
                    "VB_Description" => {
                        description = Some(value);
                    }
                    _ => {
                        // Store any other attributes in ext_key
                        ext_key.insert(key, value);
                    }
                }
            }
            self.skip_whitespace_and_newlines();
        }

        FileAttributes {
            name,
            global_name_space,
            creatable,
            predeclared_id,
            exposed,
            description,
            ext_key,
        }
    }

    /// Parse a single Attribute statement line
    /// ```Attribute VB_Name = "Value"```
    /// Or
    /// ```Attribute VB_GlobalNameSpace = True```
    fn parse_single_attribute_direct(&mut self) -> Option<(String, String)> {
        // Expect "Attribute" keyword
        if !self.at_token(Token::AttributeKeyword) {
            return None;
        }
        self.consume_advance(); // Attribute
        self.skip_whitespace();

        // Parse attribute key (e.g., "VB_Name")
        let key = if let Some((text, token)) = self.tokens.get(self.pos) {
            if *token == Token::Identifier {
                let k = text.to_string();
                self.consume_advance();
                k
            } else {
                return None;
            }
        } else {
            return None;
        };

        self.skip_whitespace();

        // Expect "="
        if !self.at_token(Token::EqualityOperator) {
            return None;
        }
        self.consume_advance(); // =
        self.skip_whitespace();

        // Parse value (can be string, True/False, or number)
        let mut value = String::new();
        let mut found_value = false;

        // Check for negative sign first (for values like "-1")
        if self.at_token(Token::SubtractionOperator) {
            value.push('-');
            self.consume_advance();
            self.skip_whitespace();
        }

        if let Some((text, token)) = self.tokens.get(self.pos) {
            match token {
                Token::StringLiteral => {
                    // Remove surrounding quotes
                    value.push_str(text.trim().trim_matches('"'));
                    self.consume_advance();
                    found_value = true;
                }
                Token::TrueKeyword => {
                    value.push_str("True");
                    self.consume_advance();
                    found_value = true;
                }
                Token::FalseKeyword => {
                    value.push_str("False");
                    self.consume_advance();
                    found_value = true;
                }
                Token::IntegerLiteral | Token::LongLiteral => {
                    value.push_str(text.trim());
                    self.consume_advance();
                    found_value = true;
                }
                _ => {}
            }
        }

        // Consume the rest of the line (for complex attributes like VB_Ext_KEY)
        // Skip until we hit a newline or end of tokens
        while self.pos < self.tokens.len() {
            if let Some((_, token)) = self.tokens.get(self.pos) {
                if *token == Token::Newline {
                    break;
                }
                self.consume_advance();
            } else {
                break;
            }
        }

        if found_value {
            Some((key, value))
        } else {
            None
        }
    }

    /// Build `FormRoot` from control type string and properties
    ///
    /// This function is used for parsing top-level form elements.
    /// Only `VB.Form` and `VB.MDIForm` are valid top-level types.
    fn build_form_root(
        control_type: &str,
        control_name: String,
        tag: String,
        index: i32,
        properties: Properties,
        groups: &[PropertyGroup],
        child_controls: Vec<Control>,
        menus: Vec<MenuControl>,
    ) -> Result<FormRoot, ErrorKind> {
        match control_type {
            "VB.Form" => {
                let mut form_properties: FormProperties = properties.into();
                // Override with property group if present
                let extracted_groups = extract_property_groups(groups);
                if let Some(font) = extracted_groups.font {
                    form_properties.font = Some(font);
                }

                Ok(FormRoot::Form(Form {
                    name: control_name,
                    tag,
                    index,
                    properties: form_properties,
                    controls: child_controls,
                    menus,
                }))
            }
            "VB.MDIForm" => {
                let mut mdi_form_properties: MDIFormProperties = properties.into();
                // Override with property group if present
                let extracted_groups = extract_property_groups(groups);
                if let Some(font) = extracted_groups.font {
                    mdi_form_properties.font = Some(font);
                }

                Ok(FormRoot::MDIForm(MDIForm {
                    name: control_name,
                    tag,
                    index,
                    properties: mdi_form_properties,
                    controls: child_controls,
                    menus,
                }))
            }
            _ => Err(ErrorKind::Form(FormError::InvalidTopLevelControl {
                control_type: control_type.to_string(),
            })),
        }
    }

    /// Build `ControlKind` from control type string and properties
    ///
    /// Note: This function rejects `VB.Form` and `VB.MDIForm` as they are now
    /// top-level types only and cannot be child controls.
    fn build_control_kind(
        control_type: &str,
        properties: Properties,
        child_controls: Vec<Control>,
        menus: Vec<MenuControl>,
        property_groups: Vec<PropertyGroup>,
    ) -> ControlKind {
        use ControlKind;
        // Extract typed property groups
        let groups = extract_property_groups(&property_groups);

        match control_type {
            "VB.CommandButton" => {
                let mut props: CommandButtonProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::CommandButton { properties: props }
            }
            "VB.Data" => {
                let mut props: DataProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::Data { properties: props }
            }
            "VB.TextBox" => {
                let mut props: TextBoxProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::TextBox { properties: props }
            }
            "VB.Label" => {
                let mut props: LabelProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::Label { properties: props }
            }
            "VB.CheckBox" => {
                let mut props: CheckBoxProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::CheckBox { properties: props }
            }
            "VB.Line" => ControlKind::Line {
                properties: properties.into(),
            },
            "VB.Shape" => ControlKind::Shape {
                properties: properties.into(),
            },
            "VB.ListBox" => {
                let mut props: ListBoxProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::ListBox { properties: props }
            }
            "VB.ComboBox" => {
                let mut props: ComboBoxProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::ComboBox { properties: props }
            }
            "VB.Timer" => ControlKind::Timer {
                properties: properties.into(),
            },
            "VB.HScrollBar" => ControlKind::HScrollBar {
                properties: properties.into(),
            },
            "VB.VScrollBar" => ControlKind::VScrollBar {
                properties: properties.into(),
            },
            "VB.Frame" => {
                let mut props: FrameProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::Frame {
                    properties: props,
                    controls: child_controls,
                }
            }
            "VB.PictureBox" => {
                let mut props: PictureBoxProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::PictureBox {
                    properties: props,
                    controls: child_controls,
                }
            }
            "VB.FileListBox" => {
                let mut props: FileListBoxProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::FileListBox { properties: props }
            }
            "VB.DirListBox" => {
                let mut props: DirListBoxProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::DirListBox { properties: props }
            }
            "VB.DriveListBox" => {
                let mut props: DriveListBoxProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::DriveListBox { properties: props }
            }
            "VB.Image" => ControlKind::Image {
                properties: properties.into(),
            },
            "VB.OptionButton" => {
                let mut props: OptionButtonProperties = properties.into();
                // Override with property group if present
                if let Some(font) = groups.font {
                    props.font = Some(font);
                }
                ControlKind::OptionButton { properties: props }
            }
            "VB.OLE" => ControlKind::Ole {
                properties: properties.into(),
            },
            "VB.Menu" => ControlKind::Menu {
                properties: properties.into(),
                sub_menus: menus,
            },
            _ => ControlKind::Custom {
                properties: properties.into(),
                property_groups,
            },
        }
    }

    /// Parse properties block directly to Control without building CST
    /// Phase 4: Full implementation with nested controls and property groups
    pub(crate) fn parse_properties_block_to_control(&mut self) -> ParseResult<'a, Control> {
        self.skip_whitespace();

        // Expect BEGIN keyword
        if !self.at_token(Token::BeginKeyword) {
            return ParseResult::new(None, Vec::new());
        }

        self.consume_advance(); // BEGIN
        self.skip_whitespace();

        // Parse control type (e.g., "VB.Form")
        let control_type = self.parse_control_type_direct();
        self.skip_whitespace();

        // Parse control name
        let control_name = self.parse_control_name_direct();
        self.skip_whitespace_and_newlines();

        // Parse properties, child controls, and property groups
        let mut properties = Properties::new();
        let mut child_controls = Vec::new();
        let mut menus = Vec::new();
        let mut property_groups = Vec::new();
        let mut failures = Vec::new();

        while !self.is_at_end() && !self.at_token(Token::EndKeyword) {
            self.skip_whitespace();

            if self.at_token(Token::EndKeyword) {
                break;
            }

            if self.at_token(Token::BeginKeyword) {
                // Nested control (Begin VB.xxx)
                let child_result = self.parse_properties_block_to_control();
                let (child_opt, child_failures) = child_result.unpack();
                failures.extend(child_failures);

                if let Some(child) = child_opt {
                    // Check if it's a menu control
                    if matches!(child.kind(), ControlKind::Menu { .. }) {
                        menus.push(Self::control_to_menu(child));
                    } else {
                        child_controls.push(child);
                    }
                }
            } else if self.is_begin_property() {
                // Parse property group (BeginProperty)
                if let Some(group) = self.parse_property_group_direct() {
                    property_groups.push(group);
                }
            } else if self.is_identifier() || self.at_keyword() {
                // Parse property (Key = Value)
                if let Some((key, value)) = self.parse_property_direct() {
                    // Remove surrounding quotes if this is a simple string literal
                    // BUT NOT if it's a resource reference (contains ":digit" pattern)
                    let is_resource_reference = value.contains(':')
                        && value
                            .split(':')
                            .next_back()
                            .is_some_and(|part| part.chars().all(|c| c.is_ascii_digit()));

                    let cleaned_value = if !is_resource_reference
                        && value.starts_with('"')
                        && value.ends_with('"')
                        && value.len() >= 2
                    {
                        &value[1..value.len() - 1]
                    } else {
                        &value
                    };
                    properties.insert(&key, cleaned_value);
                }
            } else {
                // Skip unknown token
                self.consume_advance();
            }
        }

        // Parse END keyword
        if self.at_token(Token::EndKeyword) {
            self.consume_advance();
            self.skip_whitespace_and_newlines();
        }

        // Extract tag and index from properties
        let tag = properties.get("Tag").cloned().unwrap_or_default();
        let index = properties
            .get("Index")
            .and_then(|s| s.parse().ok())
            .unwrap_or(0);

        // Build control kind with all components
        let kind = Self::build_control_kind(
            &control_type,
            properties,
            child_controls,
            menus,
            property_groups,
        );

        let control = Control::new(control_name, tag, index, kind);

        ParseResult::new(Some(control), failures)
    }

    /// Parse properties block directly to `FormRoot` for top-level form elements.
    ///
    /// This function is specifically for parsing the top-level form element in
    /// `.frm`, `.ctl`, and `.dob` files. It enforces that only `VB.Form` or
    /// `VB.MDIForm` can be used as the root element.
    ///
    /// # Returns
    ///
    /// A `ParseResult` containing either a `FormRoot` (`Form` or `MDIForm`) or `None`,
    /// along with any parsing failures encountered.
    pub(crate) fn parse_properties_block_to_form_root(&mut self) -> ParseResult<'a, FormRoot> {
        use Properties;

        let mut groups = Vec::new();

        self.skip_whitespace();

        // Expect BEGIN keyword
        if !self.at_token(Token::BeginKeyword) {
            return ParseResult::new(None, Vec::new());
        }

        self.consume_advance(); // BEGIN
        self.skip_whitespace();

        // Parse control type (e.g., "VB.Form" or "VB.MDIForm")
        let control_type = self.parse_control_type_direct();
        self.skip_whitespace();

        // Parse control name
        let control_name = self.parse_control_name_direct();
        self.skip_whitespace_and_newlines();

        // Parse properties, child controls, and property groups
        let mut properties = Properties::new();
        let mut child_controls = Vec::new();
        let mut menus = Vec::new();
        let mut failures = Vec::new();

        while !self.is_at_end() && !self.at_token(Token::EndKeyword) {
            self.skip_whitespace();

            if self.at_token(Token::EndKeyword) {
                break;
            }

            if self.at_token(Token::BeginKeyword) {
                // Nested control (Begin VB.xxx) - use parse_properties_block_to_control for children
                let child_result = self.parse_properties_block_to_control();
                let (child_opt, child_failures) = child_result.unpack();
                failures.extend(child_failures);

                if let Some(child) = child_opt {
                    // Check if it's a menu control
                    if matches!(child.kind(), ControlKind::Menu { .. }) {
                        menus.push(Self::control_to_menu(child));
                    } else {
                        child_controls.push(child);
                    }
                }
            } else if self.is_begin_property() {
                // Parse property group (BeginProperty)
                if let Some(group) = self.parse_property_group_direct() {
                    // Property groups are not used in Form/MDIForm, but we parse them anyway
                    // to avoid errors if they appear
                    groups.push(group);
                }
            } else if self.is_identifier() || self.at_keyword() {
                // Parse property (Key = Value)
                if let Some((key, value)) = self.parse_property_direct() {
                    // Remove surrounding quotes if this is a simple string literal
                    // BUT NOT if it's a resource reference (contains ":digit" pattern)
                    let is_resource_reference = value.contains(':')
                        && value
                            .split(':')
                            .next_back()
                            .is_some_and(|part| part.chars().all(|c| c.is_ascii_digit()));

                    let cleaned_value = if !is_resource_reference
                        && value.starts_with('"')
                        && value.ends_with('"')
                        && value.len() >= 2
                    {
                        &value[1..value.len() - 1]
                    } else {
                        &value
                    };
                    properties.insert(&key, cleaned_value);
                }
            } else {
                // Skip unknown token
                self.consume_advance();
            }
        }

        // Parse END keyword
        if self.at_token(Token::EndKeyword) {
            self.consume_advance();
            self.skip_whitespace_and_newlines();
        }

        // Extract tag and index from properties
        let tag = properties.get("Tag").cloned().unwrap_or_default();
        let index = properties
            .get("Index")
            .and_then(|s| s.parse().ok())
            .unwrap_or(0);

        // Build FormRoot with all components
        if let Ok(form_root) = Self::build_form_root(
            &control_type,
            control_name,
            tag,
            index,
            properties,
            &groups,
            child_controls,
            menus,
        ) {
            ParseResult::new(Some(form_root), failures)
        } else {
            // If invalid top-level control type, return a default Form as fallback
            let default_form = FormRoot::Form(Form {
                name: String::new(),
                tag: String::new(),
                index: 0,
                properties: FormProperties::default(),
                controls: Vec::new(),
                menus: Vec::new(),
            });
            ParseResult::new(Some(default_form), failures)
        }
    }

    /// Parse a complete module/class/form (the top-level structure)
    ///
    /// This function loops through all tokens and identifies what kind of
    /// VB6 construct to parse based on the current token. As more VB6 syntax
    /// is supported, additional branches can be added to this loop.
    fn parse_root(mut self) -> ConcreteSyntaxTree {
        self.builder.start_node(SyntaxKind::Root.to_raw());

        // Parse VERSION statement (if present)
        if self.at_token(Token::VersionKeyword) {
            self.parse_version_statement();
        }

        // Parse BEGIN ... END block (if present)
        if self.at_token(Token::BeginKeyword) {
            self.parse_properties_block();
        }

        // Parse Attribute statements (if present)
        // These come after the PropertiesBlock in forms/classes
        while self.at_token(Token::AttributeKeyword) {
            self.parse_attribute_statement();
        }

        self.parse_module_body();
        self.builder.finish_node(); // Root

        let root = self.builder.finish();
        ConcreteSyntaxTree::new(root)
    }

    fn parse_module_body(&mut self) {
        while !self.is_at_end() {
            // For a CST, we need to consume ALL tokens, including whitespace and comments
            // We look ahead to determine structure, but still consume everything

            // Check what kind of statement or declaration we're looking at
            match self.current_token() {
                // BEGIN ... END block (for forms/classes with properties)
                // This can appear after Object statements in form files
                Some(Token::BeginKeyword) => {
                    self.parse_properties_block();
                }
                // Object statement: Object = "{UUID}#version#flags"; "filename"
                // Only parse as ObjectStatement if it matches the proper format
                Some(Token::ObjectKeyword) if self.is_object_statement() => {
                    self.parse_object_statement();
                }
                // Attribute statement: Attribute VB_Name = "..."
                Some(Token::AttributeKeyword) => {
                    self.parse_attribute_statement();
                }
                Some(Token::OptionKeyword) => {
                    // Peek ahead to check if this is Option Base, Option Compare, or Option Private
                    if let Some(Token::BaseKeyword) = self.peek_next_keyword() {
                        self.parse_option_base_statement();
                    } else if let Some(Token::CompareKeyword) = self.peek_next_keyword() {
                        self.parse_option_compare_statement();
                    } else if let Some(Token::PrivateKeyword) = self.peek_next_keyword() {
                        self.parse_option_private_statement();
                    } else {
                        self.parse_option_statement();
                    }
                }
                // DefType statements: DefInt, DefLng, DefStr, etc.
                Some(
                    Token::DefBoolKeyword
                    | Token::DefByteKeyword
                    | Token::DefIntKeyword
                    | Token::DefLngKeyword
                    | Token::DefCurKeyword
                    | Token::DefSngKeyword
                    | Token::DefDblKeyword
                    | Token::DefDecKeyword
                    | Token::DefDateKeyword
                    | Token::DefStrKeyword
                    | Token::DefObjKeyword
                    | Token::DefVarKeyword,
                ) => {
                    self.parse_deftype_statement();
                }
                // Declare statement: Declare Sub/Function Name Lib "..."
                Some(Token::DeclareKeyword) => {
                    self.parse_declare_statement();
                }
                // Event statement: Event Name(...)
                Some(Token::EventKeyword) => {
                    self.parse_event_statement();
                }
                // Implements statement: Implements InterfaceName
                Some(Token::ImplementsKeyword) => {
                    self.parse_implements_statement();
                }
                // Enum statement: Enum Name ... End Enum
                Some(Token::EnumKeyword) => {
                    self.parse_enum_statement();
                }
                // Type statement: Type Name ... End Type
                Some(Token::TypeKeyword) => {
                    self.parse_type_statement();
                }
                // Sub procedure: Sub Name(...)
                Some(Token::SubKeyword) => {
                    self.parse_sub_statement();
                }
                // Function Procedure Syntax:
                //
                // [Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]
                //
                Some(Token::FunctionKeyword) => {
                    self.parse_function_statement();
                }
                // Property Procedure Syntax:
                //
                // [Public | Private | Friend] [ Static ] Property Get|Let|Set name [ ( arglist ) ] [ As type ]
                //
                Some(Token::PropertyKeyword) => {
                    self.parse_property_statement();
                }
                // Variable declarations: Dim/Const
                // For Public/Private/Friend/Static, we need to look ahead to see if it's a
                // function/sub declaration or a variable declaration
                Some(Token::DimKeyword | Token::ConstKeyword) => {
                    self.parse_dim();
                }
                // Public/Private/Friend/Static - could be function/sub/property or declaration
                Some(
                    Token::PrivateKeyword
                    | Token::PublicKeyword
                    | Token::FriendKeyword
                    | Token::StaticKeyword,
                ) => {
                    // Look ahead to see if this is a function/sub/property/enum declaration
                    // Peek at the next 2 keywords to handle cases like "Public Static Function"
                    let next_keywords: Vec<_> = self
                        .peek_next_count_keywords(NonZeroUsize::new(2).unwrap())
                        .collect();

                    match next_keywords.as_slice() {
                        // Direct: Public/Private/Friend Function, Sub, Property, Enum, Type, Declare, or Event
                        [Token::FunctionKeyword, ..] => self.parse_function_statement(), // Function
                        [Token::SubKeyword, ..] => self.parse_sub_statement(),           // Sub
                        [Token::PropertyKeyword, ..] => self.parse_property_statement(), // Property
                        [Token::DeclareKeyword, ..] => self.parse_declare_statement(),   // Declare
                        [Token::EnumKeyword, ..] => self.parse_enum_statement(),         // Enum
                        [Token::TypeKeyword, ..] => self.parse_type_statement(),         // Type
                        [Token::EventKeyword, ..] => self.parse_event_statement(),       // Event
                        [Token::ImplementsKeyword, ..] => self.parse_implements_statement(), // Implements
                        // With Static: Public/Private/Friend Static Function, Sub, or Property
                        [Token::StaticKeyword, Token::FunctionKeyword] => {
                            self.parse_function_statement();
                        }
                        [Token::StaticKeyword, Token::SubKeyword] => {
                            self.parse_sub_statement();
                        }
                        [Token::StaticKeyword, Token::PropertyKeyword] => {
                            self.parse_property_statement();
                        }
                        // Anything else is a declaration
                        _ => self.parse_dim(),
                    }
                }
                // Anything else - check if it's a statement, label, assignment, or unknown
                _ => {
                    // Whitespace, newlines, and comments - consume directly FIRST
                    // This must be checked before is_at_procedure_call to avoid
                    // treating REM comments as procedure calls
                    if matches!(
                        self.current_token(),
                        Some(
                            Token::Whitespace
                                | Token::Newline
                                | Token::EndOfLineComment
                                | Token::RemComment
                        )
                    ) {
                        self.consume_token();
                    // Try control flow statements
                    } else if self.is_control_flow_keyword() {
                        self.parse_control_flow_statement();
                    // Try built-in statements
                    } else if self.is_library_statement_keyword() {
                        self.parse_library_statement();
                    // Try array statements
                    } else if self.is_variable_declaration_keyword() {
                        self.parse_array_statement();
                    // Try to parse common statements using centralized dispatcher
                    } else if self.is_statement_keyword() {
                        self.parse_statement();
                    // Check if this is a label (identifier followed by colon)
                    } else if self.is_at_label() {
                        self.parse_label_statement();
                    // Check for Let statement (optional assignment keyword)
                    } else if self.at_token(Token::LetKeyword) {
                        self.parse_let_statement();
                    // Check if this looks like an assignment statement (identifier = expression)
                    // This must come BEFORE at_keyword() check to handle keywords used as variables
                    } else if self.is_at_assignment() {
                        self.parse_assignment_statement();
                    // Check if this looks like a procedure call (identifier without assignment)
                    } else if self.is_at_procedure_call() {
                        self.parse_procedure_call();
                    } else if self.is_identifier() || self.at_keyword() {
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
    }

    /// Check if the current token is a control flow keyword.
    /// Checks both current position and next non-whitespace token.
    fn is_control_flow_keyword(&self) -> bool {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        matches!(
            token,
            Some(
                Token::IfKeyword
                    | Token::SelectKeyword
                    | Token::ForKeyword
                    | Token::DoKeyword
                    | Token::WhileKeyword
                    | Token::GotoKeyword
                    | Token::GoSubKeyword
                    | Token::ReturnKeyword
                    | Token::ResumeKeyword
                    | Token::ExitKeyword
                    | Token::OnKeyword
            )
        )
    }

    /// Dispatch control flow statement parsing to the appropriate parser.
    fn parse_control_flow_statement(&mut self) {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        match token {
            Some(Token::IfKeyword) => {
                self.parse_if_statement();
            }
            Some(Token::SelectKeyword) => {
                self.parse_select_case_statement();
            }
            Some(Token::ForKeyword) => {
                // Peek ahead to see if next keyword is "Each"
                // Need to peek TWO keywords ahead if we're currently at whitespace
                let next_kw = if self.at_token(Token::Whitespace) {
                    // We peeked to find "For", now peek one more for "Each"
                    self.peek_next_count_keywords(NonZeroUsize::new(2).unwrap())
                        .nth(1)
                } else {
                    // We're directly at "For", peek once for "Each"
                    self.peek_next_keyword()
                };

                if let Some(Token::EachKeyword) = next_kw {
                    self.parse_for_each_statement();
                } else {
                    self.parse_for_statement();
                }
            }
            Some(Token::DoKeyword) => {
                self.parse_do_statement();
            }
            Some(Token::WhileKeyword) => {
                self.parse_while_statement();
            }
            Some(Token::GotoKeyword) => {
                self.parse_goto_statement();
            }
            Some(Token::GoSubKeyword) => {
                self.parse_gosub_statement();
            }
            Some(Token::ReturnKeyword) => {
                self.parse_return_statement();
            }
            Some(Token::ResumeKeyword) => {
                self.parse_resume_statement();
            }
            Some(Token::ExitKeyword) => {
                self.parse_exit_statement();
            }
            Some(Token::OnKeyword) => {
                // Look ahead to distinguish between On Error, On GoTo, and On GoSub
                // Need to peek different amounts depending on whether we're at whitespace
                let next_kw = if self.at_token(Token::Whitespace) {
                    // We peeked to find "On", now peek one more for "Error/GoTo/GoSub"
                    self.peek_next_count_keywords(NonZeroUsize::new(2).unwrap())
                        .nth(1)
                } else {
                    // We're directly at "On", peek once for next keyword
                    self.peek_next_keyword()
                };

                if let Some(Token::ErrorKeyword) = next_kw {
                    self.parse_on_error_statement();
                } else {
                    // Need to scan ahead to find GoTo or GoSub keyword
                    // to distinguish between On GoTo and On GoSub
                    let peek_start = if self.at_token(Token::Whitespace) {
                        2
                    } else {
                        1
                    };
                    let keywords: Vec<Token> = self
                        .peek_next_count_keywords(NonZeroUsize::new(20).unwrap())
                        .skip(peek_start)
                        .collect();

                    let has_goto = keywords.contains(&Token::GotoKeyword);
                    let has_gosub = keywords.contains(&Token::GoSubKeyword);

                    if has_goto {
                        self.parse_on_goto_statement();
                    } else if has_gosub {
                        self.parse_on_gosub_statement();
                    } else {
                        // Fallback - treat as On Error if we can't determine
                        self.parse_on_error_statement();
                    }
                }
            }
            _ => {}
        }
    }

    /// Check if the current token is an array statement keyword.
    /// Checks both current position and next non-whitespace token.
    fn is_variable_declaration_keyword(&self) -> bool {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        matches!(token, Some(Token::ReDimKeyword | Token::EraseKeyword))
    }

    /// Check if we're at an Object statement with proper format.
    ///
    /// Object statements in VB6 forms have the format:
    /// `Object = "{GUID}#version#flags"; "filename"`
    /// or
    /// `Object = *\G{GUID}#version#flags; "filename"`
    ///
    /// This checks if the pattern matches before committing to parse as `ObjectStatement`.
    #[allow(clippy::needless_continue)] // continue on whitespace is needed but clippy is incorrectly catching here.
    fn is_object_statement(&self) -> bool {
        // Must start with Object keyword
        if !self.at_token(Token::ObjectKeyword) {
            return false;
        }

        // Look ahead to verify it matches Object statement pattern
        // Skip whitespace, should find =, then whitespace, then string or *\G pattern
        let mut found_equals = false;
        for (_text, token) in self.tokens.iter().skip(self.pos + 1) {
            match token {
                // TODO: Change this parsing to better handle leading whitespace on object statements.
                Token::Whitespace => continue,
                Token::EqualityOperator if !found_equals => {
                    found_equals = true;
                }
                // After =, we expect either a quoted string starting with { or * for type library refs
                Token::StringLiteral | Token::MultiplicationOperator if found_equals => {
                    // Valid Object statement - string literal after =
                    // or
                    // Could be *\G{ pattern for type libraries
                    return true;
                }
                // If we hit anything else after =, not an Object statement
                _ if found_equals => return false,
                // If we hit a newline before =, not an Object statement
                Token::Newline | Token::EndOfLineComment | Token::RemComment => {
                    return false;
                }
                _ => return false,
            }
        }
        false
    }

    /// Dispatch array statement parsing to the appropriate parser.
    fn parse_array_statement(&mut self) {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        match token {
            Some(Token::ReDimKeyword) => {
                self.parse_redim_statement();
            }
            Some(Token::EraseKeyword) => {
                self.parse_erase_statement();
            }
            _ => {}
        }
    }

    /// Parse a statement list, consuming tokens until a termination condition is met.
    ///
    /// This is a generic statement list parser that can handle different termination conditions:
    /// - End Sub, End Function, End If, etc.
    /// - `ElseIf` or Else (for If statements)
    ///
    /// # Arguments
    /// * `stop_conditions` - A closure that returns true when the block should stop parsing
    pub(crate) fn parse_statement_list<F>(&mut self, stop_conditions: F)
    where
        F: Fn(&Parser) -> bool,
    {
        // Statement lists can appear in both header and body, so we do not modify parsing_header here.

        // Start a StatementList node
        self.builder.start_node(SyntaxKind::StatementList.to_raw());

        while !self.is_at_end() {
            if stop_conditions(self) {
                break;
            }

            // Try control flow statements first
            if self.is_control_flow_keyword() {
                self.parse_control_flow_statement();
                continue;
            }

            // Try built-in library statements
            if self.is_library_statement_keyword() {
                self.parse_library_statement();
                continue;
            }

            // Try array statements
            if self.is_variable_declaration_keyword() {
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
                // Whitespace, newlines, and comments - consume directly FIRST
                // This must be checked before is_at_procedure_call to avoid
                // treating REM comments as procedure calls
                Some(
                    Token::Whitespace
                    | Token::Newline
                    | Token::EndOfLineComment
                    | Token::RemComment,
                ) => {
                    self.consume_token();
                }
                // Variable declarations: Dim/Private/Public/Const/Static
                Some(
                    Token::DimKeyword
                    | Token::PrivateKeyword
                    | Token::PublicKeyword
                    | Token::ConstKeyword
                    | Token::StaticKeyword,
                ) => {
                    self.parse_dim();
                }
                // Anything else - check if it's a label, assignment, procedure call, or unknown
                _ => {
                    // Check if this is a label (identifier followed by colon)
                    if self.is_at_label() {
                        self.parse_label_statement();
                    // Check for Let statement (optional assignment keyword)
                    } else if self.at_token(Token::LetKeyword)
                        || (self.at_token(Token::Whitespace)
                            && self.peek_next_keyword() == Some(Token::LetKeyword))
                    {
                        self.parse_let_statement();
                    // Check if this looks like an assignment statement (identifier = expression)
                    } else if self.is_at_assignment() {
                        self.parse_assignment_statement();
                    // Check if this looks like a procedure call (identifier without assignment)
                    } else if self.is_at_procedure_call() {
                        self.parse_procedure_call();
                    } else {
                        self.consume_token_as_unknown();
                    }
                }
            }
        }
        self.builder.finish_node(); // StatementList
    }
}

#[cfg(test)]
mod tests {
    use super::Parser;
    use crate::parsers::cst::{
        ControlKind, Creatable, Exposed, FormRoot, NameSpace, ObjectReference, PreDeclaredID,
    };
    use crate::*;

    use assert_matches::assert_matches;

    #[test]
    fn parse_single_quote_comment() {
        let code = "' This is a comment\nSub Main()\n";

        let mut source_stream = SourceStream::new("test.bas", code);
        let result = tokenize(&mut source_stream);
        let (token_stream_opt, _failures) = result.unpack();

        let token_stream = token_stream_opt.expect("Tokenization failed");
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
        assert_eq!(first.kind(), SyntaxKind::EndOfLineComment);
        assert!(first.is_token());
    }

    #[test]
    fn syntax_kind_conversions() {
        // Test keyword conversions
        assert_eq!(
            SyntaxKind::from(Token::FunctionKeyword),
            SyntaxKind::FunctionKeyword
        );
        assert_eq!(SyntaxKind::from(Token::IfKeyword), SyntaxKind::IfKeyword);
        assert_eq!(SyntaxKind::from(Token::ForKeyword), SyntaxKind::ForKeyword);

        // Test operators
        assert_eq!(
            SyntaxKind::from(Token::AdditionOperator),
            SyntaxKind::AdditionOperator
        );
        assert_eq!(
            SyntaxKind::from(Token::EqualityOperator),
            SyntaxKind::EqualityOperator
        );

        // Test literals
        assert_eq!(
            SyntaxKind::from(Token::StringLiteral),
            SyntaxKind::StringLiteral
        );
        assert_eq!(
            SyntaxKind::from(Token::IntegerLiteral),
            SyntaxKind::IntegerLiteral
        );
        assert_eq!(
            SyntaxKind::from(Token::LongLiteral),
            SyntaxKind::LongLiteral
        );
        assert_eq!(
            SyntaxKind::from(Token::SingleLiteral),
            SyntaxKind::SingleLiteral
        );
        assert_eq!(
            SyntaxKind::from(Token::DoubleLiteral),
            SyntaxKind::DoubleLiteral
        );
        assert_eq!(
            SyntaxKind::from(Token::DateTimeLiteral),
            SyntaxKind::DateLiteral
        );
    }

    #[test]
    fn parse_empty_stream() {
        let source = "";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 0);
    }

    #[test]
    fn parse_rem_comment() {
        let source = "REM This is a REM comment\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

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
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        // Should have 5 children: EndOfLineComment, Newline, RemComment, Newline, SubStatement
        assert_eq!(cst.child_count(), 5);
        assert!(cst.text().contains("' Single quote comment"));
        assert!(cst.text().contains("REM REM comment"));

        // Use navigation methods
        let children = cst.children();
        assert_eq!(children[0].kind(), SyntaxKind::EndOfLineComment);
        assert_eq!(children[1].kind(), SyntaxKind::Newline);
        assert_eq!(children[2].kind(), SyntaxKind::RemComment);
        assert_eq!(children[3].kind(), SyntaxKind::Newline);
        assert_eq!(children[4].kind(), SyntaxKind::SubStatement);

        assert!(cst.contains_kind(SyntaxKind::EndOfLineComment));
        assert!(cst.contains_kind(SyntaxKind::RemComment));
    }

    #[test]
    fn cst_with_comments() {
        let source = "' This is a comment\nSub Main()\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Now has 3 children: comment token, newline token, SubStatement
        assert_eq!(cst.child_count(), 3);
        assert!(cst.text().contains("' This is a comment"));
        assert!(cst.text().contains("Sub Main()"));
    }

    #[test]
    fn cst_serializable_tree() {
        let source = "Sub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        // Convert to serializable format
        let serializable = cst.to_serializable();

        // Verify structure
        assert_eq!(serializable.root.kind(), SyntaxKind::Root);
        assert!(!serializable.root.is_token());
        assert_eq!(serializable.root.children().len(), 1);
        assert_eq!(
            serializable.root.children()[0].kind(),
            SyntaxKind::SubStatement
        );

        // Can be used with insta for snapshot testing:
        // insta::assert_yaml_snapshot!(serializable);
    }

    #[test]
    fn cst_serializable_with_insta() {
        let source = "Dim x As Integer\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let serializable = cst.to_serializable();

        // Example of using with insta (commented out to not create snapshot files in normal test runs)
        // insta::assert_yaml_snapshot!(serializable);

        // Verify it's serializable by checking structure
        assert!(!serializable.root.children().is_empty());
    }

    // Phase 1 Tests: Parser Modes and Constructors

    #[test]
    fn parser_mode_full_cst_default() {
        let source = "Sub Test()\nEnd Sub\n";
        let mut stream = SourceStream::new("test.bas".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");

        let parser = Parser::new(token_stream);
        // Verify parser was created successfully
        assert_eq!(parser.pos, 0);
    }

    #[test]
    fn parser_mode_direct_extraction() {
        let source = "Sub Test()\nEnd Sub\n";
        let mut stream = SourceStream::new("test.bas".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let parser = Parser::new_direct_extraction(tokens, 0);
        assert_eq!(parser.pos, 0);
    }

    #[test]
    fn parser_constructors_preserve_tokens() {
        let source = "VERSION 5.00\n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");

        let tokens_vec = token_stream.into_tokens();
        let token_count = tokens_vec.len();

        let parser = Parser::new_direct_extraction(tokens_vec, 0);
        assert_eq!(parser.tokens.len(), token_count);
        assert!(parser.tokens[0].1 == Token::VersionKeyword);
    }

    #[test]
    fn parser_new_with_position() {
        let source = "VERSION 5.00\nSub Test()\nEnd Sub\n";
        let mut stream = SourceStream::new("test.bas".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        // Create parser starting at position 3 (after VERSION keyword, whitespace, and version number)
        let parser = Parser::new_direct_extraction(tokens, 3);
        assert_eq!(parser.pos, 3);
    }

    // Phase 2 Tests: VERSION Parsing

    #[test]
    fn parse_version_direct_with_version() {
        let source = "VERSION 5.00\nSub Test()\nEnd Sub\n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (version_opt, failures) = parser.parse_version_direct().unpack();

        assert!(version_opt.is_some());
        let version = version_opt.unwrap();
        assert_eq!(version.major, 5);
        assert_eq!(version.minor, 0);
        assert!(failures.is_empty());
    }

    #[test]
    fn parse_version_direct_without_version() {
        let source = "Sub Test()\nEnd Sub\n";
        let mut stream = SourceStream::new("test.bas".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (version_opt, failures) = parser.parse_version_direct().unpack();

        assert!(version_opt.is_none());
        assert!(failures.is_empty());
    }

    #[test]
    fn parse_version_direct_with_class_keyword() {
        let source = "VERSION 1.0 CLASS\nSub Test()\nEnd Sub\n";
        let mut stream = SourceStream::new("test.cls".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (version_opt, failures) = parser.parse_version_direct().unpack();

        assert!(version_opt.is_some());
        let version = version_opt.unwrap();
        assert_eq!(version.major, 1);
        assert_eq!(version.minor, 0);
        assert!(failures.is_empty());
    }

    #[test]
    fn parse_version_direct_version_100() {
        let source = "VERSION 1.00\n";
        let mut stream = SourceStream::new("test.cls".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (version_opt, _failures) = parser.parse_version_direct().unpack();

        assert!(version_opt.is_some());
        let version = version_opt.unwrap();
        assert_eq!(version.major, 1);
        assert_eq!(version.minor, 0);
    }

    #[test]
    fn parse_version_direct_with_whitespace() {
        let source = "  VERSION   5.00  \nSub Test()\n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (version_opt, _failures) = parser.parse_version_direct().unpack();

        assert!(version_opt.is_some());
        let version = version_opt.unwrap();
        assert_eq!(version.major, 5);
        assert_eq!(version.minor, 0);
    }

    #[test]
    fn parse_version_direct_position_advances() {
        let source = "VERSION 5.00\nBegin VB.Form Form1\nEnd\n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let initial_pos = parser.pos;
        let _result = parser.parse_version_direct();

        // Position should have advanced past VERSION statement
        assert!(parser.pos > initial_pos);

        // Should now be positioned at Begin keyword
        assert_eq!(parser.current_token(), Some(&Token::BeginKeyword));
    }

    #[test]
    fn parse_version_direct_accuracy() {
        let test_cases = vec![
            ("VERSION 5.00\n", Some((5, 0))),
            ("VERSION 1.0\n", Some((1, 0))),
            ("VERSION 6.00 CLASS\n", Some((6, 0))),
            ("VERSION 4.00\n", Some((4, 0))),
            ("Sub Test()\n", None), // No VERSION
        ];

        for (source, expected) in test_cases {
            let mut stream = SourceStream::new("test.vb".to_string(), source);
            let (token_stream_opt, _) = tokenize(&mut stream).unpack();
            let token_stream = token_stream_opt.expect("Tokenization failed");
            let tokens = token_stream.into_tokens();

            let mut parser = Parser::new_direct_extraction(tokens, 0);
            let (version_opt, _failures) = parser.parse_version_direct().unpack();

            match expected {
                Some((major, minor)) => {
                    assert!(version_opt.is_some(), "Expected version for: {source}");
                    let version = version_opt.unwrap();
                    assert_eq!(version.major, major, "Major mismatch for: {source}");
                    assert_eq!(version.minor, minor, "Minor mismatch for: {source}");
                }
                None => {
                    assert!(version_opt.is_none(), "Expected no version for: {source}");
                }
            }
        }
    }

    // Phase 3 Tests: Core Control Extraction

    #[test]
    fn parse_control_type_direct_simple() {
        let source = "VB.Form Form1\n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let control_type = parser.parse_control_type_direct();

        assert_eq!(control_type, "VB.Form");
    }

    #[test]
    fn parse_control_name_direct_simple() {
        let source = "Form1 \n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let control_name = parser.parse_control_name_direct();

        assert_eq!(control_name, "Form1");
    }

    #[test]
    fn parse_property_direct_simple() {
        let source = "Caption = \"Hello World\"\n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let property = parser.parse_property_direct();

        assert!(property.is_some());
        let (key, value) = property.unwrap();
        assert_eq!(key, "Caption");
        assert_eq!(value, "\"Hello World\"");
    }

    #[test]
    fn parse_properties_block_to_control_simple_form() {
        let source = r#"Begin VB.Form Form1
   Caption = "Test Form"
   ClientHeight = 3000
   ClientWidth = 4000
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (control_opt, failures) = parser.parse_properties_block_to_form_root().unpack();

        assert!(failures.is_empty(), "Expected no failures");
        assert!(control_opt.is_some(), "Expected form root to be parsed");

        let form_root = control_opt.unwrap();
        assert_eq!(form_root.name(), "Form1");

        // Verify it's a Form
        assert!(form_root.is_form());
    }

    #[test]
    fn parse_properties_block_to_control_command_button() {
        let source = r#"Begin VB.CommandButton Command1
   Caption = "Click Me"
   Height = 495
   Width = 1215
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (control_opt, failures) = parser.parse_properties_block_to_control().unpack();

        assert!(failures.is_empty());
        assert!(control_opt.is_some());

        let control = control_opt.unwrap();
        assert_eq!(control.name(), "Command1");
        assert_matches!(control.kind(), ControlKind::CommandButton { .. });
    }

    #[test]
    fn parse_properties_block_to_control_textbox() {
        let source = r#"Begin VB.TextBox Text1
   Text = "Initial Text"
   Height = 300
   Width = 2000
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (control_opt, _failures) = parser.parse_properties_block_to_control().unpack();

        assert!(control_opt.is_some());
        let control = control_opt.unwrap();
        assert_eq!(control.name(), "Text1");
        assert_matches!(control.kind(), ControlKind::TextBox { .. });
    }

    #[test]
    fn parse_properties_block_without_begin() {
        let source = "Caption = \"Test\"\nEnd\n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (control_opt, _failures) = parser.parse_properties_block_to_control().unpack();

        // Should return None when BEGIN is missing
        assert!(control_opt.is_none());
    }

    // Phase 4 Tests: Nested Controls and Property Groups

    #[test]
    fn parse_form_with_nested_control() {
        let source = r#"Begin VB.Form Form1
   Caption = "Main Form"
   Begin VB.CommandButton Command1
      Caption = "Click Me"
      Height = 400
   End
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (form_root_opt, failures) = parser.parse_properties_block_to_form_root().unpack();

        assert!(failures.is_empty(), "Should have no failures");
        assert!(form_root_opt.is_some());
        let form_root = form_root_opt.unwrap();
        assert_eq!(form_root.name(), "Form1");

        // Check form has child controls
        if let FormRoot::Form(form) = &form_root {
            assert_eq!(form.controls.len(), 1);
            assert_eq!(form.controls[0].name(), "Command1");
            assert_matches!(form.controls[0].kind(), ControlKind::CommandButton { .. });
        } else {
            panic!("Expected Form");
        }
    }

    #[test]
    fn parse_frame_with_multiple_nested_controls() {
        let source = r#"Begin VB.Frame Frame1
   Caption = "Options"
   Begin VB.CheckBox Check1
      Caption = "Option 1"
   End
   Begin VB.CheckBox Check2
      Caption = "Option 2"
   End
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (control_opt, failures) = parser.parse_properties_block_to_control().unpack();

        assert!(failures.is_empty());
        assert!(control_opt.is_some());
        let control = control_opt.unwrap();
        assert_eq!(control.name(), "Frame1");

        // Check frame has 2 child checkboxes
        if let ControlKind::Frame { controls, .. } = control.kind() {
            assert_eq!(controls.len(), 2);
            assert_eq!(controls[0].name(), "Check1");
            assert_eq!(controls[1].name(), "Check2");
        } else {
            panic!("Expected Frame control kind");
        }
    }

    #[test]
    fn parse_control_with_property_group() {
        let source = r#"Begin VB.CommandButton Command1
   Caption = "Button"
   BeginProperty Font
      Name = "Arial"
      Size = 12
   EndProperty
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (control_opt, failures) = parser.parse_properties_block_to_control().unpack();

        assert!(failures.is_empty());
        assert!(control_opt.is_some());
        let control = control_opt.unwrap();
        assert_eq!(control.name(), "Command1");

        // Check for property - CommandButton should have parsed successfully
        // Property groups are stored in Custom control kind, not specific control types
        assert_matches!(control.kind(), ControlKind::CommandButton { .. });
    }

    #[test]
    fn parse_custom_control_with_property_group() {
        let source = r#"Begin MSComctlLib.TreeView TreeView1
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
      Name = "MS Sans Serif"
      Size = 8.25
      Charset = 0
   EndProperty
   Caption = "Tree"
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (control_opt, failures) = parser.parse_properties_block_to_control().unpack();

        assert!(failures.is_empty());
        assert!(control_opt.is_some());
        let control = control_opt.unwrap();
        assert_eq!(control.name(), "TreeView1");

        // Check for Custom control with property groups
        if let ControlKind::Custom {
            property_groups, ..
        } = control.kind()
        {
            assert_eq!(property_groups.len(), 1);
            assert_eq!(property_groups[0].name, "Font");
            assert!(property_groups[0].guid.is_some());
        } else {
            panic!("Expected Custom control kind");
        }
    }

    // Phase 5 Tests: Direct Object Parsing
    #[test]
    fn parse_simple_object_statement() {
        let source = r#"Object = "{12345678-1234-1234-1234-123456789ABC}#1.0#0"; "MyLib.dll""#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let objects = parser.parse_objects_direct();

        assert_eq!(objects.len(), 1);
        match &objects[0] {
            ObjectReference::Compiled {
                uuid,
                version,
                unknown1,
                file_name,
            } => {
                assert_eq!(
                    uuid.to_string().to_uppercase(),
                    "12345678-1234-1234-1234-123456789ABC"
                );
                assert_eq!(version, "1.0");
                assert_eq!(unknown1, "0");
                assert_eq!(file_name, "MyLib.dll");
            }
            ObjectReference::Project { .. } => {
                panic!("Expected Compiled object reference")
            }
        }
    }

    #[test]
    fn parse_multiple_object_statements() {
        let source = r#"Object = "{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}#1.0#0"; "Lib1.dll"
Object = "{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}#2.0#1"; "Lib2.ocx"
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let objects = parser.parse_objects_direct();

        assert_eq!(objects.len(), 2);

        match &objects[0] {
            ObjectReference::Compiled { file_name, .. } => {
                assert_eq!(file_name, "Lib1.dll");
            }
            ObjectReference::Project { .. } => {
                panic!("Expected Compiled object reference")
            }
        }

        match &objects[1] {
            ObjectReference::Compiled { file_name, .. } => {
                assert_eq!(file_name, "Lib2.ocx");
            }
            ObjectReference::Project { .. } => {
                panic!("Expected Compiled object reference")
            }
        }
    }

    #[test]
    fn parse_embedded_object_statement() {
        let source = r#"Object = *\G{87654321-4321-4321-4321-CBA987654321}#3.0#5; "Embedded.ocx""#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let objects = parser.parse_objects_direct();

        assert_eq!(objects.len(), 1);
        match &objects[0] {
            ObjectReference::Compiled {
                uuid,
                version,
                file_name,
                ..
            } => {
                assert_eq!(
                    uuid.to_string().to_uppercase(),
                    "87654321-4321-4321-4321-CBA987654321"
                );
                assert_eq!(version, "3.0");
                assert_eq!(file_name, "Embedded.ocx");
            }
            ObjectReference::Project { .. } => {
                panic!("Expected Compiled object reference")
            }
        }
    }

    #[test]
    fn parse_nested_property_groups() {
        use either::Either;

        let source = r#"Begin Custom.Control Ctrl1
   BeginProperty Outer
      Value1 = "Test"
      BeginProperty Inner
         Value2 = "Nested"
      EndProperty
   EndProperty
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (control_opt, failures) = parser.parse_properties_block_to_control().unpack();

        assert!(failures.is_empty());
        assert!(control_opt.is_some());
        let control = control_opt.unwrap();

        // Check for nested property groups
        if let ControlKind::Custom {
            property_groups, ..
        } = control.kind()
        {
            assert_eq!(property_groups.len(), 1);
            assert_eq!(property_groups[0].name, "Outer");

            // Check for nested group

            if let Some(Either::Right(inner)) = property_groups[0].properties.get("Inner") {
                assert_eq!(inner.name, "Inner");
            } else {
                panic!("Expected nested Inner property group");
            }
        } else {
            panic!("Expected Custom control kind");
        }
    }

    #[test]
    fn parse_deeply_nested_controls() {
        let source = r#"Begin VB.Form Form1
   Caption = "Outer"
   Begin VB.PictureBox Picture1
      Begin VB.Frame Frame1
         Begin VB.Label Label1
            Caption = "Deep"
         End
      End
   End
End
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let (form_root_opt, failures) = parser.parse_properties_block_to_form_root().unpack();

        assert!(failures.is_empty());
        assert!(form_root_opt.is_some());
        let form_root = form_root_opt.unwrap();

        // Verify deep nesting: Form > PictureBox > Frame > Label
        if let FormRoot::Form(form) = &form_root {
            assert_eq!(form.controls.len(), 1);
            if let ControlKind::PictureBox { controls, .. } = form.controls[0].kind() {
                assert_eq!(controls.len(), 1);
                if let ControlKind::Frame { controls, .. } = controls[0].kind() {
                    assert_eq!(controls.len(), 1);
                    assert_eq!(controls[0].name(), "Label1");
                } else {
                    panic!("Expected Frame");
                }
            } else {
                panic!("Expected PictureBox");
            }
        } else {
            panic!("Expected Form");
        }
    }

    // Phase 6 Tests: Direct Attribute Parsing
    #[test]
    fn parse_simple_string_attribute() {
        let source = r#"Attribute VB_Name = "Form1"
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let attrs = parser.parse_attributes_direct();

        assert_eq!(attrs.name, "Form1");
        assert_eq!(attrs.global_name_space, NameSpace::Local);
        assert_eq!(attrs.creatable, Creatable::True);
        assert_eq!(attrs.predeclared_id, PreDeclaredID::False);
        assert_eq!(attrs.exposed, Exposed::False);
        assert_eq!(attrs.description, None);
    }

    #[test]
    fn parse_boolean_attributes() {
        let source = r"Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let attrs = parser.parse_attributes_direct();

        assert_eq!(attrs.global_name_space, NameSpace::Local);
        assert_eq!(attrs.creatable, Creatable::True);
        assert_eq!(attrs.predeclared_id, PreDeclaredID::True);
        assert_eq!(attrs.exposed, Exposed::False);
    }

    #[test]
    fn parse_numeric_attribute() {
        let source = r"Attribute VB_PredeclaredId = -1
";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let attrs = parser.parse_attributes_direct();

        // -1 is truthy in VB6, so should be parsed as true
        assert_eq!(attrs.predeclared_id, PreDeclaredID::True);
    }

    #[test]
    fn parse_multiple_attributes() {
        let source = r#"Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "This is a test form"
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let attrs = parser.parse_attributes_direct();

        assert_eq!(attrs.name, "MyForm");
        assert_eq!(attrs.global_name_space, NameSpace::Local);
        assert_eq!(attrs.creatable, Creatable::False);
        assert_eq!(attrs.predeclared_id, PreDeclaredID::True);
        assert_eq!(attrs.exposed, Exposed::False);
        assert_eq!(attrs.description, Some("This is a test form".to_string()));
    }

    #[test]
    fn parse_ext_key_attributes() {
        let source = r#"Attribute VB_Name = "Form1"
Attribute VB_Ext_KEY = "CustomKey" ,"CustomValue"
Attribute VB_Description = "Test"
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let attrs = parser.parse_attributes_direct();

        assert_eq!(attrs.name, "Form1");
        assert_eq!(attrs.description, Some("Test".to_string()));
        assert_eq!(attrs.ext_key.len(), 1);
        assert!(attrs.ext_key.contains_key("VB_Ext_KEY"));
    }

    #[test]
    fn parse_empty_attributes() {
        let source = r"";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let attrs = parser.parse_attributes_direct();

        assert_eq!(attrs.name, "");
        assert_eq!(attrs.global_name_space, NameSpace::Local);
        assert_eq!(attrs.creatable, Creatable::True);
        assert_eq!(attrs.predeclared_id, PreDeclaredID::False);
        assert_eq!(attrs.exposed, Exposed::False);
        assert_eq!(attrs.description, None);
        assert!(attrs.ext_key.is_empty());
    }

    #[test]
    fn parse_resource_reference_property() {
        let source = r#"Caption = $"Gradient.frx":0000
"#;
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let mut parser = Parser::new_direct_extraction(tokens, 0);
        let property = parser.parse_property_direct();

        assert!(property.is_some());
        let (key, value) = property.unwrap();
        assert_eq!(key, "Caption");
        assert_eq!(value, r#"$"Gradient.frx":0000"#);
    }
}
