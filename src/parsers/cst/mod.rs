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
//! use vb6parse::tokenstream::TokenStream;
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
//! # Design Principles
//!
//! 1. **No rowan types exposed**: All public APIs use custom types that don't expose rowan.
//! 2. **Complete representation**: The CST includes all tokens, including whitespace and comments.
//! 3. **Efficient**: Uses rowan's red-green tree architecture for memory efficiency.
//! 4. **Type-safe**: All syntax kinds are represented as a Rust enum for compile-time safety.

use std::num::NonZeroUsize;

use crate::language::Token;
use crate::parsers::SyntaxKind;
use crate::tokenize::tokenize;
use crate::tokenstream::TokenStream;
use crate::CodeErrorKind;
use crate::ParseResult;
use crate::SourceFile;
use crate::SourceStream;

use rowan::{GreenNode, GreenNodeBuilder, Language};
use serde::Serialize;

// Submodules for organized CST parsing
mod assignment;
mod attribute_statements;
mod controlflow;
mod declarations;
mod deftype_statements;
mod enum_statements;
mod expressions;
mod for_statements;
mod function_statements;
mod helpers;
mod if_statements;
mod library_functions;
mod library_statements;
mod loop_statements;
mod navigation;
mod object_statements;
mod option_statements;
mod parameters;
mod properties;
mod property_statements;
mod select_statements;
mod sub_statements;
mod type_statements;
mod variable_declarations;

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
    pub fn from_source(source_file: &SourceFile) -> ParseResult<'_, Self, CodeErrorKind> {
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
    pub fn from_text<S>(file_name: S, contents: &str) -> ParseResult<'_, Self, CodeErrorKind>
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
    fn to_root_node(&self) -> CstNode {
        CstNode {
            kind: SyntaxKind::Root,
            text: self.text(),
            is_token: false,
            children: self.children(),
        }
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
/// use vb6parse::tokenstream::TokenStream;
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

/// Parser mode determines whether to build a full CST or extract structures directly
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ParserMode {
    /// Build full CST including VERSION, control, attributes, and code
    FullCst,
    /// Extract structures directly without CST building
    DirectExtraction,
    /// Extract VERSION/control/attributes, then build CST for code only
    Hybrid,
}

/// Internal parser state for building the CST
pub(crate) struct Parser<'a> {
    pub(crate) tokens: Vec<(&'a str, Token)>,
    pub(crate) pos: usize,
    builder: GreenNodeBuilder<'static>,
    parsing_header: bool,
    mode: ParserMode,
}

impl<'a> Parser<'a> {
    fn new(token_stream: TokenStream<'a>) -> Self {
        Parser {
            tokens: token_stream.into_tokens(),
            pos: 0,
            builder: GreenNodeBuilder::new(),
            parsing_header: true,
            mode: ParserMode::FullCst,
        }
    }

    /// Create parser for direct extraction mode (control-only parsing)
    pub(crate) fn new_direct_extraction(tokens: Vec<(&'a str, Token)>, pos: usize) -> Self {
        Parser {
            tokens,
            pos,
            builder: GreenNodeBuilder::new(),
            parsing_header: true,
            mode: ParserMode::DirectExtraction,
        }
    }

    /// Create parser for hybrid mode (`FormFile` optimization)
    pub(crate) fn new_hybrid(token_stream: TokenStream<'a>) -> Self {
        Parser {
            tokens: token_stream.into_tokens(),
            pos: 0,
            builder: GreenNodeBuilder::new(),
            parsing_header: true,
            mode: ParserMode::Hybrid,
        }
    }

    // ==================== Direct Extraction Helpers ====================
    // These methods support direct extraction without CST building

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
    ///
    /// # Example
    ///
    /// ```ignore
    /// let mut parser = Parser::new_direct_extraction(tokens, 0);
    /// let result = parser.parse_version_direct();
    /// if let Some(version) = result.result {
    ///     println!(\"Version: {}.{}\", version.major, version.minor);
    /// }
    /// ```
    pub(crate) fn parse_version_direct(
        &mut self,
    ) -> crate::ParseResult<'a, crate::parsers::FileFormatVersion, crate::errors::FormErrorKind>
    {
        use crate::parsers::FileFormatVersion;

        self.skip_whitespace();

        // Check if VERSION keyword is present
        if !self.at_token(Token::VersionKeyword) {
            return crate::ParseResult::new(None, Vec::new());
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

        crate::ParseResult::new(version_result, Vec::new())
    }

    // ==================== Core Control Extraction Methods ====================

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
        let mut value_parts = Vec::new();
        while !self.is_at_end()
            && !self.at_token(Token::Newline)
            && !self.at_token(Token::ColonOperator)
        {
            if let Some((text, _)) = self.tokens.get(self.pos) {
                value_parts.push(text.to_string());
                self.consume_advance();
            } else {
                break;
            }
        }

        // Skip newline
        self.skip_whitespace_and_newlines();

        let value = value_parts.join("").trim().to_string();
        Some((key, value))
    }

    /// Build `ControlKind` from control type string and properties
    fn build_control_kind(
        control_type: &str,
        properties: crate::Properties,
        child_controls: Vec<crate::language::Control>,
        menus: Vec<crate::language::MenuControl>,
    ) -> crate::language::ControlKind {
        use crate::language::ControlKind;

        match control_type {
            "VB.Form" => ControlKind::Form {
                properties: properties.into(),
                controls: child_controls,
                menus,
            },
            "VB.MDIForm" => ControlKind::MDIForm {
                properties: properties.into(),
                controls: child_controls,
                menus,
            },
            "VB.CommandButton" => ControlKind::CommandButton {
                properties: properties.into(),
            },
            "VB.TextBox" => ControlKind::TextBox {
                properties: properties.into(),
            },
            "VB.Label" => ControlKind::Label {
                properties: properties.into(),
            },
            "VB.CheckBox" => ControlKind::CheckBox {
                properties: properties.into(),
            },
            "VB.ListBox" => ControlKind::ListBox {
                properties: properties.into(),
            },
            "VB.Timer" => ControlKind::Timer {
                properties: properties.into(),
            },
            "VB.Frame" => ControlKind::Frame {
                properties: properties.into(),
                controls: child_controls,
            },
            "VB.PictureBox" => ControlKind::PictureBox {
                properties: properties.into(),
                controls: child_controls,
            },
            _ => ControlKind::Custom {
                properties: properties.into(),
                property_groups: Vec::new(),
            },
        }
    }

    /// Parse properties block directly to Control without building CST
    /// Phase 3: Simple implementation without nesting support
    pub(crate) fn parse_properties_block_to_control(
        &mut self,
    ) -> crate::ParseResult<'a, crate::language::Control, crate::errors::FormErrorKind> {
        use crate::language::Control;
        use crate::Properties;

        self.skip_whitespace();

        // Expect BEGIN keyword
        if !self.at_token(Token::BeginKeyword) {
            return crate::ParseResult::new(None, Vec::new());
        }

        self.consume_advance(); // BEGIN
        self.skip_whitespace();

        // Parse control type (e.g., "VB.Form")
        let control_type = self.parse_control_type_direct();
        self.skip_whitespace();

        // Parse control name
        let control_name = self.parse_control_name_direct();
        self.skip_whitespace_and_newlines();

        // Parse properties until END
        let mut properties = Properties::new();

        while !self.is_at_end() && !self.at_token(Token::EndKeyword) {
            self.skip_whitespace();

            if self.at_token(Token::EndKeyword) {
                break;
            }

            // Parse property (Key = Value)
            if self.is_identifier() || self.at_keyword() {
                if let Some((key, value)) = self.parse_property_direct() {
                    properties.insert(&key, &value);
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

        // Build control kind (Phase 3: no child controls yet)
        let kind = Self::build_control_kind(&control_type, properties, Vec::new(), Vec::new());

        let control = Control {
            name: control_name,
            tag,
            index,
            kind,
        };

        crate::ParseResult::new(Some(control), Vec::new())
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
                // Whitespace and newlines - consume directly
                Some(
                    Token::Whitespace
                    | Token::Newline
                    | Token::EndOfLineComment
                    | Token::RemComment,
                ) => {
                    self.consume_token();
                }
                // Anything else - check if it's a statement, label, assignment, or unknown
                _ => {
                    // Try control flow statements
                    if self.is_control_flow_keyword() {
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
    fn is_control_flow_keyword(&self) -> bool {
        matches!(
            self.current_token(),
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
        match self.current_token() {
            Some(Token::IfKeyword) => {
                self.parse_if_statement();
            }
            Some(Token::SelectKeyword) => {
                self.parse_select_case_statement();
            }
            Some(Token::ForKeyword) => {
                // Peek ahead to see if next keyword is "Each"
                if let Some(Token::EachKeyword) = self.peek_next_keyword() {
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
                if let Some(Token::ErrorKeyword) = self.peek_next_keyword() {
                    self.parse_on_error_statement();
                } else {
                    // Need to scan ahead to find GoTo or GoSub keyword
                    // to distinguish between On GoTo and On GoSub
                    use std::num::NonZeroUsize;
                    let keywords: Vec<Token> = self
                        .peek_next_count_keywords(NonZeroUsize::new(20).unwrap())
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
    fn is_variable_declaration_keyword(&self) -> bool {
        matches!(
            self.current_token(),
            Some(Token::ReDimKeyword | Token::EraseKeyword)
        )
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
        match self.current_token() {
            Some(Token::ReDimKeyword) => {
                self.parse_redim_statement();
            }
            Some(Token::EraseKeyword) => {
                self.parse_erase_statement();
            }
            _ => {}
        }
    }

    /// Parse a code block, consuming tokens until a termination condition is met.
    ///
    /// This is a generic code block parser that can handle different termination conditions:
    /// - End Sub, End Function, End If, etc.
    /// - `ElseIf` or Else (for If statements)
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
                // Whitespace and newlines - consume directly
                Some(
                    Token::Whitespace
                    | Token::Newline
                    | Token::EndOfLineComment
                    | Token::RemComment,
                ) => {
                    self.consume_token();
                }
                // Anything else - check if it's a label, assignment, procedure call, or unknown
                _ => {
                    // Check if this is a label (identifier followed by colon)
                    if self.is_at_label() {
                        self.parse_label_statement();
                    // Check for Let statement (optional assignment keyword)
                    } else if self.at_token(Token::LetKeyword) {
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
        self.builder.finish_node(); // CodeBlock
    }
}

#[cfg(test)]
mod test {
    use super::{Parser, ParserMode};
    use crate::*;

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
        assert_eq!(first.kind, SyntaxKind::EndOfLineComment);
        assert!(first.is_token);
    }

    #[test]
    fn syntax_kind_conversions() {
        use crate::language::Token;

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
            SyntaxKind::from(Token::DateLiteral),
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
        assert_eq!(serializable.root.kind, SyntaxKind::Root);
        assert!(!serializable.root.is_token);
        assert_eq!(serializable.root.children.len(), 1);
        assert_eq!(serializable.root.children[0].kind, SyntaxKind::SubStatement);

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
        assert!(!serializable.root.children.is_empty());
    }

    // Phase 1 Tests: Parser Modes and Constructors

    #[test]
    fn parser_mode_full_cst_default() {
        let source = "Sub Test()\nEnd Sub\n";
        let mut stream = SourceStream::new("test.bas".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");

        let parser = Parser::new(token_stream);
        assert_eq!(parser.mode, ParserMode::FullCst);
    }

    #[test]
    fn parser_mode_direct_extraction() {
        let source = "Sub Test()\nEnd Sub\n";
        let mut stream = SourceStream::new("test.bas".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");
        let tokens = token_stream.into_tokens();

        let parser = Parser::new_direct_extraction(tokens, 0);
        assert_eq!(parser.mode, ParserMode::DirectExtraction);
        assert_eq!(parser.pos, 0);
    }

    #[test]
    fn parser_mode_hybrid() {
        let source = "VERSION 5.00\nBegin VB.Form Form1\nEnd\n";
        let mut stream = SourceStream::new("test.frm".to_string(), source);
        let (token_stream_opt, _) = tokenize(&mut stream).unpack();
        let token_stream = token_stream_opt.expect("Tokenization failed");

        let parser = Parser::new_hybrid(token_stream);
        assert_eq!(parser.mode, ParserMode::Hybrid);
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
        assert_eq!(parser.mode, ParserMode::DirectExtraction);
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
        let (control_opt, failures) = parser.parse_properties_block_to_control().unpack();

        assert!(failures.is_empty(), "Expected no failures");
        assert!(control_opt.is_some(), "Expected control to be parsed");

        let control = control_opt.unwrap();
        assert_eq!(control.name, "Form1");

        // Verify it's a Form
        assert!(matches!(
            control.kind,
            crate::language::ControlKind::Form { .. }
        ));
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
        assert_eq!(control.name, "Command1");
        assert!(matches!(
            control.kind,
            crate::language::ControlKind::CommandButton { .. }
        ));
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
        assert_eq!(control.name, "Text1");
        assert!(matches!(
            control.kind,
            crate::language::ControlKind::TextBox { .. }
        ));
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
}
