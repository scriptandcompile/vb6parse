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
mod assignment;
mod attribute_statements;
mod conditionals;
mod controlflow;
mod declarations;
mod deftype_statements;
mod enum_statements;
mod for_statements;
mod function_statements;
mod helpers;
mod if_statements;
mod library_statements;
mod loop_statements;
mod navigation;
mod object_statements;
mod option_statements;
mod parameters;
mod property_statements;
mod select_statements;
mod sub_statements;
mod variable_declarations;

// Re-export navigation types
pub use navigation::CstNode;

/// A serializable representation of the CST for snapshot testing.
///
/// This struct wraps the tree structure in a way that can be serialized
/// with serde, making it suitable for use with snapshot testing tools like insta.
#[derive(Debug, Clone, PartialEq, Eq, serde::Serialize)]
pub struct SerializableTree {
    /// The root node of the tree
    pub root: CstNode,
    /// The complete text content
    pub text: String,
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
    /// let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
    /// let serializable = cst.to_serializable();
    ///
    /// // Can now be used with insta::assert_yaml_snapshot!
    /// ```
    pub fn to_serializable(&self) -> SerializableTree {
        SerializableTree {
            root: self.to_root_node(),
            text: self.text(),
        }
    }

    /// Convert the internal rowan tree to a root CstNode.
    fn to_root_node(&self) -> CstNode {
        CstNode {
            kind: SyntaxKind::Root,
            text: self.text(),
            is_token: false,
            children: self.children(),
        }
    }
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
                    // Peek ahead to check if this is Option Base or Option Compare
                    if let Some(VB6Token::BaseKeyword) = self.peek_next_keyword() {
                        self.parse_option_base_statement();
                    } else if let Some(VB6Token::CompareKeyword) = self.peek_next_keyword() {
                        self.parse_option_compare_statement();
                    } else {
                        self.parse_option_statement();
                    }
                }
                // DefType statements: DefInt, DefLng, DefStr, etc.
                Some(VB6Token::DefBoolKeyword)
                | Some(VB6Token::DefByteKeyword)
                | Some(VB6Token::DefIntKeyword)
                | Some(VB6Token::DefLngKeyword)
                | Some(VB6Token::DefCurKeyword)
                | Some(VB6Token::DefSngKeyword)
                | Some(VB6Token::DefDblKeyword)
                | Some(VB6Token::DefDecKeyword)
                | Some(VB6Token::DefDateKeyword)
                | Some(VB6Token::DefStrKeyword)
                | Some(VB6Token::DefObjKeyword)
                | Some(VB6Token::DefVarKeyword) => {
                    self.parse_deftype_statement();
                }
                // Declare statement: Declare Sub/Function Name Lib "..."
                Some(VB6Token::DeclareKeyword) => {
                    self.parse_declare_statement();
                }
                // Event statement: Event Name(...)
                Some(VB6Token::EventKeyword) => {
                    self.parse_event_statement();
                }
                // Implements statement: Implements InterfaceName
                Some(VB6Token::ImplementsKeyword) => {
                    self.parse_implements_statement();
                }
                // Enum statement: Enum Name ... End Enum
                Some(VB6Token::EnumKeyword) => {
                    self.parse_enum_statement();
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
                    // Look ahead to see if this is a function/sub/property/enum declaration
                    // Peek at the next 2 keywords to handle cases like "Public Static Function"
                    let next_keywords: Vec<_> = self
                        .peek_next_count_keywords(NonZeroUsize::new(2).unwrap())
                        .collect();

                    match next_keywords.as_slice() {
                        // Direct: Public/Private/Friend Function, Sub, Property, Enum, Declare, or Event
                        [VB6Token::FunctionKeyword, ..] => self.parse_function_statement(), // Function
                        [VB6Token::SubKeyword, ..] => self.parse_sub_statement(),           // Sub
                        [VB6Token::PropertyKeyword, ..] => self.parse_property_statement(), // Property
                        [VB6Token::DeclareKeyword, ..] => self.parse_declare_statement(), // Declare
                        [VB6Token::EnumKeyword, ..] => self.parse_enum_statement(),       // Enum
                        [VB6Token::EventKeyword, ..] => self.parse_event_statement(),     // Event
                        [VB6Token::ImplementsKeyword, ..] => self.parse_implements_statement(), // Implements
                        // With Static: Public/Private/Friend Static Function, Sub, or Property
                        [VB6Token::StaticKeyword, VB6Token::FunctionKeyword] => {
                            self.parse_function_statement()
                        }
                        [VB6Token::StaticKeyword, VB6Token::SubKeyword] => {
                            self.parse_sub_statement()
                        }
                        [VB6Token::StaticKeyword, VB6Token::PropertyKeyword] => {
                            self.parse_property_statement()
                        }
                        // Anything else is a declaration
                        _ => self.parse_dim(),
                    };
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
                    } else if self.at_token(VB6Token::LetKeyword) {
                        self.parse_let_statement();
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

    /// Check if the current token is a control flow keyword.
    fn is_control_flow_keyword(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::IfKeyword)
                | Some(VB6Token::SelectKeyword)
                | Some(VB6Token::ForKeyword)
                | Some(VB6Token::DoKeyword)
                | Some(VB6Token::GotoKeyword)
                | Some(VB6Token::GoSubKeyword)
                | Some(VB6Token::ReturnKeyword)
                | Some(VB6Token::ExitKeyword)
                | Some(VB6Token::OnKeyword)
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
            Some(VB6Token::GoSubKeyword) => {
                self.parse_gosub_statement();
            }
            Some(VB6Token::ReturnKeyword) => {
                self.parse_return_statement();
            }
            Some(VB6Token::ExitKeyword) => {
                self.parse_exit_statement();
            }
            Some(VB6Token::OnKeyword) => {
                // Look ahead to distinguish between On Error, On GoTo, and On GoSub
                if let Some(VB6Token::ErrorKeyword) = self.peek_next_keyword() {
                    self.parse_on_error_statement();
                } else {
                    // Need to scan ahead to find GoTo or GoSub keyword
                    // to distinguish between On GoTo and On GoSub
                    use std::num::NonZeroUsize;
                    let keywords: Vec<VB6Token> = self
                        .peek_next_count_keywords(NonZeroUsize::new(20).unwrap())
                        .collect();

                    let has_goto = keywords.iter().any(|t| *t == VB6Token::GotoKeyword);
                    let has_gosub = keywords.iter().any(|t| *t == VB6Token::GoSubKeyword);

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
            Some(VB6Token::ReDimKeyword) | Some(VB6Token::EraseKeyword)
        )
    }

    /// Dispatch array statement parsing to the appropriate parser.
    fn parse_array_statement(&mut self) {
        match self.current_token() {
            Some(VB6Token::ReDimKeyword) => {
                self.parse_redim_statement();
            }
            Some(VB6Token::EraseKeyword) => {
                self.parse_erase_statement();
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
                    // Check for Let statement (optional assignment keyword)
                    } else if self.at_token(VB6Token::LetKeyword) {
                        self.parse_let_statement();
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

#[cfg(test)]
mod test {
    use crate::*;

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

    #[test]
    fn syntax_kind_conversions() {
        use crate::language::VB6Token;

        // Test keyword conversions
        assert_eq!(
            SyntaxKind::from(VB6Token::FunctionKeyword),
            SyntaxKind::FunctionKeyword
        );
        assert_eq!(SyntaxKind::from(VB6Token::IfKeyword), SyntaxKind::IfKeyword);
        assert_eq!(
            SyntaxKind::from(VB6Token::ForKeyword),
            SyntaxKind::ForKeyword
        );

        // Test operators
        assert_eq!(
            SyntaxKind::from(VB6Token::AdditionOperator),
            SyntaxKind::AdditionOperator
        );
        assert_eq!(
            SyntaxKind::from(VB6Token::EqualityOperator),
            SyntaxKind::EqualityOperator
        );

        // Test literals
        assert_eq!(
            SyntaxKind::from(VB6Token::StringLiteral),
            SyntaxKind::StringLiteral
        );
        assert_eq!(SyntaxKind::from(VB6Token::Number), SyntaxKind::Number);
    }

    #[test]
    fn parse_empty_stream() {
        let source = "";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 0);
    }

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

    #[test]
    fn cst_serializable_tree() {
        let source = "Sub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        // Convert to serializable format
        let serializable = cst.to_serializable();

        // Verify structure
        assert_eq!(serializable.text, source);
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
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let serializable = cst.to_serializable();

        // Example of using with insta (commented out to not create snapshot files in normal test runs)
        // insta::assert_yaml_snapshot!(serializable);

        // Verify it's serializable by checking structure
        assert!(serializable.root.children.len() > 0);
    }
}
