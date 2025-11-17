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
            .map(|child| Self::build_cst_node(child))
            .collect()
    }
    
    /// Recursively build a CstNode from a rowan NodeOrToken
    fn build_cst_node(node_or_token: rowan::NodeOrToken<rowan::SyntaxNode<VB6Language>, rowan::SyntaxToken<VB6Language>>) -> CstNode {
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
                // Function Procedure Syntax:
                //
                // [Public | Private | Friend] [ Static ] Function name [ ( arglist ) ] [ As type ]
                // [ statements ]
                // [ name = expression ]
                // [ Exit Function ]
                // [ statements ]
                // [ name = expression ]
                // End Function
                //
                Some(VB6Token::FunctionKeyword) => {
                    self.parse_function_statement();
                }
                // Variable declarations: Dim/Const
                // For Public/Private/Friend/Static, we need to look ahead to see if it's a 
                // function/sub declaration or a variable declaration
                Some(VB6Token::DimKeyword) 
                | Some(VB6Token::ConstKeyword) => {
                    self.parse_declaration();
                }
                // Public/Private/Friend/Static - could be function/sub or declaration
                Some(VB6Token::PrivateKeyword)
                | Some(VB6Token::PublicKeyword)
                | Some(VB6Token::FriendKeyword)
                | Some(VB6Token::StaticKeyword) => {
                    // Look ahead to see if this is a function/sub declaration
                    // Peek at the next 2 keywords to handle cases like "Public Static Function"
                    let next_keywords: Vec<_> = self.peek_next_count_keywords(2).collect();
                    
                    let is_function_or_sub = match next_keywords.as_slice() {
                        // Direct: Public/Private/Friend Function or Sub
                        [VB6Token::FunctionKeyword, ..] => Some(true),  // Function
                        [VB6Token::SubKeyword, ..] => Some(false),      // Sub
                        // With Static: Public/Private/Friend Static Function or Sub
                        [VB6Token::StaticKeyword, VB6Token::FunctionKeyword] => Some(true),
                        [VB6Token::StaticKeyword, VB6Token::SubKeyword] => Some(false),
                        // Anything else is a declaration
                        _ => None,
                    };
                    
                    match is_function_or_sub {
                        Some(true) => self.parse_function_statement(),  // Function
                        Some(false) => self.parse_sub_statement(),      // Sub
                        None => self.parse_declaration(),               // Declaration
                    }
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
        
        // Consume optional Public/Private/Friend keyword
        if self.at_token(VB6Token::PublicKeyword) 
            || self.at_token(VB6Token::PrivateKeyword) 
            || self.at_token(VB6Token::FriendKeyword) {
            self.consume_token();
            
            // Consume any whitespace after visibility modifier
            while self.at_token(VB6Token::Whitespace) {
                self.consume_token();
            }
        }
        
        // Consume optional Static keyword
        if self.at_token(VB6Token::StaticKeyword) {
            self.consume_token();
            
            // Consume any whitespace after Static
            while self.at_token(VB6Token::Whitespace) {
                self.consume_token();
            }
        }
        
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
        self.parse_code_block(|parser| {
            parser.at_keyword(VB6Token::EndKeyword) 
                && parser.peek_next_keyword() == Some(VB6Token::SubKeyword)
        });
        
        // Consume "End Sub" and trailing tokens
        if self.at_keyword(VB6Token::EndKeyword) {
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
        }
        
        self.builder.finish_node(); // SubStatement
    }
    
    /// Parse a Visual Basic 6 function with syntax:
    /// 
    /// \[ Public | Private | Friend \] \[ Static \] Function name \[ ( arglist ) \] \[ As type \]
    /// \[ statements \]
    /// \[ name = expression \]
    /// \[ Exit Function \]
    /// \[ statements \]
    /// \[ name = expression \]
    /// End Function
    /// 
    /// The Function statement syntax has these parts:
    /// 
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public   	  | Optional | Indicates that the Function procedure is accessible to all other procedures in all modules. If used in a module that contains an Option Private, the procedure is not available outside the project. |
    /// | Private  	  | Optional | Indicates that the Function procedure is accessible only to other procedures in the module where it is declared. |
    /// | Friend 	  | Optional | Used only in a class module. Indicates that the Function procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
    /// | Static 	  | Optional | Indicates that the Function procedure's local variables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the Function, even if they are used in the procedure. |
    /// | name 	      | Required | Name of the Function; follows standard variable naming conventions. |
    /// | arglist 	  | Optional | List of variables representing arguments that are passed to the Function procedure when it is called. Multiple variables are separated by commas. |
    /// | type 	      | Optional | Data type of the value returned by the Function procedure; may be Byte, Boolean, Integer, Long, Currency, Single, Double, Decimal (not currently supported), Date, String (except fixed length), Object, Variant, or any user-defined type. |
    /// | statements  | Optional | Any group of statements to be executed within the Function procedure.
    /// | expression  | Optional | Return value of the Function. |
    /// 
    /// The arglist argument has the following syntax and parts:
    /// 
    /// \[ Optional \] \[ ByVal | ByRef \] \[ ParamArray \] varname \[ ( ) \] \[ As type \] \[ = defaultvalue \]
    /// 
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)
    fn parse_function_statement(&mut self) {
        self.builder.start_node(SyntaxKind::FunctionStatement.to_raw());
        
        // Consume optional Public/Private/Friend keyword
        if self.at_token(VB6Token::PublicKeyword) 
            || self.at_token(VB6Token::PrivateKeyword) 
            || self.at_token(VB6Token::FriendKeyword) {
            self.consume_token();
            
            // Consume any whitespace after visibility modifier
            while self.at_token(VB6Token::Whitespace) {
                self.consume_token();
            }
        }
        
        // Consume optional Static keyword
        if self.at_token(VB6Token::StaticKeyword) {
            self.consume_token();
            
            // Consume any whitespace after Static
            while self.at_token(VB6Token::Whitespace) {
                self.consume_token();
            }
        }
        
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
        self.parse_code_block(|parser| {
            parser.at_keyword(VB6Token::EndKeyword) 
                && parser.peek_next_keyword() == Some(VB6Token::FunctionKeyword)
        });
        
        // Consume "End Function" and trailing tokens
        if self.at_keyword(VB6Token::EndKeyword) {
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
    
    /// Parse an If statement: If condition Then ... End If
    /// Handles both single-line and multi-line If statements
    /// 
    /// IfStatement
    /// ├─ If keyword
    /// ├─ condition tokens
    /// ├─ Then keyword
    /// ├─ body tokens
    /// ├─ ElseIfClause (if present)
    /// │  ├─ ElseIf keyword
    /// │  ├─ condition tokens
    /// │  ├─ Then keyword
    /// │  └─ body tokens
    /// ├─ ElseClause (if present)
    /// │  ├─ Else keyword
    /// │  └─ body tokens
    /// ├─ End keyword
    /// └─ If keyword
    /// 
    fn parse_if_statement(&mut self) {
        self.builder.start_node(SyntaxKind::IfStatement.to_raw());
        
        // Consume "If" keyword
        self.consume_token();
        
        // Parse the conditional expression
        self.parse_conditional();
        
        // Consume "Then" if present
        if self.at_token(VB6Token::ThenKeyword) {
            self.consume_token();
        }
        
        // Consume any whitespace after Then
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
        }
        
        // Check if this is a single-line If statement (has code on the same line after Then)
        let is_single_line = !self.at_token(VB6Token::Newline) && !self.is_at_end();
        
        if is_single_line {
            // Single-line If: consume everything until newline
            while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
            
            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        } else {
            // Multi-line If: consume newline after Then
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
            
            // Parse body until "End If", "Else", or "ElseIf"
            self.parse_code_block(|parser| {
                (parser.at_keyword(VB6Token::EndKeyword) && parser.peek_next_keyword() == Some(VB6Token::IfKeyword))
                    || parser.at_keyword(VB6Token::ElseIfKeyword)
                    || parser.at_keyword(VB6Token::ElseKeyword)
            });
            
            // Handle ElseIf and Else clauses
            while !self.is_at_end() {
                if self.at_keyword(VB6Token::ElseIfKeyword) {
                    // Parse ElseIf clause
                    self.parse_elseif_clause();
                } else if self.at_keyword(VB6Token::ElseKeyword) {
                    // Parse Else clause
                    self.parse_else_clause();
                } else {
                    break;
                }
            }
            
            // Consume "End If" and trailing tokens
            if self.at_keyword(VB6Token::EndKeyword) {
                // Consume "End"
                self.consume_token();
                
                // Consume any whitespace between "End" and "If"
                while self.at_token(VB6Token::Whitespace) {
                    self.consume_token();
                }
                
                // Consume "If"
                self.consume_token();
                
                // Consume until newline (including it)
                while !self.is_at_end() && !self.at_token(VB6Token::Newline) {
                    self.consume_token();
                }
                if self.at_token(VB6Token::Newline) {
                    self.consume_token();
                }
            }
        }
        
        self.builder.finish_node(); // IfStatement
    }
    
    /// Parse an ElseIf clause: ElseIf condition Then ... 
    fn parse_elseif_clause(&mut self) {
        self.builder.start_node(SyntaxKind::ElseIfClause.to_raw());
        
        // Consume "ElseIf" keyword
        self.consume_token();
        
        // Parse the conditional expression
        self.parse_conditional();
        
        // Consume "Then" if present
        if self.at_token(VB6Token::ThenKeyword) {
            self.consume_token();
        }
        
        // Consume any whitespace after Then
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
        }
        
        // Consume the newline after Then
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Parse body until "End If", "Else", or another "ElseIf"
        self.parse_code_block(|parser| {
            parser.at_keyword(VB6Token::ElseIfKeyword) 
                || parser.at_keyword(VB6Token::ElseKeyword) 
                || (parser.at_keyword(VB6Token::EndKeyword) && parser.peek_next_keyword() == Some(VB6Token::IfKeyword))
        });
        
        self.builder.finish_node(); // ElseIfClause
    }
    
    /// Parse an Else clause: Else ...
    fn parse_else_clause(&mut self) {
        self.builder.start_node(SyntaxKind::ElseClause.to_raw());
        
        // Consume "Else" keyword
        self.consume_token();
        
        // Consume any whitespace after Else
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
        }
        
        // Consume the newline after Else
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }
        
        // Parse body until "End If"
        self.parse_code_block(|parser| {
            parser.at_keyword(VB6Token::EndKeyword) && parser.peek_next_keyword() == Some(VB6Token::IfKeyword)
        });
        
        self.builder.finish_node(); // ElseClause
    }
    
    /// Parse a conditional expression.
    /// 
    /// This handles both:
    /// - Binary conditionals: `a = b`, `x > 5`, `name <> ""`
    /// - Unary conditionals: `Not condition`, `Not IsEmpty(x)`
    /// 
    /// The conditional is parsed until "Then" or newline is encountered.
    fn parse_conditional(&mut self) {
        // Skip any leading whitespace
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
        }
        
        // Check if this is a unary conditional starting with "Not"
        if self.at_keyword(VB6Token::NotKeyword) {
            self.builder.start_node(SyntaxKind::UnaryConditional.to_raw());
            
            // Consume "Not" keyword
            self.consume_token();
            
            // Consume any whitespace after "Not"
            while self.at_token(VB6Token::Whitespace) {
                self.consume_token();
            }
            
            // Consume the rest of the conditional expression until "Then" or newline
            while !self.is_at_end() 
                && !self.at_token(VB6Token::ThenKeyword) 
                && !self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
            
            self.builder.finish_node(); // UnaryConditional
        } else {
            // Binary conditional - parse left side, operator, right side
            self.builder.start_node(SyntaxKind::BinaryConditional.to_raw());
            
            // Consume tokens until we hit a comparison operator
            while !self.is_at_end() 
                && !self.at_token(VB6Token::ThenKeyword) 
                && !self.at_token(VB6Token::Newline) {
                
                // Check if we've hit a comparison operator
                if self.is_comparison_operator() {
                    // Consume the operator
                    self.consume_token();
                    
                    // Consume any whitespace after the operator
                    while self.at_token(VB6Token::Whitespace) {
                        self.consume_token();
                    }
                    
                    // Now consume the right side until "Then" or newline
                    while !self.is_at_end() 
                        && !self.at_token(VB6Token::ThenKeyword) 
                        && !self.at_token(VB6Token::Newline) {
                        self.consume_token();
                    }
                    break;
                }
                
                self.consume_token();
            }
            
            // If we didn't find an operator, we still consumed everything until "Then"
            // This handles cases like function calls that return boolean values
            
            self.builder.finish_node(); // BinaryConditional
        }
    }
    
    /// Check if the current token is a comparison operator
    fn is_comparison_operator(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::EqualityOperator)
                | Some(VB6Token::LessThanOperator)
                | Some(VB6Token::GreaterThanOperator)
        )
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
        F: Fn(&Parser) -> bool 
    {
        while !self.is_at_end() {
            if stop_conditions(self) {
                break;
            }
            
            // Dispatch to appropriate parsing methods based on current token
            match self.current_token() {
                // If statement: If condition Then ... End If
                Some(VB6Token::IfKeyword) => {
                    self.parse_if_statement();
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
                    self.consume_token_as_unknown();
                }
            }
        }
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
        self.peek_next_count_keywords(1).next()
    }
    
    /// Peek ahead and get the next `count` non-whitespace keywords from the current position.
    /// 
    /// # Arguments
    /// * `count` - Number of keywords to peek ahead (must be non-zero)
    /// 
    /// # Returns
    /// An iterator over the next `count` keywords (non-whitespace tokens)
    /// 
    /// # Panics
    /// Panics if `count` is zero
    fn peek_next_count_keywords(&self, count: usize) -> impl Iterator<Item = VB6Token> + '_ {
        assert!(count > 0, "count must be non-zero");
        
        self.tokens
            .iter()
            .skip(self.pos + 1)
            .filter(|(_, token)| *token != VB6Token::Whitespace)
            .take(count)
            .map(|(_, token)| *token)
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
