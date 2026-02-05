//! WebAssembly bindings for the VB6 parser.
//!
//! This module provides functions for parsing and tokenizing VB6 code in a WebAssembly environment.
//! It exposes these functions to JavaScript via WebAssembly, allowing VB6 code analysis in the browser.
//!
//! Predominantly, this is designed for the needs of the `VB6Parser` playground.
//!

use crate::{parsers, tokenize, Token, TokenStream};
use serde::{Deserialize, Serialize};
use serde_wasm_bindgen::to_value;
use wasm_bindgen::prelude::*;

/// Initializes the panic hook for better error messages in the browser console.
#[wasm_bindgen]
pub fn init_panic_hook() {
    console_error_panic_hook::set_once();
}

const IDENTIFIER: &str = "identifier";
const WHITESPACE: &str = "whitespace";
const KEYWORD: &str = "keyword";
const LITERAL: &str = "literal";
const OPERATOR: &str = "operator";
const COMMENT: &str = "comment";

/// Information about a single token in the source code.
#[derive(Serialize, Deserialize)]
pub struct TokenInfo {
    /// The kind of token (e.g., identifier, keyword, literal).
    pub kind: String,
    /// The actual text content of the token.
    pub content: String,
    /// The line number where the token appears.
    /// Lines are 1-indexed.
    pub line: u32,
    /// The column number where the token starts.
    /// Columns are 1-indexed.
    pub column: u32,
    /// The length of the token in characters.
    pub length: u32,
}

/// Information about a single error in the source code.
#[derive(Serialize, Deserialize)]
pub struct ErrorInfo {
    /// The line number where the error occurred.
    /// Lines are 1-indexed.
    pub line: usize,
    /// The column number where the error occurred.
    /// Columns are 1-indexed.
    pub column: usize,
    /// A descriptive error message.
    pub message: String,
}

/// Stats for the parse results
#[derive(Serialize, Deserialize)]
pub struct ParseStats {
    /// The total number of tokens in the source code.
    pub token_count: u32,
    /// The total number of nodes in the CST.
    pub node_count: u32,
    /// The maximum depth of the CST.
    pub tree_depth: u32,
}

/// Information about the output of the VB6 playground, including tokens, CST, and errors.
#[derive(Serialize, Deserialize)]
pub struct PlaygroundOutput {
    /// A list of tokens found in the source code, if tokenization was successful.
    pub tokens: Option<Vec<TokenInfo>>,
    /// The concrete syntax tree (CST) of the source code, if parsing was successful.
    pub cst: Option<CstNode>,
    /// A list of errors encountered during parsing or tokenization.
    pub errors: Vec<ErrorInfo>,
    /// The time taken to parse the source code, in milliseconds.
    pub parse_time_ms: f64,
    /// Statistics about the parse results.
    pub stats: ParseStats,
}

/// A node in the concrete syntax tree (CST).
#[derive(Serialize, Deserialize)]
pub struct CstNode {
    /// The kind of CST node (e.g., `FunctionDeclaration`, `IfStatement`).
    pub kind: String,
    /// The range of the node in the source code as [start, end] positions.
    pub range: [u32; 2],
    /// The child nodes of this CST node.
    pub children: Vec<CstNode>,
}

/// Convert a `parsers::cst::CstNode` to the wasm-facing `CstNode` recursively.
fn convert_cst_node(node: &parsers::cst::CstNode) -> CstNode {
    convert_cst_node_with_offset(node, 0).0
}

/// Convert a `parsers::cst::CstNode` to the wasm-facing `CstNode` recursively,
/// tracking the current byte offset. Returns the converted node and the next offset.
fn convert_cst_node_with_offset(node: &parsers::cst::CstNode, start_offset: u32) -> (CstNode, u32) {
    let text_len = u32::try_from(node.text().len()).unwrap_or(0);
    let end_offset = start_offset + text_len;

    let mut current_offset = start_offset;
    let children: Vec<CstNode> = node
        .children()
        .iter()
        .map(|child| {
            let (child_node, next_offset) = convert_cst_node_with_offset(child, current_offset);
            current_offset = next_offset;
            child_node
        })
        .collect();

    let wasm_node = CstNode {
        kind: format!("{:?}", node.kind()),
        range: [start_offset, end_offset],
        children,
    };

    (wasm_node, end_offset)
}

/// Helper to count nodes recursively
fn count_nodes(node: &CstNode) -> u32 {
    1 + node.children.iter().map(count_nodes).sum::<u32>()
}

/// Helper to compute tree depth recursively
fn tree_depth(node: &CstNode) -> u32 {
    if node.children.is_empty() {
        1
    } else {
        1 + node.children.iter().map(tree_depth).max().unwrap_or(0)
    }
}

/// Parses VB6 code and returns a `PlaygroundOutput` object containing tokens, CST, and errors.
///
/// # Errors
///
/// So far we do not correctly handle errors and failures and just panic but this must eventually
/// be converted into an error value.
///
/// # Panics
///
/// Currently, we are doing minimal error recovery and checking for the playground as this
/// is an attempt to get the system up and working well enough to demonstrate the possibilities.
/// As is, we can produce a panic if the input can not be tokenized.
///
#[wasm_bindgen]
pub fn parse_vb6_code(
    code: &str,
    _file_type: &str, // "project", "class", "module", "form"
) -> Result<JsValue, JsError> {
    // Implementation that calls appropriate parser
    // Returns serialized PlaygroundOutput

    let mut source_stream = crate::SourceStream::new("test.bas", code);

    let (token_stream_opt, _failures) = tokenize(&mut source_stream).unpack();

    let token_stream = token_stream_opt.unwrap();

    let tokens = produce_tokens(token_stream.clone());

    // Parse CST using the token stream and convert to CstNode
    let cst = parsers::cst::parse(token_stream.clone());
    let cst_node = convert_cst_node(&cst.to_root_node());

    let token_count = u32::try_from(tokens.len())?;

    let parse_stats = ParseStats {
        token_count,
        node_count: count_nodes(&cst_node),
        tree_depth: tree_depth(&cst_node),
    };

    let playground_output = PlaygroundOutput {
        tokens: Some(tokens),
        cst: Some(cst_node),
        errors: vec![],
        parse_time_ms: 0.0f64,
        stats: parse_stats,
    };

    Ok(to_value(&playground_output).unwrap())
}

/// Produces a list of `TokenInfo` objects from a `TokenStream`.
#[must_use]
pub fn produce_tokens(token_stream: TokenStream) -> Vec<TokenInfo> {
    let mut tokens = vec![];

    let mut column = 1;
    let mut line = 1;

    for (text, token) in token_stream.into_tokens() {
        let kind = match token {
            Token::Whitespace => WHITESPACE.to_string(),
            Token::Identifier => IDENTIFIER.to_string(),
            Token::DateTimeLiteral
            | Token::DecimalLiteral
            | Token::SingleLiteral
            | Token::DoubleLiteral
            | Token::StringLiteral
            | Token::IntegerLiteral
            | Token::LongLiteral => LITERAL.to_string(),
            Token::EndOfLineComment | Token::RemComment => COMMENT.to_string(),
            _ => {
                if token.is_keyword() {
                    KEYWORD.to_string()
                } else if token.is_operator() {
                    OPERATOR.to_string()
                } else {
                    format!("{token:?}")
                }
            }
        };

        if token == Token::Newline {
            line += 1;
            column = 0;
        }

        let content = text.to_string();
        // Calculate the length of the token in characters, safely converting to u32
        // If the character count exceeds u32::MAX, default to 0
        // VB6 is limited by 32-bit integer sizes for column and length of source code
        // so this conversion should be unnecessary in practice.
        let length = u32::try_from(content.chars().count()).unwrap_or(0);

        let token = TokenInfo {
            kind,
            content,
            line,
            column,
            length,
        };

        // We update the column position here because the column is supposed to be
        // the *starting* position of the token and the length represents the number
        // of characters in the token which should be added to the column for the next token.
        // This ensures that the column number correctly reflects the position of the next token
        // on the next run of the loop.
        column += length;

        tokens.push(token);
    }

    tokens
}

/// Tokenizes VB6 code and returns a list of `TokenInfo` objects for quick preview.
///
/// # Errors
///
/// So far we do not correctly handle errors and failures and just panic but this must eventually
/// be converted into an error value.
///
/// # Panics
///
/// Currently, we are doing minimal error recovery and checking for the playground as this
/// is an attempt to get the system up and working well enough to demonstrate the possibilities.
/// As is, we can produce a panic if the input can not be tokenized.
///
#[wasm_bindgen]
pub fn tokenize_vb6_code(code: &str) -> Result<JsValue, JsError> {
    // Returns just tokens for quick preview

    let mut source_stream = crate::SourceStream::new("test.bas", code);

    let (token_stream_opt, _failures) = tokenize(&mut source_stream).unpack();

    let token_stream = token_stream_opt.unwrap();

    let tokens = produce_tokens(token_stream);

    Ok(to_value(&tokens).unwrap())
}
