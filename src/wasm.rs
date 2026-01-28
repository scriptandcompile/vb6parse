//! WebAssembly bindings for the VB6 parser.
//!
//! This module provides functions for parsing and tokenizing VB6 code in a WebAssembly environment.
//! It exposes these functions to JavaScript via WebAssembly, allowing VB6 code analysis in the browser.
//!
//! Predominantly, this is designed for the needs of the `VB6Parser` playground.
//!

use crate::*;
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

/// Information about the output of the VB6 playground, including tokens, CST, and errors.
#[derive(Serialize, Deserialize)]
pub struct PlaygroundOutput {
    /// A list of tokens found in the source code, if tokenization was successful.
    pub tokens: Option<Vec<TokenInfo>>,
    /// The concrete syntax tree (CST) of the source code, if parsing was successful.
    pub cst: Option<String>,
    /// A list of errors encountered during parsing or tokenization.
    pub errors: Vec<ErrorInfo>,
    /// The time taken to parse the source code, in milliseconds.
    pub parse_time_ms: f64,
}

/// Parses VB6 code and returns a `PlaygroundOutput` object containing tokens, CST, and errors.
#[wasm_bindgen]
pub fn parse_vb6_code(
    _code: &str,
    _file_type: &str, // "project", "class", "module", "form"
) -> Result<JsValue, JsValue> {
    // Implementation that calls appropriate parser
    // Returns serialized PlaygroundOutput

    Ok(JsValue::FALSE)
}

/// Tokenizes VB6 code and returns a list of `TokenInfo` objects for quick preview.
#[wasm_bindgen]
pub fn tokenize_vb6_code(code: &str) -> Result<JsValue, JsValue> {
    // Returns just tokens for quick preview

    let mut tokens = vec![];

    let mut source_stream = crate::SourceStream::new("test.bas", code);

    let (token_stream_opt, _failures) = tokenize(&mut source_stream).unpack();

    let token_stream = token_stream_opt.unwrap();

    let mut column = 1;
    let mut line = 1;

    for (text, token) in token_stream {
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

    Ok(to_value(&tokens).unwrap())
}
