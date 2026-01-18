use crate::*;
use serde::{Deserialize, Serialize};
use serde_wasm_bindgen::to_value;
use wasm_bindgen::prelude::*;

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

#[derive(Serialize, Deserialize)]
pub struct TokenInfo {
    pub kind: String,
    pub content: String,
    pub line: u32,
    pub column: u32,
    pub length: u32,
}

#[derive(Serialize, Deserialize)]
pub struct ErrorInfo {
    pub line: usize,
    pub column: usize,
    pub message: String,
}

#[derive(Serialize, Deserialize)]
pub struct PlaygroundOutput {
    pub tokens: Option<Vec<TokenInfo>>,
    pub cst: Option<String>,
    pub errors: Vec<ErrorInfo>,
    pub parse_time_ms: f64,
}

#[wasm_bindgen]
pub fn parse_vb6_code(
    code: &str,
    file_type: &str, // "project", "class", "module", "form"
) -> Result<JsValue, JsValue> {
    // Implementation that calls appropriate parser
    // Returns serialized PlaygroundOutput

    Ok(JsValue::FALSE)
}

#[wasm_bindgen]
pub fn tokenize_vb6_code(code: &str) -> Result<JsValue, JsValue> {
    // Returns just tokens for quick preview

    let mut tokens = vec![];

    let mut source_stream = crate::SourceStream::new("test.bas", code);

    let (token_stream_opt, _failures) = tokenize(&mut source_stream).unpack();

    let token_stream = token_stream_opt.unwrap();

    for (text, token) in token_stream {
        let kind = match token {
            Token::Whitespace => WHITESPACE.to_string(),
            Token::Identifier => IDENTIFIER.to_string(),
            Token::DateLiteral
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
                    "".to_string()
                }
            }
        };

        let token = TokenInfo {
            kind: kind,
            content: text.to_string(),
            line: 1,
            column: 1,
            length: 1,
        };

        tokens.push(token);
    }

    Ok(to_value(&tokens).unwrap())
}
