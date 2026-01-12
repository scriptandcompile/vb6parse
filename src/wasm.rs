use crate::language::Token;
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub fn init_panic_hook() {
    console_error_panic_hook::set_once();
}

#[derive(Serialize, Deserialize)]
pub struct TokenInfo {
    pub name: String,
    pub offset: u32,
    pub token: Token,
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

    Ok(JsValue::FALSE)
}
