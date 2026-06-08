//! Example demonstrating various functionalities of the `SourceStream` API.
//! This example shows how to navigate, peek, and extract text from a VB6 source code stream.
//! It includes examples of case-sensitive and case-insensitive operations,
//! as well as line-based operations and error handling.
//!

use vb6parse::io::{Comparator, SourceStream};

// Sample VB6 code content
const VB6_CODE: &str = r#"Private Sub Form_Load()
    Dim x As Integer
    Dim message As String
    x = 42
    message = "Hello, World!"
    MsgBox message
End Sub

' This is a comment
Public Function Calculate(a As Integer, b As Integer) As Integer
    Calculate = a + b
End Function"#;

fn main() {
    // Create a SourceStream
    let mut stream = SourceStream::new("example.bas", VB6_CODE);

    println!("=== SourceStream Example ===");
    println!("File: {}", stream.file_name());
    println!("================================");
    println!();
    println!("{VB6_CODE}");
    println!();
    println!("================================");
    println!();
    println!("Content length: {} characters", stream.contents.len());
    println!("Initial offset: {}\n", stream.offset());
    println!();

    println!("=== Parsing Tokens Example ===");

    let mut tokens = Vec::new();

    while !stream.is_empty() {
        // Skip whitespace
        if stream.take_ascii_whitespaces().is_some() {
            continue;
        }

        // Try to take an identifier or keyword
        if let Some(identifier) = stream.take_ascii_underscore_alphanumerics() {
            tokens.push(("Identifier", identifier));
        }
        // Try to take punctuation
        else if let Some(punctuation) = stream.take_ascii_punctuation() {
            tokens.push(("Punctuation", punctuation));
        }
        // Try to take a string literal (simplified)
        else if stream.peek_text("\"", Comparator::CaseSensitive).is_some() {
            stream.forward(1); // Skip opening quote
            if let Some((string_content, _)) = stream.take_until("\"", Comparator::CaseSensitive) {
                stream.forward(1); // Skip closing quote
                tokens.push(("String", string_content));
            }
        }
        // Try to take digits
        else if let Some(digits) = stream.take_ascii_digits() {
            tokens.push(("Number", digits));
        }
        // Skip newlines
        else if stream.take_newline().is_some() {
            tokens.push(("Newline", "\\n"));
        }
        // Unknown character - skip it
        else {
            stream.forward(1);
        }

        // Limit output for demonstration
        if tokens.len() >= 20 {
            break;
        }
    }

    println!("First 20 tokens:");
    for (i, (token_type, value)) in tokens.iter().enumerate() {
        println!("  {:2}: {:12} = '{}'", i + 1, token_type, value);
    }
}
