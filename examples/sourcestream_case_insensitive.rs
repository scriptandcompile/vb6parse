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

    println!("=== Case-Insensitive Operations ===");

    // Search for keywords case-insensitively
    let keywords_to_find = ["private", "PUBLIC", "Function", "MSGBOX"];

    for keyword in &keywords_to_find {
        let mut found_positions = Vec::new();

        while !stream.is_empty() {
            if stream
                .peek_text(*keyword, Comparator::CaseInsensitive)
                .is_some()
            {
                found_positions.push(stream.offset());
                stream.forward(keyword.len());
            } else {
                stream.forward(1);
            }

            // Limit search results
            if found_positions.len() >= 3 {
                break;
            }
        }

        if !found_positions.is_empty() {
            println!("Found '{keyword}' at positions: {found_positions:?}");
        }
    }
}
