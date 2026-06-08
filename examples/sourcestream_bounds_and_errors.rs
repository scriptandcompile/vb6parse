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

    println!("=== Basic Navigation Example ===");

    // Peek at the first 7 characters
    if let Some(peek_text) = stream.peek(7) {
        println!("Peeking at first 7 characters: '{peek_text}'");
    }

    // Check if we can find "Private" at the current position (case-sensitive)
    if let Some(private_text) = stream.peek_text("Private", Comparator::CaseSensitive) {
        println!("Found 'Private' at current position: '{private_text}'");
        stream.forward(7); // Move past "Private"
    }

    println!("Offset after moving past 'Private': {}", stream.offset());

    // Take a space
    if let Some(whitespace) = stream.take_ascii_whitespaces() {
        println!("Consumed whitespace: '{whitespace:?}'");
    }

    // Take the word "Sub"
    if let Some(sub_keyword) = stream.take_ascii_alphabetics() {
        println!("Found keyword: '{sub_keyword}'");
    }

    println!();
}
