//! Example demonstrating various functionalities of the `SourceStream` API.
//! This example shows how to navigate, peek, and extract text from a VB6 source code stream.
//! It includes examples of case-sensitive and case-insensitive operations,
//! as well as line-based operations and error handling.
//!

use vb6parse::io::SourceStream;

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

    println!("=== Bounds and Error Handling ===");

    // Try to peek beyond the end of the stream
    let beyond_end = stream.peek(stream.contents.len() + 10);
    println!("Peeking beyond end result: {beyond_end:?}");

    // Move to near the end
    let original_len = stream.contents.len();
    stream.forward(original_len - 10);

    println!("Moved to near end, offset: {}", stream.offset());
    println!("Is empty: {}", stream.is_empty());

    // Try to take more characters than available
    if let Some(remaining) = stream.peek(20) {
        println!("Last characters: '{remaining}'");
    } else {
        println!("Cannot peek 20 characters from current position");
    }

    // Take the remaining characters
    let remaining_count = stream.contents.len() - stream.offset();
    if let Some(final_chars) = stream.take_count(remaining_count) {
        println!("Final {remaining_count} characters: '{final_chars:?}'");
    }

    println!("Stream is empty: {}", stream.is_empty());
    println!("Final offset: {}", stream.offset());
}
