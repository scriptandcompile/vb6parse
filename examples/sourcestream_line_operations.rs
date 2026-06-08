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

    println!("=== Line Operations ===");

    // Take until newline to get the first line
    if let Some((line, newline)) = stream.take_until_newline() {
        println!("First line: '{line}'");
        if let Some(nl) = newline {
            println!("Newline character(s): '{nl:?}'");
        }
    }

    // Get line information
    let start_of_line = stream.start_of_line();
    let end_of_line = stream.end_of_line();
    println!("Start of current line: {start_of_line}");
    println!("End of current line: {end_of_line}");
}
