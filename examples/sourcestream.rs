use vb6parse::{
    sourcestream::{SourceStream, Comparator},
};

fn main() {
    // Sample VB6 code content
    let vb6_code = r#"Private Sub Form_Load()
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

    // Create a SourceStream
    let mut stream = SourceStream::new("example.bas", vb6_code);

    println!("=== SourceStream Example ===");
    println!("File: {}", stream.file_name());
    println!("Content length: {} characters", stream.contents.len());
    println!("Initial offset: {}\n", stream.offset());

    // Example 1: Basic navigation and peeking
    println!("=== Example 1: Basic Navigation ===");
    
    // Peek at the first 7 characters
    if let Some(peek_text) = stream.peek(7) {
        println!("Peeking at first 7 characters: '{}'", peek_text);
    }
    
    // Check if we can find "Private" at the current position (case-sensitive)
    if let Some(private_text) = stream.peek_text("Private", Comparator::CaseSensitive) {
        println!("Found 'Private' at current position: '{}'", private_text);
        stream.forward(7); // Move past "Private"
    }
    
    println!("Offset after moving past 'Private': {}", stream.offset());
    
    // Take a space
    if let Some(whitespace) = stream.take_ascii_whitespaces() {
        println!("Consumed whitespace: '{:?}'", whitespace);
    }
    
    // Take the word "Sub"
    if let Some(sub_keyword) = stream.take_ascii_alphabetics() {
        println!("Found keyword: '{}'", sub_keyword);
    }

    println!();

    // Example 2: Line-based operations
    println!("=== Example 2: Line Operations ===");
    
    // Reset to beginning
    stream.offset = 0;
    
    // Take until newline to get the first line
    if let Some((line, newline)) = stream.take_until_newline() {
        println!("First line: '{}'", line);
        if let Some(nl) = newline {
            println!("Newline character(s): '{:?}'", nl);
        }
    }
    
    // Get line information
    let start_of_line = stream.start_of_line();
    let end_of_line = stream.end_of_line();
    println!("Start of current line: {}", start_of_line);
    println!("End of current line: {}", end_of_line);
    
    // Example 3: Parsing identifiers and keywords
    println!("\n=== Example 3: Parsing Tokens ===");
    
    // Reset to the beginning and demonstrate parsing
    stream.offset = 0;
    
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

    // Example 4: Case-insensitive searching
    println!("\n=== Example 4: Case-Insensitive Operations ===");
    
    stream.offset = 0;
    
    // Search for keywords case-insensitively
    let keywords_to_find = ["private", "PUBLIC", "Function", "MSGBOX"];
    
    for keyword in &keywords_to_find {
        let mut search_stream = stream.clone();
        let mut found_positions = Vec::new();
        
        while !search_stream.is_empty() {
            if let Some(_) = search_stream.peek_text(*keyword, Comparator::CaseInsensitive) {
                found_positions.push(search_stream.offset());
                search_stream.forward(keyword.len());
            } else {
                search_stream.forward(1);
            }
            
            // Limit search results
            if found_positions.len() >= 3 {
                break;
            }
        }
        
        if !found_positions.is_empty() {
            println!("Found '{}' at positions: {:?}", keyword, found_positions);
        }
    }

    // Example 5: Taking text until specific patterns
    println!("\n=== Example 5: Pattern-based Extraction ===");
    
    stream.offset = 0;
    
    // Find and extract function definitions
    while !stream.is_empty() {
        // Look for "Function" keyword
        if let Some(_) = stream.peek_text("Function", Comparator::CaseInsensitive) {
            let function_start = stream.offset();
            
            // Take until the end of the line to get function signature
            if let Some((signature, _)) = stream.take_until_newline() {
                println!("Found function at offset {}: '{}'", function_start, signature.trim());
            }
        } else {
            stream.forward(1);
        }
    }

    // Example 6: Error handling and bounds checking
    println!("\n=== Example 6: Bounds and Error Handling ===");
    
    stream.offset = 0;
    
    // Try to peek beyond the end of the stream
    let beyond_end = stream.peek(stream.contents.len() + 10);
    println!("Peeking beyond end result: {:?}", beyond_end);
    
    // Move to near the end
    let original_len = stream.contents.len();
    stream.forward(original_len - 10);
    
    println!("Moved to near end, offset: {}", stream.offset());
    println!("Is empty: {}", stream.is_empty());
    
    // Try to take more characters than available
    if let Some(remaining) = stream.peek(20) {
        println!("Last characters: '{}'", remaining);
    } else {
        println!("Cannot peek 20 characters from current position");
    }
    
    // Take the remaining characters
    let remaining_count = stream.contents.len() - stream.offset();
    if let Some(final_chars) = stream.take_count(remaining_count) {
        println!("Final {} characters: '{:?}'", remaining_count, final_chars);
    }
    
    println!("Stream is empty: {}", stream.is_empty());
    println!("Final offset: {}", stream.offset());

    println!("\n=== SourceStream Example Complete ===");
}