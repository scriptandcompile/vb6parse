use vb6parse::{sourcestream::SourceStream, tokenize::tokenize};

fn main() {
    let code = "Dim x As Integer";
    let mut input = SourceStream::new("example.bas", code);

    // Parse the code into a TokenStream
    let result = tokenize(&mut input);

    // Show error handling first
    if result.has_failures() {
        println!("Parsing failures:");
        for failure in result.failures() {
            println!("  {failure:?}");
        }
    }

    let (token_stream_opt, _failures) = result.unpack();
    if let Some(token_stream) = token_stream_opt {
        println!("File name: {}", token_stream.file_name());
        println!("Total tokens: {}", token_stream.len());
        println!("Current offset: {}", token_stream.offset());
        println!();

        // Iterate through tokens
        println!("Tokens:");
        for (i, &(text, token_type)) in token_stream.tokens().iter().enumerate() {
            println!("  {i}: {token_type:?} = '{text}'");
        }

        // Demonstrate navigation
        println!("\nDemonstrating TokenStream navigation:");
        let mut stream = token_stream;

        // Get first few tokens using next()
        for i in 0..3 {
            if let Some((text, token)) = stream.next() {
                println!("Token {i}: {token:?} = '{text}'");
            }
        }

        println!("Current offset after reading 3 tokens: {}", stream.offset());

        // Reset and use indexing
        stream.reset();
        println!("\nAfter reset, offset: {}", stream.offset());
        println!("First token by index: {:?}", stream[0]);
        println!(
            "Current token by current(): {:?}",
            stream.current().unwrap()
        );
    }
}
