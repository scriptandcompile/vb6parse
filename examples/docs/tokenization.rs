use vb6parse::{tokenize, SourceStream};

fn main() {
    let code = "Dim x As Integer ' Declare a variable";
    let mut source = SourceStream::new("test.bas", code);
    let result = tokenize(&mut source);
    let (token_stream, _) = result.unpack();

    if let Some(tokens) = token_stream {
        println!("Tokens found: {}", tokens.len());

        for (text, token) in tokens.tokens() {
            println!("  {token:?}: '{text}'");
        }
    }
}
