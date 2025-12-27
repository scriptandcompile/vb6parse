#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::SourceFile;

fuzz_target!(|data: &[u8]| {
    if let Ok(source_file) = SourceFile::decode_with_replacement("fuzz.bas", data) {
        let mut stream = source_file.source_stream();
        let result = vb6parse::lexer::tokenize(&mut stream);

        // Get the TokenStream from the ParseResult
        let (token_stream_opt, _failures) = result.unpack();

        // If we got a token stream, iterate through tokens
        if let Some(mut token_stream) = token_stream_opt {
            for (_text, _token_kind) in &mut token_stream {
                // Just iterate - accessing the data is enough to test for panics
            }
        }
    }
});
