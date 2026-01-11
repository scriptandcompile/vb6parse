#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::parsers::Comparator;
use vb6parse::SourceFile;

fuzz_target!(|data: &[u8]| {
    if let Ok(source_file) = SourceFile::decode_with_replacement("fuzz.bas", data) {
        let mut stream = source_file.source_stream();

        // Exercise various stream operations
        while !stream.is_empty() {
            let _ = stream.peek(1);
            let _ = stream.peek(5);

            // Test pattern matching
            let _ = stream.peek_text("test", Comparator::CaseInsensitive);
            let _ = stream.peek_text("END", Comparator::CaseSensitive);

            stream.forward(1);
        }
    }
});
