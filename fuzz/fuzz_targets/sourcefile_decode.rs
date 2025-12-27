#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::SourceFile;

fuzz_target!(|data: &[u8]| {
    // Test with arbitrary filenames and byte sequences
    // This tests Windows-1252 decoding robustness with arbitrary byte sequences
    let _ = SourceFile::decode_with_replacement("fuzz.vbp", data);
    let _ = SourceFile::decode_with_replacement("fuzz.bas", data);
    let _ = SourceFile::decode_with_replacement("fuzz.cls", data);
    let _ = SourceFile::decode_with_replacement("fuzz.frm", data);
});
