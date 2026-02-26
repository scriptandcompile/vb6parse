//! A simple example to demonstrate how to debug the Concrete Syntax Tree (CST)
//! using a VB6 source file.
//!
//! This example reads a VB6 source file, tokenizes its content, parses it into a CST,
//! and then prints the full structure of the CST for debugging purposes.
//!

use vb6parse::io::SourceFile;
use vb6parse::lexer::tokenize;
use vb6parse::parsers::cst::parse;

fn main() {
    let source = b"Function GetValues() As Variant
    GetValues = Array(10, 20, 30)
End Function
";

    let source_file = SourceFile::decode_with_replacement("test.bas", source)
        .expect("Unable to decode the sourcefile with replacements.");
    let mut stream = source_file.source_stream();
    let (tokens_opt, errors, warnings) = tokenize(&mut stream).unpack_with_severity();

    if !errors.is_empty() {
        eprintln!("Errors during tokenization:");
        for error in errors {
            error.print();
        }
    }

    if !warnings.is_empty() {
        eprintln!("Warnings during tokenization:");
        for warning in warnings {
            warning.print();
        }
    }

    let tokens = tokens_opt.expect("Tokens should be present.");

    let cst = parse(tokens);

    println!("Full CST structure:");
    println!("{}", cst.debug_tree());
}
