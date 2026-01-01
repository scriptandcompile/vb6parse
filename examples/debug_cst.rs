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

    let source_file = SourceFile::decode_with_replacement("test.bas", source).unwrap();
    let mut stream = source_file.source_stream();
    let tokens = tokenize(&mut stream).unwrap();
    let cst = parse(tokens);

    println!("Full CST structure:");
    println!("{}", cst.debug_tree());
}
