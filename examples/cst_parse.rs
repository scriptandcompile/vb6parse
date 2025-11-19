//! Example demonstrating how to use the CST parser with a TokenStream
//!
//! This example shows how to create a TokenStream and parse it into a
//! Concrete Syntax Tree (CST) that represents VB6 code structure.

use vb6parse::language::VB6Token;
use vb6parse::parsers::cst::parse;
use vb6parse::tokenstream::TokenStream;

fn main() {
    // Create a token stream representing a simple VB6 subroutine
    let tokens = vec![
        ("Sub", VB6Token::SubKeyword),
        (" ", VB6Token::Whitespace),
        ("HelloWorld", VB6Token::Identifier),
        ("(", VB6Token::LeftParenthesis),
        (")", VB6Token::RightParenthesis),
        ("\n", VB6Token::Newline),
        ("    ", VB6Token::Whitespace),
        ("' This is a comment\n", VB6Token::EndOfLineComment),
        ("    ", VB6Token::Whitespace),
        ("Dim", VB6Token::DimKeyword),
        (" ", VB6Token::Whitespace),
        ("x", VB6Token::Identifier),
        (" ", VB6Token::Whitespace),
        ("As", VB6Token::AsKeyword),
        (" ", VB6Token::Whitespace),
        ("Integer", VB6Token::IntegerKeyword),
        ("\n", VB6Token::Newline),
        ("End", VB6Token::EndKeyword),
        (" ", VB6Token::Whitespace),
        ("Sub", VB6Token::SubKeyword),
        ("\n", VB6Token::Newline),
    ];

    // Create a TokenStream
    let token_stream = TokenStream::new("example.bas".to_string(), tokens);

    // Parse the TokenStream into a CST
    let cst = parse(token_stream);

    // Display information about the CST
    println!("CST Root Kind: {:?}", cst.root_kind());
    println!("Number of children: {}", cst.child_count());
    println!("\nFull text of the CST:");
    println!("{}", cst.text());
    println!("\nDebug tree structure:");
    println!("{}", cst.debug_tree());
}
