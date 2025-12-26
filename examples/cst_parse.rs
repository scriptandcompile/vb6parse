//! Example demonstrating how to use the CST parser with a `TokenStream`
//!
//! This example shows how to create a `TokenStream` and parse it into a
//! Concrete Syntax Tree (CST) that represents VB6 code structure.

use vb6parse::language::Token;
use vb6parse::lexer::TokenStream;
use vb6parse::parsers::cst::parse;

fn main() {
    // Create a token stream representing a simple VB6 subroutine
    let tokens = vec![
        ("Sub", Token::SubKeyword),
        (" ", Token::Whitespace),
        ("HelloWorld", Token::Identifier),
        ("(", Token::LeftParenthesis),
        (")", Token::RightParenthesis),
        ("\n", Token::Newline),
        ("    ", Token::Whitespace),
        ("' This is a comment\n", Token::EndOfLineComment),
        ("    ", Token::Whitespace),
        ("Dim", Token::DimKeyword),
        (" ", Token::Whitespace),
        ("x", Token::Identifier),
        (" ", Token::Whitespace),
        ("As", Token::AsKeyword),
        (" ", Token::Whitespace),
        ("Integer", Token::IntegerKeyword),
        ("\n", Token::Newline),
        ("End", Token::EndKeyword),
        (" ", Token::Whitespace),
        ("Sub", Token::SubKeyword),
        ("\n", Token::Newline),
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
