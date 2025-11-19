use vb6parse::{parse, tokenize::tokenize, sourcestream::SourceStream};

#[test]
fn function_with_do_loop_ending_at_end_function() {
    let code = r#"Function Test()
Do
Loop
End Function
"#;

    let mut stream = SourceStream::new("test.bas", code);
    let token_result = tokenize(&mut stream);
    assert!(!token_result.has_failures(), "Tokenization should succeed");
    
    let tokens = token_result.result.unwrap();
    let cst = parse(tokens);
    
    println!("{}", cst.debug_tree());
    
    // Check if "End" appears as Unknown in the tree
    let tree_str = cst.debug_tree();
    assert!(!tree_str.contains("Unknown"), "Should not have any Unknown tokens\n{}", tree_str);
}

#[test]
fn function_with_do_until_loop() {
    let code = r#"Function Test()
Do Until x = ""
  y = z
Loop
End Function
"#;

    let mut stream = SourceStream::new("test.bas", code);
    let token_result = tokenize(&mut stream);
    assert!(!token_result.has_failures(), "Tokenization should succeed");
    
    let tokens = token_result.result.unwrap();
    let cst = parse(tokens);
    
    println!("{}", cst.debug_tree());
    
    // Check if "End" appears as Unknown in the tree
    let tree_str = cst.debug_tree();
    assert!(!tree_str.contains("Unknown"), "Should not have any Unknown tokens\n{}", tree_str);
}
