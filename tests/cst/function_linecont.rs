use vb6parse::{parse, tokenize, SourceStream};

#[test]
fn function_with_line_continuation_in_params() {
    let source = r#"
Public Function Test( _
  ByVal x As Long _
) As String
    Test = "hello"
End Function
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    println!("{}", debug);

    // Should have a FunctionStatement node
    assert!(
        debug.contains("FunctionStatement"),
        "Should be FunctionStatement"
    );
    // The function itself should not be parsed as a DimStatement
    // (although it may contain DimStatement nodes inside for variable declarations)
    assert!(
        debug.contains("  FunctionStatement@"),
        "Function should be at root level, not inside DimStatement"
    );
}

#[test]
fn function_with_line_continuation_after_open_paren() {
    // This is the exact pattern from audiostation modArgs.bas argGetSwitchArg
    let source = r#"
Public Function argGetSwitchArg( _
  ByRef Switch As String, _
  Optional ByRef Position As Long, _
  Optional ByVal UseWildcard As Boolean _
) As String
Dim I&
argGetSwitchArg = ""
End Function
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    println!("{}", debug);

    // Should have a FunctionStatement node
    assert!(
        debug.contains("FunctionStatement"),
        "Should be FunctionStatement"
    );
    // The function itself should not be parsed as a DimStatement
    assert!(
        debug.contains("  FunctionStatement@"),
        "Function should be at root level, not inside DimStatement"
    );
}

#[test]
fn function_with_do_loop_before_end() {
    // Test that "End Function" after a DO loop is recognized correctly
    let source = r#"
Public Function Test(ByVal x As Long) As String
Dim i As Long
Do
    i = i + 1
Loop
End Function
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    println!("{}", debug);

    assert!(
        debug.contains("FunctionStatement"),
        "Should have FunctionStatement"
    );
    assert!(
        !debug.contains("Unknown"),
        "Should not have any Unknown tokens"
    );
    assert!(
        debug.contains("  FunctionStatement@"),
        "Function should be at root level"
    );
}

#[test]
fn function_with_line_continuation_in_if_condition() {
    // Test from audiostation modArgs.bas - line continuation in IF condition
    let source = r#"
Public Function argGetArgs(ByRef argv() As String, ByRef argc As Long, _
 Optional ByVal Args As String)
Dim strArgTemp As String
Do Until strArgTemp = ""
  If InStr(1, strArgTemp, Chr$(34)) <> 0 And _
     InStr(1, strArgTemp, Chr$(34)) < InStr(1, strArgTemp, " ") Then
    strArgTemp = ""
  End If
Loop
End Function
"#;

    let mut source_stream = SourceStream::new("test.bas", source);
    let result = tokenize(&mut source_stream);
    let token_stream = result.result.expect("Tokenization should succeed");
    let cst = parse(token_stream);

    let debug = cst.debug_tree();
    println!("{}", debug);

    assert!(
        debug.contains("FunctionStatement"),
        "Should have FunctionStatement"
    );
    assert!(
        !debug.contains("Unknown"),
        "Should not have any Unknown tokens"
    );
    assert!(
        debug.contains("  FunctionStatement@"),
        "Function should be at root level"
    );
}
