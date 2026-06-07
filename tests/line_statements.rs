//! Tests for VB6 Line graphics statement with coordinate pair syntax.
//!
//! The Line method in VB6 has a special syntax for drawing lines and shapes:
//! `object.Line (x1, y1)-(x2, y2) [, color] [, BF]`
//!
//! The key challenge is parsing the `-(` between coordinate pairs correctly,
//! as this involves a minus operator followed by a parenthesized expression.
//! Previously, the `-(` between coordinate pairs was parsed as Unknown tokens.

use vb6parse::*;

#[test]
fn line_statement_simple_coordinates() {
    let source = r"
Sub Test()
    Picture1.Line (0, 0)-(100, 100)
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    // Verify the CST structure contains a CallStatement for the Line method
    let serializable = cst.to_serializable();
    let text = format!("{serializable:#?}");

    if !text.contains("CallStatement") {
        println!("DEBUG: CST does not contain CallStatement");
        println!("{text}");
    }

    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(
        text.contains("Line"),
        "Should contain Line keyword/identifier"
    );
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}

#[test]
fn line_statement_with_color_argument() {
    let source = r"
Sub Test()
    Picture1.Line (10, 20)-(200, 150), RGB(255, 0, 0)
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(text.contains("RGB"), "Should contain RGB function call");
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}

#[test]
fn line_statement_with_variables() {
    let source = r"
Sub DrawGradient()
    Dim x As Long, iHeight As Long
    x = 50
    iHeight = 100
    DstObject.Line (x, 0)-(x, iHeight), RGB(255, 128, 0)
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(text.contains("Line"), "Should contain Line");
    assert!(
        text.contains("DstObject"),
        "Should contain DstObject identifier"
    );
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}

#[test]
fn line_statement_box_filled() {
    let source = r"
Sub Test()
    Picture1.Line (10, 10)-(200, 200), vbRed, BF
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(
        text.contains("BF"),
        "Should contain BF identifier for filled box"
    );
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}

#[test]
fn line_statement_box_filled_no_color() {
    let source = r"
Sub Test()
    Picture1.Line (10, 10)-(200, 200), , BF
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}

#[test]
fn line_statement_multiple_in_loop() {
    let source = r"
Sub Test()
    Dim i As Integer
    For i = 0 To 10
        Picture1.Line (i * 10, 0)-(i * 10, 100)
    Next i
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());
    assert!(text.contains("ForStatement"), "Should contain ForStatement");
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}

#[test]
fn line_statement_expressions_in_coordinates() {
    let source = r"
Sub Test()
    Dim R2 As Long, G2 As Long, B2 As Long
    Picture1.Line (x + 5, y * 2)-(width - x, height / 2), RGB(R2, G2, B2)
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(
        text.contains("BinaryExpression"),
        "Should contain BinaryExpression for coordinate calculations"
    );
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}

#[test]
fn line_statement_from_histograms_example() {
    // This is the actual pattern from the Histograms-basic VB6 code that was producing Unknown tokens
    let source = r"
Sub CreateGradient()
    Dim x As Long, iHeight As Long, R2 As Long, G2 As Long, B2 As Long
    For x = 0 To 100
        DstObject.Line (x, 0)-(x, iHeight), RGB(R2, G2, B2)
    Next x
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());

    // Verify key structural elements
    assert!(text.contains("ForStatement"), "Should contain ForStatement");
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement for Line method"
    );
    assert!(text.contains("ArgumentList"), "Should contain ArgumentList");

    // Verify the minus operator and parentheses are parsed as part of expressions, not Unknown
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );

    // The coordinate pairs should be parsed as part of the argument list
    assert!(text.contains("RGB"), "Should contain RGB function call");
}

#[test]
fn line_statement_on_form() {
    let source = r"
Sub Form_Paint()
    Me.Line (0, 0)-(Me.ScaleWidth, Me.ScaleHeight), vbBlue, B
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(
        text.contains("ScaleWidth"),
        "Should contain ScaleWidth property"
    );
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}

#[test]
fn line_statement_no_color_argument() {
    let source = r"
Sub Test()
    Picture1.Line (50, 50)-(150, 150)
End Sub
";
    let (cst_opt, failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert!(failures.is_empty(), "Expected no parse failures");

    let text = format!("{:#?}", cst.to_serializable());
    assert!(
        text.contains("CallStatement"),
        "Should contain CallStatement"
    );
    assert!(
        !text.contains("Unknown"),
        "Should not contain Unknown tokens"
    );
}
