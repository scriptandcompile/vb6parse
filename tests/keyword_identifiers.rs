//! Tests for keywords used as identifiers in various positions
//!
//! VB6 allows keywords to be used as identifiers (variable names, procedure names, etc.)
//! in most contexts. This test file verifies that keywords are properly converted to
//! Identifier tokens when they appear in identifier positions.

use vb6parse::*;

#[test]
fn keyword_as_sub_name() {
    let source = "Sub Text()\nEnd Sub\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert_tree!(cst, [
        SubStatement {
            SubKeyword,
            Whitespace,
            Identifier ("Text"),
            ParameterList {
                LeftParenthesis,
                RightParenthesis,
            },
            Newline,
            StatementList,
            EndKeyword,
            Whitespace,
            SubKeyword,
            Newline,
        },
    ]);
}

#[test]
fn keyword_as_function_name() {
    let source = "Function Database() As String\nEnd Function\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert_tree!(cst, [
        FunctionStatement {
            FunctionKeyword,
            Whitespace,
            Identifier ("Database"),
            ParameterList {
                LeftParenthesis,
                RightParenthesis,
            },
            Whitespace,
            AsKeyword,
            Whitespace,
            StringKeyword,
            Newline,
            StatementList,
            EndKeyword,
            Whitespace,
            FunctionKeyword,
            Newline,
        },
    ]);
}

#[test]
fn keyword_as_property_name() {
    let source = "Property Get Binary() As Integer\nEnd Property\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert_tree!(cst, [
        PropertyStatement {
            PropertyKeyword,
            Whitespace,
            GetKeyword,
            Whitespace,
            Identifier ("Binary"),
            ParameterList {
                LeftParenthesis,
                RightParenthesis,
            },
            Whitespace,
            AsKeyword,
            Whitespace,
            IntegerKeyword,
            Newline,
            StatementList,
            EndKeyword,
            Whitespace,
            PropertyKeyword,
            Newline,
        },
    ]);
}

#[test]
fn keyword_as_variable_in_assignment() {
    let source = "text = \"hello\"\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert_tree!(cst, [
        AssignmentStatement {
            IdentifierExpression {
                TextKeyword,
            },
            Whitespace,
            EqualityOperator,
            Whitespace,
            StringLiteralExpression {
                StringLiteral ("\"hello\""),
            },
            Newline,
        },
    ]);
}

#[test]
fn keyword_as_property_in_assignment() {
    let source = "obj.text = \"hello\"\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert_tree!(cst, [
        AssignmentStatement {
            MemberAccessExpression {
                Identifier ("obj"),
                PeriodOperator,
                TextKeyword,
            },
            Whitespace,
            EqualityOperator,
            Whitespace,
            StringLiteralExpression {
                StringLiteral ("\"hello\""),
            },
            Newline,
        },
    ]);
}

#[test]
fn multiple_keywords_as_identifiers() {
    let source = r#"
database = "mydb.mdb"
text = "hello"
obj.binary = True
"#;
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert_tree!(cst, [
        Newline,
        AssignmentStatement {
            IdentifierExpression {
                DatabaseKeyword,
            },
            Whitespace,
            EqualityOperator,
            Whitespace,
            StringLiteralExpression {
                StringLiteral ("\"mydb.mdb\""),
            },
            Newline,
        },
        AssignmentStatement {
            IdentifierExpression {
                TextKeyword,
            },
            Whitespace,
            EqualityOperator,
            Whitespace,
            StringLiteralExpression {
                StringLiteral ("\"hello\""),
            },
            Newline,
        },
        AssignmentStatement {
            MemberAccessExpression {
                Identifier ("obj"),
                PeriodOperator,
                BinaryKeyword,
            },
            Whitespace,
            EqualityOperator,
            Whitespace,
            BooleanLiteralExpression {
                TrueKeyword,
            },
            Newline,
        },
    ]);
}

#[test]
fn keyword_as_enum_name() {
    let source = "Enum Random\n    Value1\n    Value2\nEnd Enum\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert_tree!(cst, [
        EnumStatement {
            EnumKeyword,
            Whitespace,
            Identifier ("Random"),
            Newline,
            Whitespace,
            Identifier ("Value1"),
            Newline,
            Whitespace,
            Identifier ("Value2"),
            Newline,
            EndKeyword,
            Whitespace,
            EnumKeyword,
            Newline,
        },
    ]);
}

#[test]
fn keyword_after_keyword_converted() {
    // Even when a keyword follows another keyword in procedure definition
    let source = "Sub Output()\nEnd Sub\n";
    let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
    let cst = cst_opt.expect("CST should be parsed");

    assert_tree!(cst, [
        SubStatement {
            SubKeyword,
            Whitespace,
            Identifier ("Output"),
            ParameterList {
                LeftParenthesis,
                RightParenthesis,
            },
            Newline,
            StatementList,
            EndKeyword,
            Whitespace,
            SubKeyword,
            Newline,
        },
    ]);
}
