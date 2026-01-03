//! `DefType` statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 `DefType` statements which set default data types
//! for variables based on their first letter.
//!
//! `DefType` statements include:
//! - `DefBool`: Boolean type
//! - `DefByte`: Byte type
//! - `DefInt`: Integer type
//! - `DefLng`: Long type
//! - `DefCur`: Currency type
//! - `DefSng`: Single type
//! - `DefDbl`: Double type
//! - `DefDec`: Decimal type
//! - `DefDate`: Date type
//! - `DefStr`: String type
//! - `DefObj`: Object type
//! - `DefVar`: Variant type
//!
//! Syntax: `DefType` letterrange [, letterrange] ...
//! where letterrange is a single letter or a range like A-Z
//!
//! [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Visual Basic 6 `DefType` statement with syntax:
    ///
    /// `DefType` letterrange \[, letterrange\] ...
    ///
    /// The `DefType` statement syntax has these parts:
    ///
    /// | Part          | Description |
    /// |---------------|-------------|
    /// | `DefBool`       | Sets default type to Boolean for variables starting with specified letters. |
    /// | `DefByte`       | Sets default type to Byte for variables starting with specified letters. |
    /// | `DefInt`        | Sets default type to Integer for variables starting with specified letters. |
    /// | `DefLng`        | Sets default type to Long for variables starting with specified letters. |
    /// | `DefCur`        | Sets default type to Currency for variables starting with specified letters. |
    /// | `DefSng`        | Sets default type to Single for variables starting with specified letters. |
    /// | `DefDbl`        | Sets default type to Double for variables starting with specified letters. |
    /// | `DefDec`        | Sets default type to Decimal for variables starting with specified letters. |
    /// | `DefDate`       | Sets default type to Date for variables starting with specified letters. |
    /// | `DefStr`        | Sets default type to String for variables starting with specified letters. |
    /// | `DefObj`        | Sets default type to Object for variables starting with specified letters. |
    /// | `DefVar`        | Sets default type to Variant for variables starting with specified letters. |
    /// | letterrange   | A single letter or range of letters (e.g., A, A-Z, M-P). Multiple ranges separated by commas. |
    ///
    /// Examples:
    /// - `DefInt` A-Z (all variables default to Integer)
    /// - `DefStr` S (variables starting with S default to String)
    /// - `DefLng` L, M-N (variables starting with L, M, or N default to Long)
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    pub(super) fn parse_deftype_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::DefTypeStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume the DefType keyword (DefBool, DefByte, DefInt, etc.)
        // TODO: Validate that the keyword is one of the valid DefType keywords
        self.consume_token();

        // Consume any whitespace after DefType keyword
        self.consume_whitespace();

        // TODO: Validate letter ranges

        // Parse letter ranges until newline
        // Letter ranges can be:
        // - Single letter: A
        // - Range: A-Z
        // - Multiple ranges separated by commas: A, M-Z
        loop {
            // Check if we've reached the end of the line
            if self.at_token(Token::Newline) || self.is_at_end() {
                break;
            }

            // Consume the letter or range
            // This includes identifiers (for letters) and minus signs (for ranges)
            self.consume_token();
        }

        // Consume the newline
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // DefTypeStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn deftype_defint_single_letter() {
        // Test DefInt with single letter
        let source = "DefInt I\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefIntKeyword,
                Whitespace,
                Identifier ("I"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defint_range() {
        // Test DefInt with letter range
        let source = "DefInt A-Z\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefIntKeyword,
                Whitespace,
                Identifier ("A"),
                SubtractionOperator,
                Identifier ("Z"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_deflng_multiple_ranges() {
        // Test DefLng with multiple ranges
        let source = "DefLng L, M-N\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefLngKeyword,
                Whitespace,
                Identifier ("L"),
                Comma,
                Whitespace,
                Identifier ("M"),
                SubtractionOperator,
                Identifier ("N"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defstr_single() {
        // Test DefStr with single letter
        let source = "DefStr S\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefStrKeyword,
                Whitespace,
                Identifier ("S"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defdbl_range() {
        // Test DefDbl with range
        let source = "DefDbl D-F\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefDblKeyword,
                Whitespace,
                Identifier ("D"),
                SubtractionOperator,
                Identifier ("F"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defobj_full_range() {
        // Test DefObj A-Z (common pattern)
        let source = "DefObj A-Z\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefObjKeyword,
                Whitespace,
                Identifier ("A"),
                SubtractionOperator,
                Identifier ("Z"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defbool() {
        // Test DefBool
        let source = "DefBool B\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefBoolKeyword,
                Whitespace,
                Identifier ("B"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defbyte() {
        // Test DefByte
        let source = "DefByte B\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefByteKeyword,
                Whitespace,
                Identifier ("B"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defcur() {
        // Test DefCur
        let source = "DefCur C\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefCurKeyword,
                Whitespace,
                Identifier ("C"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defsng() {
        // Test DefSng
        let source = "DefSng F-G\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefSngKeyword,
                Whitespace,
                Identifier ("F"),
                SubtractionOperator,
                Identifier ("G"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defdec() {
        // Test DefDec
        let source = "DefDec D\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefDecKeyword,
                Whitespace,
                Identifier ("D"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defdate() {
        // Test DefDate
        let source = "DefDate D\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefDateKeyword,
                Whitespace,
                Identifier ("D"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_defvar() {
        // Test DefVar
        let source = "DefVar V-Z\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefVarKeyword,
                Whitespace,
                Identifier ("V"),
                SubtractionOperator,
                Identifier ("Z"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_multiple_single_letters() {
        // Test multiple single letters
        let source = "DefInt A, B, C\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefIntKeyword,
                Whitespace,
                Identifier ("A"),
                Comma,
                Whitespace,
                Identifier ("B"),
                Comma,
                Whitespace,
                Identifier ("C"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_mixed_ranges_and_singles() {
        // Test mixed ranges and single letters
        let source = "DefLng A-C, E, G-Z\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefLngKeyword,
                Whitespace,
                Identifier ("A"),
                SubtractionOperator,
                Identifier ("C"),
                Comma,
                Whitespace,
                Identifier ("E"),
                Comma,
                Whitespace,
                Identifier ("G"),
                SubtractionOperator,
                Identifier ("Z"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_multiple_statements() {
        // Test multiple DefType statements
        let source = "DefInt I-N\nDefLng L\nDefStr S\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefIntKeyword,
                Whitespace,
                Identifier ("I"),
                SubtractionOperator,
                Identifier ("N"),
                Newline,
            },
            DefTypeStatement {
                DefLngKeyword,
                Whitespace,
                Identifier ("L"),
                Newline,
            },
            DefTypeStatement {
                DefStrKeyword,
                Whitespace,
                Identifier ("S"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_with_spaces() {
        // Test with various spacing
        let source = "DefInt  A - Z\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefIntKeyword,
                Whitespace,
                Identifier ("A"),
                Whitespace,
                SubtractionOperator,
                Whitespace,
                Identifier ("Z"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_lowercase_range() {
        // Test with lowercase letters (should still work)
        let source = "DefStr a-z\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefStrKeyword,
                Whitespace,
                Identifier ("a"),
                SubtractionOperator,
                Identifier ("z"),
                Newline,
            },
        ]);
    }

    #[test]
    fn deftype_partial_alphabet() {
        // Test partial alphabet ranges
        let source = "DefInt A-M\nDefLng N-Z\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DefTypeStatement {
                DefIntKeyword,
                Whitespace,
                Identifier ("A"),
                SubtractionOperator,
                Identifier ("M"),
                Newline,
            },
            DefTypeStatement {
                DefLngKeyword,
                Whitespace,
                Identifier ("N"),
                SubtractionOperator,
                Identifier ("Z"),
                Newline,
            },
        ]);
    }
}
