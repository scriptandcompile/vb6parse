//! Test utilities for CST assertions
//!
//! This module provides macros and functions to assist in writing tests
//! for the Concrete Syntax Tree (CST) produced by the parser. It includes
//! macros to assert the structure and content of the CST nodes.
//!
//! The main macro `assert_tree!` allows for concise and readable assertions
//! of the CST structure in tests. It uses a hybrid approach that captures
//! line information from the tree specification and reports precise error
//! locations when assertions fail.
//!
//! ## Compile-Time Validation
//!
//! As of version 0.5.1, the `assert_tree!` macro now includes compile-time
//! validation of `SyntaxKind` identifiers. This means typos or invalid node
//! kinds will be caught at compile time with helpful error messages.
//!
//! ### Example Error Messages
//!
//! If you mistype a `SyntaxKind` like `SubStatment` (missing 'e'), you'll get:
//! ```text
//! error[E0599]: no variant or associated item named `SubStatment` found for enum `SyntaxKind`
//!   --> src/your_test.rs:123:9
//!    |
//! 123|         SubStatment {
//!    |         ^^^^^^^^^^^ variant or associated item not found in `SyntaxKind`
//!    |
//! help: there is a variant with a similar name
//!    |
//! 123-         SubStatment {
//! 123+         SubStatement {
//!    |
//! ```
//!
//! # Example
//! ```rust
//! use vb6parse::*;
//!
//! let source = "Sub Test()\nEnd Sub\n";
//!
//! let (cst_opt, _) = ConcreteSyntaxTree::from_text("file.bas", source).unpack();
//! let cst = cst_opt.expect("Failed to parse CST");
//!
//! // This will assert the structure of the CST
//! // but can only be used within module tests due to its internal visibility.
//!
//! // assert_tree!(cst, [
//! //    SubStatement {
//! //        SubKeyword,
//! //        Whitespace (" "),
//! //        Identifier ("Test"),
//! //        ParameterList {
//! //            LeftParenthesis,
//! //            RightParenthesis,
//! //        },
//! //        Newline,
//! //    },
//! //    EndKeyword,
//! //    Whitespace (" "),
//! //    SubKeyword,
//! //    Newline,
//! // ]);
//! ```

// The items in this module are used by the assert_tree! macro but appear unused
// to the compiler because they're only invoked through macro expansion.
#![allow(dead_code)]

use crate::{parsers::cst::CstNode, ConcreteSyntaxTree};

/// Helper macro to validate that a token tree contains valid SyntaxKind identifiers.
///
/// This macro validates each identifier in the tree specification against the
/// `SyntaxKind` enum by attempting to construct the enum variant. If an invalid
/// identifier is used, this will produce a compile error pointing to the exact
/// location of the typo.
#[doc(hidden)]
#[macro_export]
macro_rules! validate_syntax_kinds {
    // Base case: empty
    () => {};

    // Skip commas
    (, $($rest:tt)* ) => {
        $crate::validate_syntax_kinds!($($rest)*);
    };

    // Match identifier followed by text literal: Kind ("text")
    ($kind:ident ( $text:literal ) $($rest:tt)* ) => {
        // Validate this identifier references a SyntaxKind variant
        const _: () = {
            let _validation: $crate::parsers::syntaxkind::SyntaxKind = $crate::parsers::syntaxkind::SyntaxKind::$kind;
        };
        $crate::validate_syntax_kinds!($($rest)*);
    };

    // Match identifier followed by children: Kind { ... }
    ($kind:ident { $($children:tt)* } $($rest:tt)* ) => {
        // Validate this identifier references a SyntaxKind variant
        const _: () = {
            let _validation: $crate::parsers::syntaxkind::SyntaxKind = $crate::parsers::syntaxkind::SyntaxKind::$kind;
        };
        // Recursively validate children
        $crate::validate_syntax_kinds!($($children)*);
        // Continue with remaining siblings
        $crate::validate_syntax_kinds!($($rest)*);
    };

    // Match simple identifier: Kind
    ($kind:ident $($rest:tt)* ) => {
        // Validate this identifier references a SyntaxKind variant
        const _: () = {
            let _validation: $crate::parsers::syntaxkind::SyntaxKind = $crate::parsers::syntaxkind::SyntaxKind::$kind;
        };
        $crate::validate_syntax_kinds!($($rest)*);
    };
}

/// Macro to assert the structure of a CST node against an expected pattern.
///
/// This macro uses a hybrid approach: it captures the tree specification as a string
/// using `stringify!()`, parses it at runtime, and provides detailed error messages
/// that include the specific line from the tree specification where a mismatch occurred.
///
/// This macro now includes compile-time validation of `SyntaxKind` identifiers, which
/// means typos in node kinds will be caught at compile time with helpful error messages
/// pointing to the exact location of the mistake.
///
/// This macro is internal to the crate and not exported.
///
/// # Example Error Messages
///
/// If you mistype a SyntaxKind like `SubStatment` (missing 'e'), you'll get a compile error:
/// ```text
/// error[E0599]: no variant named `SubStatment` found for enum `SyntaxKind`
/// ```
///
/// The error will point to the exact line in your test where the typo appears.
#[macro_export]
macro_rules! assert_tree {
    ($node:expr, [ $($tree:tt)* ]) => {{
        // Validate all SyntaxKind identifiers at compile time
        $crate::validate_syntax_kinds!($($tree)*);

        // Run the actual tree assertion at runtime
        let tree_spec = stringify!($($tree)*);
        if let Err(e) = $crate::test_utils::check_tree(&$node, tree_spec, file!(), line!() as usize) {
            panic!("{}", e);
        }
    }};
}

/// Represents a parsed node in the expected tree structure
#[derive(Debug, Clone)]
enum ExpectedNode {
    /// Simple node: `Kind`
    Simple { kind: String, line_offset: usize },
    /// Node with text: `Kind ("text")`
    WithText {
        kind: String,
        text: String,
        line_offset: usize,
    },
    /// Node with children: `Kind { ... }`
    WithChildren {
        kind: String,
        children: Vec<ExpectedNode>,
        line_offset: usize,
    },
}

impl ExpectedNode {
    fn kind(&self) -> &str {
        match self {
            ExpectedNode::Simple { kind, .. }
            | ExpectedNode::WithText { kind, .. }
            | ExpectedNode::WithChildren { kind, .. } => kind,
        }
    }

    fn line_offset(&self) -> usize {
        match self {
            ExpectedNode::Simple { line_offset, .. }
            | ExpectedNode::WithText { line_offset, .. }
            | ExpectedNode::WithChildren { line_offset, .. } => *line_offset,
        }
    }
}

/// Parse the stringified tree specification into expected nodes
fn parse_tree_spec(spec: &str) -> Result<Vec<ExpectedNode>, String> {
    let mut parser = TreeParser::new(spec);
    parser.parse_nodes()
}

struct TreeParser {
    input: Vec<char>,
    pos: usize,
    line: usize,
}

impl TreeParser {
    fn new(input: &str) -> Self {
        Self {
            input: input.chars().collect(),
            pos: 0,
            line: 0,
        }
    }

    fn current(&self) -> Option<char> {
        self.input.get(self.pos).copied()
    }

    fn advance(&mut self) {
        if let Some('\n') = self.current() {
            self.line += 1;
        }
        self.pos += 1;
    }

    fn skip_whitespace_and_commas(&mut self) {
        while let Some(ch) = self.current() {
            if ch.is_whitespace() || ch == ',' {
                self.advance();
            } else {
                break;
            }
        }
    }

    fn parse_identifier(&mut self) -> Result<String, String> {
        let mut ident = String::new();
        while let Some(ch) = self.current() {
            if ch.is_alphanumeric() || ch == '_' {
                ident.push(ch);
                self.advance();
            } else {
                break;
            }
        }
        if ident.is_empty() {
            Err("Expected identifier".to_string())
        } else {
            Ok(ident)
        }
    }

    fn parse_string_literal(&mut self) -> Result<String, String> {
        if self.current() != Some('"') {
            return Err("Expected opening quote".to_string());
        }
        self.advance(); // skip opening "

        let mut text = String::new();
        while let Some(ch) = self.current() {
            if ch == '"' {
                self.advance(); // skip closing "
                return Ok(text);
            } else if ch == '\\' {
                self.advance();
                if let Some(escaped) = self.current() {
                    text.push(match escaped {
                        'n' => '\n',
                        't' => '\t',
                        'r' => '\r',
                        '\\' => '\\',
                        '"' => '"',
                        _ => escaped,
                    });
                    self.advance();
                }
            } else {
                text.push(ch);
                self.advance();
            }
        }
        Err("Unterminated string literal".to_string())
    }

    fn parse_nodes(&mut self) -> Result<Vec<ExpectedNode>, String> {
        let mut nodes = Vec::new();
        self.skip_whitespace_and_commas();

        while self.current().is_some() && self.current() != Some('}') {
            let node = self.parse_node()?;
            nodes.push(node);
            self.skip_whitespace_and_commas();
        }

        Ok(nodes)
    }

    fn parse_node(&mut self) -> Result<ExpectedNode, String> {
        self.skip_whitespace_and_commas();
        let line_offset = self.line;
        let kind = self.parse_identifier()?;

        self.skip_whitespace_and_commas();

        match self.current() {
            Some('(') => {
                // Node with text
                self.advance(); // skip '('
                self.skip_whitespace_and_commas();
                let text = self.parse_string_literal()?;
                self.skip_whitespace_and_commas();
                if self.current() == Some(')') {
                    self.advance(); // skip ')'
                }
                Ok(ExpectedNode::WithText {
                    kind,
                    text,
                    line_offset,
                })
            }
            Some('{') => {
                // Node with children
                self.advance(); // skip '{'
                let children = self.parse_nodes()?;
                if self.current() == Some('}') {
                    self.advance(); // skip '}'
                }
                Ok(ExpectedNode::WithChildren {
                    kind,
                    children,
                    line_offset,
                })
            }
            _ => {
                // Simple node
                Ok(ExpectedNode::Simple { kind, line_offset })
            }
        }
    }
}

/// Check a CST node against the expected tree specification
#[allow(clippy::cast_possible_truncation)]
pub(crate) fn check_tree(
    tree: &ConcreteSyntaxTree,
    tree_spec: &str,
    file: &str,
    base_line: usize,
) -> Result<(), String> {
    let node = tree.to_root_node();
    let expected_nodes = parse_tree_spec(tree_spec).map_err(|e| {
        format!(
            "Failed to parse tree specification: {e}\nSpec:\n{tree_spec}\nCST Debug:\n{}\n",
            tree.debug_tree()
        )
    })?;

    check_nodes(node.children(), &expected_nodes, file, base_line, &["Root"])
        .map_err(|e| format!("{}\nCST Debug:\n{}\n", e, tree.debug_tree()))?;

    // Check if there are extra children
    if node.children().len() > expected_nodes.len() {
        let extra_count = node.children().len() - expected_nodes.len();
        return Err(format!(
            "\n{file}:{base_line}: Assertion failed in tree structure\nPath: Root\nExpected {} children, but found {} ({extra_count}extra)\nCST Debug:\n{}\n",
            expected_nodes.len(),
            node.children().len(),
            tree.debug_tree()
        ));
    }

    Ok(())
}

#[allow(clippy::cast_possible_truncation)]
fn check_nodes(
    actual_nodes: &[CstNode],
    expected_nodes: &[ExpectedNode],
    file: &str,
    base_line: usize,
    path: &[&str],
) -> Result<(), String> {
    for (idx, expected) in expected_nodes.iter().enumerate() {
        let actual = actual_nodes.get(idx).ok_or_else(|| {
            format!(
                "\n{file}:{}: Assertion failed in tree structure\nPath: {}\nExpected child at index {idx}: {}\nBut node has only {} children\n",
                base_line + expected.line_offset() as usize + 1,
                path.join(" → "),
                expected.kind(),
                actual_nodes.len(),
            )
        })?;

        check_single_node(actual, expected, idx, file, base_line, path)?;
    }

    Ok(())
}

#[allow(clippy::cast_possible_truncation)]
fn check_single_node(
    actual: &CstNode,
    expected: &ExpectedNode,
    idx: usize,
    file: &str,
    base_line: usize,
    path: &[&str],
) -> Result<(), String> {
    let expected_kind_str = expected.kind();
    let actual_kind = actual.kind();
    let actual_kind_str = format!("{actual_kind:?}");

    // Compare kinds
    if expected_kind_str != actual_kind_str {
        return Err(format!(
            "\n{file}:{}: Assertion failed in tree structure\nPath: {} → [{idx}]\nExpected: {expected_kind_str}\nActual:   {actual_kind_str}\n",
            base_line + expected.line_offset() as usize + 1,
            path.join(" → "),
        ));
    }

    match expected {
        ExpectedNode::WithText {
            text: expected_text,
            kind,
            ..
        } => {
            if !actual.is_token() {
                return Err(format!(
                    "\n{file}:{}: Assertion failed in tree structure\nPath: {} → {kind} [{idx}]\nExpected a token node with text\nActual:   non-token node\n",
                    base_line + expected.line_offset() as usize + 1,
                    path.join(" → "),
                ));
            }

            let actual_text = actual.text();
            if expected_text != actual_text {
                return Err(format!(
                    "\n{file}:{}: Assertion failed in tree structure\nPath: {} → {kind} [{idx}]\nExpected text: {expected_text:?}\nActual text:   {actual_text:?}\n",
                    base_line + expected.line_offset() as usize + 1,
                    path.join(" → "),
                ));
            }
        }
        ExpectedNode::WithChildren { kind, children, .. } => {
            let mut new_path = path.to_vec();
            new_path.push(kind);
            check_nodes(actual.children(), children, file, base_line, &new_path)?;

            // Check for extra children
            if actual.children().len() > children.len() {
                let extra_count = actual.children().len() - children.len();
                return Err(format!(
                    "\n{file}:{}: Assertion failed in tree structure\nPath: {} → {kind}\nExpected {} children, but found {} ({extra_count} extra)\n",
                    base_line + expected.line_offset() as usize + 1,
                    path.join(" → "),
                    children.len(),
                    actual.children().len(),
                ));
            }
        }
        ExpectedNode::Simple { .. } => {
            // Kind already checked, nothing more to do
        }
    }

    Ok(())
}
