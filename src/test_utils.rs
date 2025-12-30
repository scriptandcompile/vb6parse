//! Test utilities for CST assertions
//!
//! This module provides macros and functions to assist in writing tests
//! for the Concrete Syntax Tree (CST) produced by the parser. It includes
//! macros to assert the structure and content of the CST nodes.
//!
//! The main macro `assert_tree!` allows for concise and readable assertions
//! of the CST structure in tests.
//!
//! # Example
//! ```rust
//! use vb6parse::parsers::cst::{ConcreteSyntaxTree, SyntaxKind};
//! use vb6parse::test_utils::assert_tree;
//!
//! let cst = ConcreteSyntaxTree::parse("Sub Test()\nEnd Sub");
//! assert_tree!(cst.to_root_node(), [
//!     SyntaxKind::Newline,
//!     SyntaxKind::SubStatement {
//!         SyntaxKind::SubKeyword,
//!         SyntaxKind::Whitespace,
//!         SyntaxKind::Identifier ("Test"),
//!         SyntaxKind::ParameterList {
//!             SyntaxKind::LeftParenthesis,
//!             SyntaxKind::RightParenthesis,
//!         },
//!         SyntaxKind::Newline,
//!     },
//!     SyntaxKind::EndSubStatement {
//!         SyntaxKind::EndSubKeyword,
//!         SyntaxKind::Whitespace,
//!         SyntaxKind::SubKeyword,
//!     },
//! ]);
//! ```
//!

/// Macro to assert the structure of a CST node against an expected pattern.
///
#[macro_export]
macro_rules! assert_tree {
    ($node:expr, [ $($tree:tt)* ]) => {
        $crate::__assert_tree_internal!($node, 0, $($tree)*);
    };
}

/// Internal helper macro for `assert_tree!`.
///
#[macro_export]
macro_rules! __assert_tree_internal {
    // Node with children: Kind { ... },
    ($node:expr, $idx:expr, $kind:ident { $($inner:tt)* }, $($rest:tt)*) => {{
        let children = &$node.children;
        if let Some(child) = children.get($idx) {
            assert_eq!(child.kind, SyntaxKind::$kind, "Expected kind {:?} at line {}", SyntaxKind::$kind, line!());
            $crate::assert_tree!(child, [ $($inner)* ]);
        } else {
            panic!("Missing child {} for kind {:?}", $idx, SyntaxKind::$kind);
        }
        $crate::__assert_tree_internal!($node, $idx + 1, $($rest)*);
    }};
    // Node with children: Kind { ... } (no trailing comma)
    ($node:expr, $idx:expr, $kind:ident { $($inner:tt)* }) => {{
        let children = $node.children;
        if let Some(child) = children.get($idx) {
            assert_eq!(child.kind, SyntaxKind::$kind, "Expected kind {:?} at line {}", SyntaxKind::$kind, line!());
            $crate::assert_tree!(child, [ $($inner)* ]);
        } else {
            panic!("Missing child {} for kind {:?}", $idx, SyntaxKind::$kind);
        }
        $crate::__assert_tree_internal!($node, $idx + 1, );
    }};
    // Token with text: Kind("text"),
    ($node:expr, $idx:expr, $kind:ident ( $text:expr ), $($rest:tt)*) => {{
        let children = &$node.children;
        if let Some(child) = children.get($idx) {
            assert_eq!(child.kind, SyntaxKind::$kind, "Expected kind {:?} at line {}", SyntaxKind::$kind, line!());
            assert!(child.is_token, "Expected token node for text assertion at line {}", line!());
            assert_eq!(child.text, $text, "Expected token text '{}' at line {}", $text, line!());
        } else {
            panic!("Missing child {} for kind {:?}", $idx, SyntaxKind::$kind);
        }
        $crate::__assert_tree_internal!($node, $idx + 1, $($rest)*);
    }};
    // Token with text: Kind("text") (no trailing comma)
    ($node:expr, $idx:expr, $kind:ident ( $text:expr )) => {{
        let children = $node.children;
        if let Some(child) = children.get($idx) {
            assert_eq!(child.kind(), SyntaxKind::$kind, "Expected kind {:?} at line {}", SyntaxKind::$kind, line!());
            assert!(child.is_token(), "Expected token node for text assertion at line {}", line!());
            assert_eq!(child.text(), $text, "Expected token text '{}' at line {}", $text, line!());
        } else {
            panic!("Missing child {} for kind {:?}", $idx, SyntaxKind::$kind);
        }
        $crate::__assert_tree_internal!($node, $idx + 1, );
    }};
    // Node kind only: Kind,
    ($node:expr, $idx:expr, $kind:ident , $($rest:tt)*) => {{
        let children = &$node.children;
        if let Some(child) = children.get($idx) {
            assert_eq!(child.kind, SyntaxKind::$kind, "Expected kind {:?} at line {}", SyntaxKind::$kind, line!());
        } else {
            panic!("Missing child {} for kind {:?}", $idx, SyntaxKind::$kind);
        }
        $crate::__assert_tree_internal!($node, $idx + 1, $($rest)*);
    }};
    // Node kind only: Kind (no trailing comma, end of list)
    ($node:expr, $idx:expr, $kind:ident) => {{
        let children = $node.children;
        if let Some(child) = children.get($idx) {
            assert_eq!(child.kind(), SyntaxKind::$kind, "Expected kind {:?} at line {}", SyntaxKind::$kind, line!());
        } else {
            panic!("Missing child {} for kind {:?}", $idx, SyntaxKind::$kind);
        }
        $crate::__assert_tree_internal!($node, $idx + 1, );
    }};
    // End of list
    ($node:expr, $idx:expr,) => {};
}
