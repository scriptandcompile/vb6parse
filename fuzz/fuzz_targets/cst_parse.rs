#![no_main]
use libfuzzer_sys::fuzz_target;
use vb6parse::SourceFile;

fuzz_target!(|data: &[u8]| {
    if let Ok(source_file) = SourceFile::decode_with_replacement("fuzz.bas", data) {
        let result = vb6parse::parsers::cst::ConcreteSyntaxTree::from_source(&source_file);

        let (cst_opt, _failures) = result.unpack();

        // If we got a CST, walk the tree to ensure no panics
        if let Some(cst) = cst_opt {
            let root = cst.to_serializable().root;

            // Recursively walk the tree
            fn walk_node(node: &vb6parse::parsers::cst::CstNode) {
                let _ = node.kind;
                let _ = node.text.as_str();
                let _ = node.is_token;

                for child in &node.children {
                    walk_node(child);
                }
            }

            walk_node(&root);

            // Also test basic tree properties
            let _ = cst.root_kind();
            let _ = cst.text();
        }
    }
});
