use vb6parse::io::SourceFile;
use vb6parse::lexer::tokenize;
use vb6parse::parsers::cst::parse;

fn main() {
    let source = b"VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   \"Test Form\"
End
Attribute VB_Name = \"frmMain\"
Attribute VB_GlobalNameSpace = False
";

    let source_file = SourceFile::decode_with_replacement("test.frm", source).unwrap();
    let mut stream = source_file.source_stream();
    let tokens = tokenize(&mut stream).unwrap();
    let cst = parse(tokens);

    println!("Full CST structure:");
    println!("{}", cst.debug_tree());
}
