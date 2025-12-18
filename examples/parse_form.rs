use vb6parse::language::ControlKind;
use vb6parse::parsers::FormFile;
use vb6parse::SourceFile;

fn main() {
    let input = b"VERSION 5.00
Begin VB.Form Form1
    BorderStyle     =   1  'Fixed Single
    Caption         =   \"Form1\"
    ClientHeight    =   3195
    ClientLeft      =   60
    ClientTop       =   345
    ClientWidth     =   4680
    LinkTopic       =   \"\"
    MaxButton       =   1   'True
    MinButton       =   1   'True
    ScaleHeight     =   3195
    ScaleWidth      =   4680
    StartUpPosition =   3  'Windows Default
    BeginProperty Font
        Name            =   \"MS Sans Serif\"
        Size            =   8.25
        Charset         =   0
        Weight          =   400
        Underline       =   0   'False
        Italic          =   0   'False
        Strikethrough   =   0   'False
    EndProperty
    Begin VB.TextBox Text1
        Height          =   315
        Left            =   120
        TabIndex        =   0
        Top             =   720
        Width           =   1215
    End
    Begin VB.CommandButton Command1
        Caption         =   \"Command1\"
        Height          =   375
        Left            =   1560
        TabIndex        =   1
        Top             =   720
        Width           =   1215
    End
End
Attribute VB_Name = \"Form1\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Some comment";

    let result = SourceFile::decode_with_replacement("form_parse.frm", input);

    let source_file = match result {
        Ok(source_file) => source_file,
        Err(e) => panic!("Failed to decode source file 'form_parse.frm': {e:?}"),
    };

    let form_file = FormFile::parse(&source_file).unwrap_or_fail();

    println!("Form Version Major: {}", form_file.version.major);
    println!("Form Version Minor: {}", form_file.version.minor);
    println!("Form Properties:");

    let (properties, controls, menus) = match form_file.form.kind {
        ControlKind::Form {
            properties,
            controls,
            menus,
        } => (properties, controls, menus),
        _ => panic!("Unexpected control kind"),
    };

    println!("\tCaption: {:?}", properties.caption);
    println!("\tBorder Style: {:?}", properties.border_style);
    println!("\tMax Button: {:?}", properties.max_button);
    println!("\tMin Button: {:?}", properties.min_button);
    println!("\tStartup Position: {:?}", properties.start_up_position);
    println!("\tClient Height: {:?}", properties.client_height);
    println!("\tClient Width: {:?}", properties.client_width);
    println!("\tLink Topic: {:?}", properties.link_topic);
    println!("\tNumber of Controls: {}", controls.len());
    println!("\tNumber of Menus: {}", menus.len());
    println!("Form Attributes:");
    println!("\tName: {:?}", form_file.attributes.name);
    println!(
        "\tGlobal Name Space: {:?}",
        form_file.attributes.global_name_space
    );
    println!("\tCreatable: {:?}", form_file.attributes.creatable);
    println!(
        "\tPredeclared ID: {:?}",
        form_file.attributes.predeclared_id
    );
    println!("\tExposed: {:?}", form_file.attributes.exposed);
    println!("Ext Attributes:");
    for ext in &form_file.attributes.ext_key {
        println!("\t{} = {}", ext.0, ext.1);
    }

    println!("CST:");
    println!("{}", form_file.cst.debug_tree());
}
