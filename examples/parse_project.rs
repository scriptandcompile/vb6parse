use vb6parse::{ObjectReference, ProjectFile, ProjectReference, SourceFile};

/// Example showing how to parse a VB6 project file from raw bytes.
/// This example uses a hardcoded string, but in a real application,
/// you would typically read the bytes from a `.vbp` file on disk.
fn main() {
    // Hardcoded example of a VB6 project file content as a string.
    // In a real application, you would read this from a file.
    let input = r#"Type=Exe
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
Module=Module1; Module1.bas
Class=Class1; Class1.cls
Form=Form1.frm
Form=Form2.frm
UserControl=UserControl1.ctl
UserDocument=UserDocument1.uds
ExeName32="Project1.exe"
Command32=""
Path32=""
Name="Project1"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
DebugStartupOption=0
NoControlUpgrade=0
ServerSupportFiles=0
VersionCompanyName="Company Name"
VersionFileDescription="File Description"
VersionLegalCopyright="Copyright"
VersionLegalTrademarks="Trademark"
VersionProductName="Product Name"
VersionComments="Comments"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
CondComp=""
ResFile32=""
IconForm=""
Startup="Form1"
HelpFile=""
Title="Project1"

[MS Transaction Server]
AutoRefresh=1
"#;

    // Decode the source file from the byte array.
    // The filename is provided for reference in error messages.
    // In a real application, use the actual filename.
    // Decode with replacement to handle any invalid characters gracefully.
    let source_file =
        SourceFile::decode_with_replacement("Project1.vbp", input.as_bytes()).unwrap();

    // Parse the project file from the decoded source file.
    let project = ProjectFile::parse(&source_file).unwrap_or_fail();

    // Print out some information about the parsed project file.
    println!("Project Name: {}", project.properties.name);
    println!("Project Type: {:?}", project.project_type);
    println!("Files in Project:");
    for reference in project.references() {
        match reference {
            ProjectReference::Compiled {
                uuid,
                unknown1,
                unknown2,
                path,
                description,
            } => {
                println!("Reference - Compiled: {uuid} {unknown1} {unknown2} {path} {description}");
            }
            ProjectReference::SubProject { path } => {
                println!("Reference - SubProject: {path}",);
            }
        }
    }
    for object in project.objects() {
        match object {
            ObjectReference::Compiled {
                uuid,
                version,
                unknown1,
                file_name,
            } => {
                println!("Object Reference - Compiled: {uuid} {version} {unknown1} {file_name}");
            }
            ObjectReference::Project { path } => {
                println!("Object Reference - Project: {path}");
            }
        }
    }
    for module in project.modules() {
        println!("Module: {} (Path: {})", module.name, module.path);
    }
    for class in project.classes() {
        println!("Class: {} (Path: {})", class.name, class.path);
    }
    for related_document_path in project.related_documents() {
        println!("Related Document - Path: {related_document_path}");
    }
    for user_document_path in project.user_documents() {
        println!("User Document - Path: {user_document_path}");
    }
    for form_path in project.forms() {
        println!("Form - Path: {form_path}");
    }
    for user_control_path in project.user_controls() {
        println!("User Control - Path: {user_control_path}");
    }
    for property_group in &project.other_properties {
        println!("Property Group: {}", property_group.0);
        for property in property_group.1 {
            println!("  {} = {}", property.0, property.1);
        }
    }
}
