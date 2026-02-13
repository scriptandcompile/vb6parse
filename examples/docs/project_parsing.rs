use vb6parse::*;

fn main() {
    let project_content = r#"Type=Exe
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
Module=Utilities; Utilities.bas
Module=DataAccess; DataAccess.bas
Form=MainForm.frm
Class=DatabaseConnection; DatabaseConnection.cls
"#;

    let source = SourceFile::from_string("MyProject.vbp", project_content);
    let result = ProjectFile::parse(&source);
    let (project, _failures) = result.unpack();

    if let Some(proj) = project {
        println!("Project Type: {:?}", proj.project_type);

        // Iterate over modules
        println!("\nModules:");
        for module in proj.modules() {
            println!("  - {} ({})", module.name, module.path);
        }

        // Iterate over forms
        println!("\nForms:");
        for form_name in proj.forms() {
            println!("  - {form_name}");
        }

        // Iterate over classes
        println!("\nClasses:");
        for class in proj.classes() {
            println!("  - {} ({})", class.name, class.path);
        }

        // Check references
        println!("\nReferences: {}", proj.references().count());
    }
}
