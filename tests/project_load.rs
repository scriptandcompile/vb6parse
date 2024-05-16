use vb6parse::parser::VB6Project;

#[test]
fn vbp_load() {
    let project_file_bytes = include_bytes!("ppdm.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.references.len(), 15);
    assert_eq!(project.objects.len(), 12);
    assert_eq!(project.modules.len(), 39);
    assert_eq!(project.classes.len(), 83);
    assert_eq!(project.designers.len(), 55);
    assert_eq!(project.forms.len(), 157);
    assert_eq!(project.usercontrols.len(), 13);

    assert_eq!(project.res_file_32_path, "..\\DBCommon\\PSFC.RES");
    assert_eq!(project.icon_form, "frmMain");
    assert_eq!(project.startup, "Sub Main");
}
