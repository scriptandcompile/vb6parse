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
    assert_eq!(project.user_controls.len(), 13);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(
        project.res_file_32_path,
        Some("..\\DBCommon\\PSFC.RES".to_owned())
    );
    assert_eq!(project.icon_form, Some("frmMain".to_owned()));
    assert_eq!(project.startup, Some("Sub Main".to_owned()));
    assert_eq!(project.help_file_path, Some("".to_owned()));
    assert_eq!(project.title, Some("PPDM".to_owned()));
    assert_eq!(project.exe_32_file_name, Some("PPDM.exe".to_owned()));
    assert_eq!(
        project.command_line_arguments,
        Some("-DisableRememberPassword%20-CHARTING -U -language %22english%7d".to_owned())
    );
    assert_eq!(project.name, Some("PPDM".to_owned()));
    assert_eq!(project.help_context_id, Some("0".to_owned()));
    assert_eq!(project.compatible_mode, Some("0".to_owned()));
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.major_version, Some("11".to_owned()));
    assert_eq!(project.minor_version, Some("0".to_owned()));
    assert_eq!(project.revision_version, Some("288".to_owned()));
    assert_eq!(project.auto_increment_revision_version, false);
}
