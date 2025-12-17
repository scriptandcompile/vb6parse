use vb6parse::*;

#[test]
fn ppdm_project_load() {
    let project_file_bytes = include_bytes!("./data/ppdm/ppdm.vbp").as_slice();

    let source_file = SourceFile::decode_with_replacement("ppdm.vbp", project_file_bytes).unwrap();

    let result = ProjectFile::parse(&source_file);

    if result.has_failures() {
        for failure in result.failures {
            failure.print();
        }

        panic!("Project parse had failures");
    }

    let project = result.unwrap();

    assert_eq!(project.project_type, CompileTargetType::Exe);
    assert_eq!(project.references.len(), 15);
    assert_eq!(project.objects.len(), 11);
    assert_eq!(project.modules.len(), 34);
    assert_eq!(project.classes.len(), 76);
    assert_eq!(project.designers.len(), 54);
    assert_eq!(project.forms.len(), 156);
    assert_eq!(project.user_controls.len(), 13);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(
        project.properties.res_file_32_path,
        "..\\DBCommon\\PSFC.RES"
    );
    assert_eq!(project.properties.icon_form, "frmMain");
    assert_eq!(project.properties.startup, "Sub Main");
    assert_eq!(project.properties.help_file_path, "");
    assert_eq!(project.properties.title, "PPDM");
    assert_eq!(project.properties.exe_32_file_name, "PPDM.exe");
    assert_eq!(
        project.properties.command_line_arguments,
        "-DisableRememberPassword  -CHARTING -U"
    );
    assert_eq!(project.properties.name, "PPDM");
    assert_eq!(project.properties.help_context_id, "0");
    assert_eq!(
        project.properties.compatibility_mode,
        CompatibilityMode::NoCompatibility
    );
    assert_eq!(
        project.properties.upgrade_controls,
        UpgradeControls::Upgrade
    );
    assert_eq!(
        project.properties.server_support_files,
        ServerSupportFiles::Local
    );
    assert_eq!(
        project.properties.conditional_compile,
        "PDMBuild = 1 : PDM_SHORTCUTS = 1"
    );
    assert_eq!(
        project.properties.compilation_type,
        CompilationType::NativeCode(NativeCodeSettings {
            optimization_type: OptimizationType::FavorFastCode,
            favor_pentium_pro: FavorPentiumPro::False,
            code_view_debug_info: CodeViewDebugInfo::NotCreated,
            aliasing: Aliasing::AssumeAliasing,
            bounds_check: BoundsCheck::CheckBounds,
            overflow_check: OverflowCheck::CheckOverflow,
            floating_point_check: FloatingPointErrorCheck::CheckFloatingPointError,
            pentium_fdiv_bug_check: PentiumFDivBugCheck::CheckPentiumFDivBug,
            unrounded_floating_point: UnroundedFloatingPoint::DoNotAllow,
        })
    );

    assert_eq!(project.properties.start_mode, StartMode::StandAlone);
    assert_eq!(project.properties.unattended, InteractionMode::Interactive);
    assert_eq!(project.properties.retained, Retained::UnloadOnExit);
    assert_eq!(project.properties.thread_per_object, 0);
    assert_eq!(project.properties.max_number_of_threads, 1);
    assert_eq!(
        project.properties.debug_startup_option,
        DebugStartupOption::WaitForComponentCreation
    );

    // version information.
    assert_eq!(project.properties.version_info.major, 11);
    assert_eq!(project.properties.version_info.minor, 0);
    assert_eq!(project.properties.version_info.revision, 177);
    assert_eq!(project.properties.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.properties.version_info.company_name,
        "Predator Software Inc."
    );
    assert_eq!(
        project.properties.version_info.file_description,
        "Predator PDM "
    );
    assert_eq!(
        project.properties.version_info.copyright,
        "Copyright ©1994 - 2017 Predator Software Inc.  All Rights Reserved."
    );
    assert_eq!(
        project.properties.version_info.trademark,
        "Predator SFC® and Predator PDM® are Registered Trademarks of Predator Software Inc."
    );
    assert_eq!(project.properties.version_info.product_name, "Predator PDM");
}
