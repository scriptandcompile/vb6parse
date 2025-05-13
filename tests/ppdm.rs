use bstr::{ByteSlice, B};

use vb6parse::parsers::{
    compilesettings::{
        Aliasing, BoundsCheck, CodeViewDebugInfo, CompilationType, FavorPentiumPro,
        FloatingPointErrorCheck, OptimizationType, OverflowCheck, PentiumFDivBugCheck,
        UnroundedFloatingPoint,
    },
    project::{
        CompatibilityMode, DebugStartupOption, InteractionMode, Retained, ServerSupportFiles,
        StartMode, UpgradeControls,
    },
    CompileTargetType, VB6Project,
};

#[test]
fn ppdm_project_load() {
    let project_file_bytes = include_bytes!("./data/ppdm/ppdm.vbp").as_slice();

    let project = VB6Project::parse("ppdm.vbp", project_file_bytes).unwrap();

    assert_eq!(project.project_type, CompileTargetType::Exe);
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
        Some(b"..\\DBCommon\\PSFC.RES".as_bstr())
    );
    assert_eq!(project.icon_form, Some(b"frmMain".as_bstr()));
    assert_eq!(project.startup, Some(b"Sub Main".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"PPDM".as_bstr()));
    assert_eq!(project.exe_32_file_name, Some(b"PPDM.exe".as_bstr()));
    assert_eq!(
        project.command_line_arguments,
        Some(b"-DisableRememberPassword%20-CHARTING -U -language %22english%7d".as_bstr())
    );
    assert_eq!(project.name, Some(b"PPDM".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(
        project.compatibility_mode,
        CompatibilityMode::NoCompatibility
    );
    assert_eq!(project.upgrade_controls, UpgradeControls::Upgrade);
    assert_eq!(project.server_support_files, ServerSupportFiles::Local);
    assert_eq!(
        project.conditional_compile,
        Some(b"PDMBuild = 1 : PDM_SHORTCUTS = 1 : PMData7Build = 0".as_bstr())
    );
    assert_eq!(
        project.compilation_type,
        CompilationType::NativeCode {
            optimization_type: OptimizationType::FavorFastCode,
            favor_pentium_pro: FavorPentiumPro::False,
            code_view_debug_info: CodeViewDebugInfo::NotCreated,
            aliasing: Aliasing::AssumeAliasing,
            bounds_check: BoundsCheck::CheckBounds,
            overflow_check: OverflowCheck::CheckOverflow,
            floating_point_check: FloatingPointErrorCheck::CheckFloatingPointError,
            pentium_fdiv_bug_check: PentiumFDivBugCheck::CheckPentiumFDivBug,
            unrounded_floating_point: UnroundedFloatingPoint::DoNotAllow,
        }
    );

    assert_eq!(project.start_mode, StartMode::StandAlone);
    assert_eq!(project.unattended, InteractionMode::Interactive);
    assert_eq!(project.retained, Retained::UnloadOnExit);
    assert_eq!(project.thread_per_object, Some(0));
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(
        project.debug_startup_option,
        DebugStartupOption::WaitForComponentCreation
    );

    // version information.
    assert_eq!(project.version_info.major, 11);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 288);
    assert_eq!(project.version_info.auto_increment_revision, 0);
    assert_eq!(
        project.version_info.company_name,
        Some(b"Predator Software Inc.".as_bstr())
    );
    assert_eq!(
        project.version_info.file_description,
        Some(b"Predator PDM ".as_bstr())
    );
    assert_eq!(
        project.version_info.copyright,
        Some(B("Copyright �1994 - 2022 Predator Software Inc.  All Rights Reserved.").as_bstr())
    );
    assert_eq!(
        project.version_info.trademark,
        Some(B("Predator SFC� and Predator PDM� are Registered Trademarks of Predator Software Inc.").as_bstr())
    );
    assert_eq!(
        project.version_info.product_name,
        Some(B("Predator PDM").as_bstr())
    );
}
