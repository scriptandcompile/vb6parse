use bstr::{ByteSlice, B};
use vb6parse::project::{ProjectType, VB6Project};

#[test]
fn artificial_life_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Artificial-life/Artificial Life.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 1);
    assert_eq!(project.classes.len(), 2);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmMain".as_bstr()));
    assert_eq!(project.startup, Some(b"frmMain".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"Artificial Life Simulator".as_bstr()));
    assert_eq!(
        project.exe_32_file_name,
        Some(b"Artificial Life.exe".as_bstr())
    );
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"Artificial_Life".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, true);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 2);
    assert_eq!(project.version_info.revision, 76);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.version_info.company_name,
        Some(b"Tanner Helland Productions".as_bstr())
    );
    assert_eq!(project.version_info.file_description, Some(b"".as_bstr()));
    assert_eq!(
        project.version_info.copyright,
        Some(b"\xA92009 Tanner Helland - www.tannerhelland.com".as_bstr())
    );
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(
        project.version_info.product_name,
        Some(B("Artificial Life Simulator").as_bstr())
    );
}

#[test]
fn blacklight_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Blacklight-effect/Blacklight.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 0);
    assert_eq!(project.classes.len(), 2);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmBlacklight".as_bstr()));
    assert_eq!(project.startup, Some(b"frmBlacklight".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"Blacklight".as_bstr()));
    assert_eq!(project.exe_32_file_name, Some(b"Blacklight.exe".as_bstr()));
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"Blacklight_Effect".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, true);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 22);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.version_info.company_name,
        Some(b"Tanner Helland Productions".as_bstr())
    );
    assert_eq!(project.version_info.file_description, Some(b"".as_bstr()));
    assert_eq!(project.version_info.copyright, Some(b"".as_bstr()));
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(project.version_info.product_name, Some(B("").as_bstr()));
}

#[test]
fn brightness_effect_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 0);
    assert_eq!(project.classes.len(), 0);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmBrightness".as_bstr()));
    assert_eq!(project.startup, Some(b"frmBrightness".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"vbBrightness".as_bstr()));
    assert_eq!(
        project.exe_32_file_name,
        Some(b"vbBrightness.exe".as_bstr())
    );
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"VB_Brightness".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, false);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 18);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(project.version_info.company_name, Some(b"".as_bstr()));
    assert_eq!(
        project.version_info.file_description,
        Some(b"Sample executable".as_bstr())
    );
    assert_eq!(
        project.version_info.copyright,
        Some(b"\xA92020 Tanner Helland - www.tannerhelland.com".as_bstr())
    );
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(
        project.version_info.product_name,
        Some(B("vbBrightness.exe").as_bstr())
    );
}

#[test]
fn color_shift_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Color-shift-effect/ShiftColor.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 0);
    assert_eq!(project.classes.len(), 2);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmColorShift".as_bstr()));
    assert_eq!(project.startup, Some(b"frmColorShift".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"ColorShifting".as_bstr()));
    assert_eq!(project.exe_32_file_name, Some(b"ColorShift.exe".as_bstr()));
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"ColorShift".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, true);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 17);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.version_info.company_name,
        Some(b"Tanner Helland Productions".as_bstr())
    );
    assert_eq!(project.version_info.file_description, Some(b"".as_bstr()));
    assert_eq!(project.version_info.copyright, Some(b"".as_bstr()));
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(project.version_info.product_name, Some(B("").as_bstr()));
}

#[test]
fn colorize_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Colorize-effect/Colorize.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 0);
    assert_eq!(project.classes.len(), 2);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmColorize".as_bstr()));
    assert_eq!(project.startup, Some(b"frmColorize".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"Colorize Application".as_bstr()));
    assert_eq!(project.exe_32_file_name, Some(b"Colorize.exe".as_bstr()));
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"Colorize".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, true);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 6);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.version_info.company_name,
        Some(b"Tanner Helland Productions".as_bstr())
    );
    assert_eq!(project.version_info.file_description, Some(b"".as_bstr()));
    assert_eq!(
        project.version_info.copyright,
        Some(b"Published in 2011 by Tanner Helland".as_bstr())
    );
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(project.version_info.product_name, Some(B("").as_bstr()));
}

#[test]
fn contrast_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Contrast-effect/Contrast.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 0);
    assert_eq!(project.classes.len(), 2);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmContrast".as_bstr()));
    assert_eq!(project.startup, Some(b"frmContrast".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"Contrast".as_bstr()));
    assert_eq!(project.exe_32_file_name, Some(b"Contrast.exe".as_bstr()));
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"Image_Contrast".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, true);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 18);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.version_info.company_name,
        Some(b"www.tannerhelland.com".as_bstr())
    );
    assert_eq!(project.version_info.file_description, Some(b"".as_bstr()));
    assert_eq!(project.version_info.copyright, Some(b"".as_bstr()));
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(
        project.version_info.product_name,
        Some(B("Visual Basic Image Contrast Example (Real-time)").as_bstr())
    );
}

#[test]
fn curves_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Curves-effect/Curves.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 0);
    assert_eq!(project.classes.len(), 2);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmCurves".as_bstr()));
    assert_eq!(project.startup, Some(b"frmCurves".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"Curves".as_bstr()));
    assert_eq!(project.exe_32_file_name, Some(b"Curves.exe".as_bstr()));
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"Curves_Dialog".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, true);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 14);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.version_info.company_name,
        Some(b"Tanner Helland Productions".as_bstr())
    );
    assert_eq!(project.version_info.file_description, Some(b"".as_bstr()));
    assert_eq!(project.version_info.copyright, Some(b"".as_bstr()));
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(project.version_info.product_name, Some(B("").as_bstr()));
}

#[test]
fn custom_image_filters_project_load() {
    let project_file_bytes =
        include_bytes!("./data/vb6-code/Custom-image-filters/CustomFilters.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 0);
    assert_eq!(project.classes.len(), 2);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmCustomFilters".as_bstr()));
    assert_eq!(project.startup, Some(b"frmCustomFilters".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"Custom Filters Application".as_bstr()));
    assert_eq!(
        project.exe_32_file_name,
        Some(b"Custom_Filters.exe".as_bstr())
    );
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"CustomFilters_Dialog".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, true);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 8);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.version_info.company_name,
        Some(b"Tanner Helland Productions".as_bstr())
    );
    assert_eq!(project.version_info.file_description, Some(b"".as_bstr()));
    assert_eq!(
        project.version_info.copyright,
        Some(b"\xA92009 Tanner Helland".as_bstr())
    );
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(project.version_info.product_name, Some(B("").as_bstr()));
}

#[test]
fn diffuse_effect_project_load() {
    let project_file_bytes = include_bytes!("./data/vb6-code/Diffuse-effect/Diffuse.vbp");

    let project = VB6Project::parse(project_file_bytes).unwrap();

    assert_eq!(project.project_type, ProjectType::Exe);
    assert_eq!(project.references.len(), 1);
    assert_eq!(project.objects.len(), 0);
    assert_eq!(project.modules.len(), 0);
    assert_eq!(project.classes.len(), 2);
    assert_eq!(project.designers.len(), 0);
    assert_eq!(project.forms.len(), 1);
    assert_eq!(project.user_controls.len(), 0);
    assert_eq!(project.user_documents.len(), 0);

    assert_eq!(project.res_file_32_path, Some(b"".as_bstr()));
    assert_eq!(project.icon_form, Some(b"frmDiffuse".as_bstr()));
    assert_eq!(project.startup, Some(b"frmDiffuse".as_bstr()));
    assert_eq!(project.help_file_path, Some(b"".as_bstr()));
    assert_eq!(project.title, Some(b"Diffuse Application".as_bstr()));
    assert_eq!(project.exe_32_file_name, Some(b"Diffuse.exe".as_bstr()));
    assert_eq!(project.command_line_arguments, Some(b"".as_bstr()));
    assert_eq!(project.name, Some(b"Diffuse_Dialog".as_bstr()));
    assert_eq!(project.help_context_id, Some(b"0".as_bstr()));
    assert_eq!(project.compatible_mode, false);
    assert_eq!(project.upgrade_activex_controls, true);
    assert_eq!(project.server_support_files, false);
    assert_eq!(project.conditional_compile, Some(b"".as_bstr()));
    assert_eq!(project.auto_refresh, true);
    assert_eq!(project.compilation_type, false);
    assert_eq!(project.optimization_type, false);
    assert_eq!(project.favor_pentium_pro, true);
    assert_eq!(project.code_view_debug_info, false);
    assert_eq!(project.aliasing, false);
    assert_eq!(project.bounds_check, true);
    assert_eq!(project.overflow_check, true);
    assert_eq!(project.floating_point_check, true);
    assert_eq!(project.pentium_fdiv_bug_check, true);
    assert_eq!(project.unrounded_floating_point, true);
    assert_eq!(project.start_mode, false);
    assert_eq!(project.unattended, false);
    assert_eq!(project.retained, false);
    assert_eq!(project.thread_per_object, 0);
    assert_eq!(project.max_number_of_threads, 1);
    assert_eq!(project.debug_startup_option, false);

    // version information.
    assert_eq!(project.version_info.major, 1);
    assert_eq!(project.version_info.minor, 0);
    assert_eq!(project.version_info.revision, 10);
    assert_eq!(project.version_info.auto_increment_revision, 1);
    assert_eq!(
        project.version_info.company_name,
        Some(b"Tanner Helland Productions".as_bstr())
    );
    assert_eq!(project.version_info.file_description, Some(b"".as_bstr()));
    assert_eq!(
        project.version_info.copyright,
        Some(b"Published in 2011 by Tanner Helland".as_bstr())
    );
    assert_eq!(project.version_info.trademark, Some(B("").as_bstr()));
    assert_eq!(project.version_info.product_name, Some(B("").as_bstr()));
}
