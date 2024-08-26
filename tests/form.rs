use vb6parse::language::{
    Appearance, DrawMode, DrawStyle, FillStyle, FormBorderStyle, FormLinkMode, MousePointer,
    OLEDropMode, PaletteMode, ScaleMode, StartUpPosition, VB6ControlKind, WindowState,
};
use vb6parse::parsers::VB6FormFile;
use vb6parse::VB6Color;

#[test]
fn artificial_life_form_load() {
    let form_file_bytes = include_bytes!("./data/vb6-code/Artificial-life/frmMain.frm");

    let form_file = VB6FormFile::parse("frmMain.frm".to_owned(), form_file_bytes).unwrap();

    assert_eq!(form_file.format_version.major, 5);
    assert_eq!(form_file.format_version.minor, 0);

    assert_eq!(form_file.form.name, "frmMain");
    assert_eq!(form_file.form.tag, "");
    assert_eq!(form_file.form.index, 0);

    assert_eq!(
        matches!(form_file.form.kind, VB6ControlKind::Form { .. }),
        true
    );

    match &form_file.form.kind {
        VB6ControlKind::Form { properties, .. } => {
            assert_eq!(properties.auto_redraw, false);
            assert_eq!(properties.appearance, Appearance::ThreeD);
            assert_eq!(properties.border_style, FormBorderStyle::Sizable);
            assert_eq!(
                properties.caption,
                "Artificial Life Simulator - www.tannerhelland.com"
            );
            assert_eq!(properties.control_box, true);
            assert_eq!(properties.back_color, VB6Color::System { index: 5 });
            assert_eq!(properties.clip_controls, true);
            assert_eq!(properties.draw_mode, DrawMode::CopyPen);
            assert_eq!(properties.draw_style, DrawStyle::Solid);
            assert_eq!(properties.draw_width, 1);
            assert_eq!(properties.enabled, true);
            assert_eq!(properties.height, 240);
            assert_eq!(properties.help_context_id, 0);
            assert_eq!(
                properties.fill_color,
                VB6Color::RGB {
                    red: 0,
                    green: 0,
                    blue: 0
                }
            );
            assert_eq!(properties.fill_style, FillStyle::Transparent);
            assert_eq!(properties.font_transparent, true);
            assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
            assert_eq!(properties.has_dc, true);
            assert_eq!(properties.icon, None);
            assert_eq!(properties.key_preview, false);
            assert_eq!(properties.left, 0);
            assert_eq!(properties.link_mode, FormLinkMode::None);
            assert_eq!(properties.link_topic, "Form1");
            assert_eq!(properties.max_button, true);
            assert_eq!(properties.mdi_child, false);
            assert_eq!(properties.min_button, true);
            assert_eq!(properties.mouse_icon, None);
            assert_eq!(properties.mouse_pointer, MousePointer::Default);
            assert_eq!(properties.moveable, true);
            assert_eq!(properties.negotiate_menus, true);
            assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
            assert_eq!(properties.palette, None);
            assert_eq!(properties.palette_mode, PaletteMode::HalfTone);
            assert_eq!(properties.picture, None);
            assert_eq!(properties.right_to_left, false);
            assert_eq!(properties.scale_height, 240);
            assert_eq!(properties.scale_left, 0);
            assert_eq!(properties.scale_mode, ScaleMode::Twip);
            assert_eq!(properties.scale_top, 0);
            assert_eq!(properties.scale_width, 240);
            assert_eq!(properties.show_in_taskbar, true);
            assert_eq!(
                properties.start_up_position,
                StartUpPosition::WindowsDefault
            );
            assert_eq!(properties.top, 0);
            assert_eq!(properties.visible, true);
            assert_eq!(properties.whats_this_button, false);
            assert_eq!(properties.whats_this_help, false);
            assert_eq!(properties.width, 240);
            assert_eq!(properties.window_state, WindowState::Normal);
        }
        _ => panic!("Expected Form"),
    };
}
