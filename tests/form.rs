use vb6parse::language::{
    Align, Alignment, Appearance, BackStyle, BorderStyle, CheckBoxValue, ComboBoxStyle, DragMode,
    DrawMode, DrawStyle, FillStyle, FormBorderStyle, FormLinkMode, JustifyAlignment, LinkMode,
    MousePointer, OLEDragMode, OLEDropMode, PaletteMode, ScaleMode, ScrollBars, StartUpPosition,
    Style, VB6ControlKind, WindowState,
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

    let VB6ControlKind::Form {
        properties: form_properties,
        controls: form_controls,
        menus: form_menus,
    } = &form_file.form.kind
    else {
        panic!("Expected Form");
    };
    assert_eq!(form_controls.len(), 10);
    assert_eq!(form_menus.len(), 0);

    assert_eq!(form_properties.auto_redraw, false);
    assert_eq!(form_properties.appearance, Appearance::ThreeD);
    assert_eq!(form_properties.border_style, FormBorderStyle::Sizable);
    assert_eq!(
        form_properties.caption,
        "Artificial Life Simulator - www.tannerhelland.com"
    );
    assert_eq!(form_properties.control_box, true);
    assert_eq!(form_properties.back_color, VB6Color::System { index: 5 });
    assert_eq!(form_properties.clip_controls, true);
    assert_eq!(form_properties.draw_mode, DrawMode::CopyPen);
    assert_eq!(form_properties.draw_style, DrawStyle::Solid);
    assert_eq!(form_properties.draw_width, 1);
    assert_eq!(form_properties.enabled, true);
    assert_eq!(form_properties.height, 240);
    assert_eq!(form_properties.help_context_id, 0);
    assert_eq!(
        form_properties.fill_color,
        VB6Color::RGB {
            red: 0,
            green: 0,
            blue: 0
        }
    );
    assert_eq!(form_properties.fill_style, FillStyle::Transparent);
    assert_eq!(form_properties.font_transparent, true);
    assert_eq!(form_properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(form_properties.has_dc, true);
    assert_eq!(form_properties.icon, None);
    assert_eq!(form_properties.key_preview, false);
    assert_eq!(form_properties.left, 0);
    assert_eq!(form_properties.link_mode, FormLinkMode::None);
    assert_eq!(form_properties.link_topic, "Form1");
    assert_eq!(form_properties.max_button, true);
    assert_eq!(form_properties.mdi_child, false);
    assert_eq!(form_properties.min_button, true);
    assert_eq!(form_properties.mouse_icon, None);
    assert_eq!(form_properties.mouse_pointer, MousePointer::Default);
    assert_eq!(form_properties.moveable, true);
    assert_eq!(form_properties.negotiate_menus, true);
    assert_eq!(form_properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(form_properties.palette, None);
    assert_eq!(form_properties.palette_mode, PaletteMode::HalfTone);
    assert_eq!(form_properties.picture, None);
    assert_eq!(form_properties.right_to_left, false);
    assert_eq!(form_properties.scale_height, 240);
    assert_eq!(form_properties.scale_left, 0);
    assert_eq!(form_properties.scale_mode, ScaleMode::Twip);
    assert_eq!(form_properties.scale_top, 0);
    assert_eq!(form_properties.scale_width, 240);
    assert_eq!(form_properties.show_in_taskbar, true);
    assert_eq!(
        form_properties.start_up_position,
        StartUpPosition::WindowsDefault
    );
    assert_eq!(form_properties.top, 0);
    assert_eq!(form_properties.visible, true);
    assert_eq!(form_properties.whats_this_button, false);
    assert_eq!(form_properties.whats_this_help, false);
    assert_eq!(form_properties.width, 240);
    assert_eq!(form_properties.window_state, WindowState::Normal);

    //
    // Check the form's, nested controls - picturebox, index 0.
    //

    let form_index = 0;
    assert_eq!(form_controls[form_index].name, "picFront");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::PictureBox { properties } = &form_controls[form_index].kind else {
        panic!("Expected PictureBox");
    };

    assert_eq!(properties.align, Align::None);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_redraw, false);
    assert_eq!(properties.auto_size, false);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H8000000F&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.clip_controls, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.draw_mode, DrawMode::CopyPen);
    assert_eq!(properties.draw_style, DrawStyle::Solid);
    assert_eq!(properties.draw_width, 1);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fill_color,
        VB6Color::from_hex("&H00000000&").unwrap()
    );
    assert_eq!(properties.fill_style, FillStyle::Solid);
    assert_eq!(properties.font_transparent, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000012&").unwrap()
    );
    assert_eq!(properties.has_dc, true);
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.picture, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scale_height, 100);
    assert_eq!(properties.scale_left, 0);
    assert_eq!(properties.scale_mode, ScaleMode::Twip);
    assert_eq!(properties.scale_top, 0);
    assert_eq!(properties.scale_width, 100);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, nested controls - CommandButton index 2.
    //

    let form_index = 1;
    assert_eq!(form_controls[form_index].name, "cmdSaveData");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::CommandButton { properties } = &form_controls[form_index].kind else {
        panic!("Expected CommandButton");
    };

    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H8000000F&").unwrap()
    );
    assert_eq!(properties.cancel, false);
    assert_eq!(properties.caption, "Command1");
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.default, false);
    assert_eq!(properties.disabled_picture, None);
    assert_eq!(properties.down_picture, None);
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(
        properties.mask_color,
        VB6Color::from_hex("&H00C0C0C0&").unwrap()
    );
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.picture, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.style, Style::Standard);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mask_color, false);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, nested controls - Checkbox, index 2.
    //

    let form_index = 2;
    assert_eq!(form_controls[form_index].name, "chkDisplayDead");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::CheckBox { properties } = &form_controls[form_index].kind else {
        panic!("Expected CheckBox");
    };

    assert_eq!(properties.alignment, JustifyAlignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H8000000F&").unwrap()
    );
    assert_eq!(properties.caption, "Check1");
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.disabled_picture, None);
    assert_eq!(properties.down_picture, None);
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000012&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(
        properties.mask_color,
        VB6Color::from_hex("&H00C0C0C0&").unwrap()
    );
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.picture, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.style, Style::Standard);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mask_color, false);
    assert_eq!(properties.value, CheckBoxValue::Unchecked);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, nested controls - Frame, index 3.
    //

    // Check the form's, frame control
    let form_index = 3;
    assert_eq!(form_controls[form_index].name, "frmStartSettings");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::Frame {
        properties,
        controls: frame_controls,
    } = &form_controls[form_index].kind
    else {
        panic!("Expected Frame");
    };

    assert_eq!(frame_controls.len(), 19);

    assert_eq!(properties.appearance, Appearance::Flat);
    assert_eq!(properties.back_color, VB6Color::System { index: 5 });
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.caption, "Initial Simulation Settings:");
    assert_eq!(properties.clip_controls, true);
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 8 });
    assert_eq!(properties.height, 4095);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 7800);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 7);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 1080);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 4215);

    //
    // Check the form's, frame's, nested controls - TextBox, index 0.
    //

    // Check the form's, frame's, nested controls
    let frame_index = 0;
    assert_eq!(frame_controls[frame_index].name, "txtMutations");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H80000005&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000008&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - TextBox, index 1.
    //

    let frame_index = 1;
    assert_eq!(frame_controls[frame_index].name, "txtMutateTurns");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H80000005&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000008&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - CheckBox, index 2.
    //

    let frame_index = 2;
    assert_eq!(frame_controls[frame_index].name, "chkMultiply");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::CheckBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected CheckBox");
    };

    assert_eq!(properties.alignment, JustifyAlignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H8000000F&").unwrap()
    );
    assert_eq!(properties.caption, "Check1");
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.disabled_picture, None);
    assert_eq!(properties.down_picture, None);
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000012&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(
        properties.mask_color,
        VB6Color::from_hex("&H00C0C0C0&").unwrap()
    );
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.picture, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.style, Style::Standard);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mask_color, false);
    assert_eq!(properties.value, CheckBoxValue::Unchecked);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - TextBox, index 3.
    //

    let frame_index = 3;
    assert_eq!(frame_controls[frame_index].name, "txtFoodWorth");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H80000005&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000008&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - TextBox, index 4.
    //

    let frame_index = 4;
    assert_eq!(frame_controls[frame_index].name, "txtFoodGen");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H80000005&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000008&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - TextBox, index 5.
    //

    let frame_index = 5;
    assert_eq!(frame_controls[frame_index].name, "txtFoodRegen");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H80000005&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000008&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - TextBox, index 6.
    //

    let frame_index = 6;
    assert_eq!(frame_controls[frame_index].name, "txtInitialEnergy");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H80000005&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000008&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - TextBox, index 7.
    //

    let frame_index = 7;
    assert_eq!(frame_controls[frame_index].name, "txtInitialFood");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H80000005&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000008&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - TextBox, index 8.
    //

    let frame_index = 8;
    assert_eq!(frame_controls[frame_index].name, "txtOrganisms");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H80000005&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000008&").unwrap()
    );
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's, nested controls - Label, index 9.
    //

    let frame_index = 9;
    assert_eq!(frame_controls[frame_index].name, "Label8");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Label { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);

    //
    // Check the form's, frame's, nested controls - Label, index 10.
    //

    let frame_index = 10;
    assert_eq!(frame_controls[frame_index].name, "Label7");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Label { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);

    //
    // Check the form's, frame's, nested controls - Line, index 11.
    //

    let frame_index = 11;
    assert_eq!(frame_controls[frame_index].name, "Line2");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Line { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Line");
    };

    assert_eq!(properties.border_color, VB6Color::System { index: 8 });
    assert_eq!(properties.border_style, DrawStyle::Solid);
    assert_eq!(properties.border_width, 1);
    assert_eq!(properties.draw_mode, DrawMode::CopyPen);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.x1, 0);
    assert_eq!(properties.y1, 0);
    assert_eq!(properties.x2, 100);
    assert_eq!(properties.y2, 100);

    //
    // Check the form's, frame's, nested controls - Label, index 12.
    //

    let frame_index = 12;
    assert_eq!(frame_controls[frame_index].name, "Label6");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Label { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);

    //
    // Check the form's, frame's, nested controls - Label, index 13.
    //

    let frame_index = 13;
    assert_eq!(frame_controls[frame_index].name, "Label5");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Label { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);

    //
    // Check the form's, frame's, nested controls - Label, index 14.
    //

    let frame_index = 14;
    assert_eq!(frame_controls[frame_index].name, "Label4");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Label { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);

    //
    // Check the form's, frame's, nested controls - Label, index 15.
    //

    let frame_index = 15;
    assert_eq!(frame_controls[frame_index].name, "Label3");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Label { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);

    //
    // Check the form's, frame's, nested controls - Label, index 16.
    //

    let frame_index = 16;
    assert_eq!(frame_controls[frame_index].name, "Label2");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Label { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);

    //
    // Check the form's, frame's, nested controls - Line, index 17.
    //

    let frame_index = 17;
    assert_eq!(frame_controls[frame_index].name, "Line1");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Line { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Line");
    };

    assert_eq!(properties.border_color, VB6Color::System { index: 8 });
    assert_eq!(properties.border_style, DrawStyle::Solid);
    assert_eq!(properties.border_width, 1);
    assert_eq!(properties.draw_mode, DrawMode::CopyPen);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.x1, 0);
    assert_eq!(properties.y1, 0);
    assert_eq!(properties.x2, 100);
    assert_eq!(properties.y2, 100);

    //
    // Check the form's, frame's, nested controls - Label, index 18.
    //

    let frame_index = 18;
    assert_eq!(frame_controls[frame_index].name, "Label1");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::Label { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);

    // back to processing the form's controls

    //
    // Check the form's, nested controls - Frame, index 4.
    //

    // Check the form's, frame control
    let form_index = 4;
    assert_eq!(form_controls[form_index].name, "frmOrganisms");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::Frame {
        properties,
        controls: frame_controls,
    } = &form_controls[form_index].kind
    else {
        panic!("Expected Frame");
    };

    assert_eq!(frame_controls.len(), 2);

    assert_eq!(properties.appearance, Appearance::Flat);
    assert_eq!(properties.back_color, VB6Color::System { index: 5 });
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(
        properties.caption,
        "Select a creature for detailed information:"
    );
    assert_eq!(properties.clip_controls, true);
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 8 });
    assert_eq!(properties.height, 3135);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 7800);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 4);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 5400);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 4215);

    //
    // Check the form's, frame's[4], nested controls - TextBox, index 0.
    //

    let frame_index = 0;
    assert_eq!(frame_controls[frame_index].name, "txtInfo");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::TextBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected TextBox");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.back_color, VB6Color::System { index: 5 });
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 8 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.hide_selection, true);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.locked, false);
    assert_eq!(properties.max_length, 0);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.multi_line, false);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.password_char, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scroll_bars, ScrollBars::None);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, frame's[4], nested controls - ComboBox, index 1.
    //

    let frame_index = 1;
    assert_eq!(frame_controls[frame_index].name, "cmbOrganisms");
    assert_eq!(frame_controls[frame_index].index, 0);
    assert_eq!(frame_controls[frame_index].tag, "");

    let VB6ControlKind::ComboBox { properties } = &frame_controls[frame_index].kind else {
        panic!("Expected ComboBox");
    };

    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.back_color, VB6Color::System { index: 5 });
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 8 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.integral_height, true);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.locked, false);
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.sorted, false);
    assert_eq!(properties.style, ComboBoxStyle::DropDownCombo);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.text, "");
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, nested controls - CommandButton, index 5.
    //

    // Check the form's, commandbutton control
    let form_index = 5;
    assert_eq!(form_controls[form_index].name, "cmdStop");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::CommandButton { properties } = &form_controls[form_index].kind else {
        panic!("Expected CommandButton");
    };

    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H8000000F&").unwrap()
    );
    assert_eq!(properties.cancel, false);
    assert_eq!(properties.caption, "Command1");
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.default, false);
    assert_eq!(properties.disabled_picture, None);
    assert_eq!(properties.down_picture, None);
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(
        properties.mask_color,
        VB6Color::from_hex("&H00C0C0C0&").unwrap()
    );
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.picture, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.style, Style::Standard);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mask_color, false);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, nested controls - CommandButton, index 6.
    //

    // Check the form's, commandbutton control
    let form_index = 6;
    assert_eq!(form_controls[form_index].name, "cmdStart");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::CommandButton { properties } = &form_controls[form_index].kind else {
        panic!("Expected CommandButton");
    };

    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H8000000F&").unwrap()
    );
    assert_eq!(properties.cancel, false);
    assert_eq!(properties.caption, "Command1");
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.default, false);
    assert_eq!(properties.disabled_picture, None);
    assert_eq!(properties.down_picture, None);
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(
        properties.mask_color,
        VB6Color::from_hex("&H00C0C0C0&").unwrap()
    );
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.picture, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.style, Style::Standard);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mask_color, false);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, nested controls - PictureBox, index 7.
    //

    // Check the form's, picturebox control
    let form_index = 7;
    assert_eq!(form_controls[form_index].name, "picMap");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::PictureBox { properties } = &form_controls[form_index].kind else {
        panic!("Expected PictureBox");
    };

    assert_eq!(properties.align, Align::None);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_redraw, false);
    assert_eq!(properties.auto_size, false);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H8000000F&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.clip_controls, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.draw_mode, DrawMode::CopyPen);
    assert_eq!(properties.draw_style, DrawStyle::Solid);
    assert_eq!(properties.draw_width, 1);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fill_color,
        VB6Color::from_hex("&H00000000&").unwrap()
    );
    assert_eq!(properties.fill_style, FillStyle::Solid);
    assert_eq!(properties.font_transparent, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000012&").unwrap()
    );
    assert_eq!(properties.has_dc, true);
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.picture, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scale_height, 100);
    assert_eq!(properties.scale_left, 0);
    assert_eq!(properties.scale_mode, ScaleMode::Twip);
    assert_eq!(properties.scale_top, 0);
    assert_eq!(properties.scale_width, 100);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, nested controls - PictureBox, index 8.
    //

    // Check the form's, picturebox control
    let form_index = 8;
    assert_eq!(form_controls[form_index].name, "picFood");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::PictureBox { properties } = &form_controls[form_index].kind else {
        panic!("Expected PictureBox");
    };

    assert_eq!(properties.align, Align::None);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_redraw, false);
    assert_eq!(properties.auto_size, false);
    assert_eq!(
        properties.back_color,
        VB6Color::from_hex("&H8000000F&").unwrap()
    );
    assert_eq!(properties.border_style, BorderStyle::FixedSingle);
    assert_eq!(properties.causes_validation, true);
    assert_eq!(properties.clip_controls, true);
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.draw_mode, DrawMode::CopyPen);
    assert_eq!(properties.draw_style, DrawStyle::Solid);
    assert_eq!(properties.draw_width, 1);
    assert_eq!(properties.enabled, true);
    assert_eq!(
        properties.fill_color,
        VB6Color::from_hex("&H00000000&").unwrap()
    );
    assert_eq!(properties.fill_style, FillStyle::Solid);
    assert_eq!(properties.font_transparent, true);
    assert_eq!(
        properties.fore_color,
        VB6Color::from_hex("&H80000012&").unwrap()
    );
    assert_eq!(properties.has_dc, true);
    assert_eq!(properties.height, 30);
    assert_eq!(properties.help_context_id, 0);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drag_mode, OLEDragMode::Manual);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.picture, None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.scale_height, 100);
    assert_eq!(properties.scale_left, 0);
    assert_eq!(properties.scale_mode, ScaleMode::Twip);
    assert_eq!(properties.scale_top, 0);
    assert_eq!(properties.scale_width, 100);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tab_stop, true);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);

    //
    // Check the form's, nested controls - Label, index 9.
    //

    // Check the form's, Label control
    let form_index = 9;
    assert_eq!(form_controls[form_index].name, "lblTitle");
    assert_eq!(form_controls[form_index].index, 0);
    assert_eq!(form_controls[form_index].tag, "");

    let VB6ControlKind::Label { properties } = &form_controls[form_index].kind else {
        panic!("Expected Label");
    };

    assert_eq!(properties.alignment, Alignment::LeftJustify);
    assert_eq!(properties.appearance, Appearance::ThreeD);
    assert_eq!(properties.auto_size, false);
    assert_eq!(properties.back_color, VB6Color::System { index: 15 });
    assert_eq!(properties.back_style, BackStyle::Opaque);
    assert_eq!(properties.border_style, BorderStyle::None);
    assert_eq!(properties.caption, "Label1");
    assert_eq!(properties.data_field, "");
    assert_eq!(properties.data_format, "");
    assert_eq!(properties.data_member, "");
    assert_eq!(properties.data_source, "");
    assert_eq!(properties.drag_icon, None);
    assert_eq!(properties.drag_mode, DragMode::Manual);
    assert_eq!(properties.enabled, true);
    assert_eq!(properties.fore_color, VB6Color::System { index: 18 });
    assert_eq!(properties.height, 30);
    assert_eq!(properties.left, 30);
    assert_eq!(properties.link_item, "");
    assert_eq!(properties.link_mode, LinkMode::None);
    assert_eq!(properties.link_timeout, 50);
    assert_eq!(properties.link_topic, "");
    assert_eq!(properties.mouse_icon, None);
    assert_eq!(properties.mouse_pointer, MousePointer::Default);
    assert_eq!(properties.ole_drop_mode, OLEDropMode::None);
    assert_eq!(properties.right_to_left, false);
    assert_eq!(properties.tab_index, 0);
    assert_eq!(properties.tool_tip_text, "");
    assert_eq!(properties.top, 30);
    assert_eq!(properties.use_mnemonic, true);
    assert_eq!(properties.visible, true);
    assert_eq!(properties.whats_this_help_id, 0);
    assert_eq!(properties.width, 100);
    assert_eq!(properties.word_wrap, false);
}
