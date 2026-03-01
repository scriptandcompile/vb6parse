use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

#[test]
fn binary_metamorphosis_v1_form_load() {
    let form_file_bytes = include_bytes!("../data/Binary-metamorphosis/Binary metamorphosis V1.0/src/Bin_To_VB.frm");

    let source_file = SourceFile::decode_with_replacement("Bin_To_VB.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Bin_To_VB.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Bin_To_VB.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/binary_metamorphosis");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn binary_metamorphosis_v2_form_load() {
    let form_file_bytes = include_bytes!("../data/Binary-metamorphosis/Binary metamorphosis V2.0/src/Bin_To_VB.frm");

    let source_file = SourceFile::decode_with_replacement("Bin_To_VB.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Bin_To_VB.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Bin_To_VB.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/binary_metamorphosis");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn binary_metamorphosis_v3_form_load() {
    let form_file_bytes = include_bytes!("../data/Binary-metamorphosis/Binary metamorphosis V3.0/src/Bin_To_VB.frm");

    let source_file = SourceFile::decode_with_replacement("Bin_To_VB.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'Bin_To_VB.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'Bin_To_VB.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/binary_metamorphosis");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}

#[test]
fn binary_metamorphosis_tini_form_load() {
    let form_file_bytes = include_bytes!("../data/Binary-metamorphosis/tini/tini.frm");

    let source_file = SourceFile::decode_with_replacement("tini.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'tini.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'tini.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/binary_metamorphosis");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}
