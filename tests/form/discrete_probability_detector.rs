use vb6parse::files::FormFile;
use vb6parse::io::SourceFile;

#[test]
fn discrete_probability_detector_dpd_form_load() {
    let form_file_bytes = include_bytes!("../data/Discrete-Probability-Detector-in-VB6/DPD.frm");

    let source_file = SourceFile::decode_with_replacement("DPD.frm", form_file_bytes);

    let source_file = match source_file {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("Failed to decode 'DPD.frm' source file.");
        }
    };

    let result = FormFile::parse(&source_file);

    let (form_file_opt, failures) = result.unpack();
    let Some(form_file) = form_file_opt else {
        for failure in &failures {
            failure.eprint();
        }
        panic!("Failed to parse 'DPD.frm' form file");
    };

    let mut settings = insta::Settings::clone_current();
    settings.set_snapshot_path("../../snapshots/tests/form/discrete_probability_detector");
    settings.set_prepend_module_to_snapshot(false);
    let _guard = settings.bind_to_scope();
    insta::assert_yaml_snapshot!(form_file);
}
