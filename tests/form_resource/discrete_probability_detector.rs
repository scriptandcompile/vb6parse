use vb6parse::files::resource::FormResourceFile;

#[test]
fn discrete_probability_detector_dpd_frx() {
    let result = FormResourceFile::from_file("./tests/data/Discrete-Probability-Detector-in-VB6/DPD.frx")
        .expect("Failed to read file");

    assert!(!result.has_failures());
    let resource_file = result.unwrap_or_fail();

    let mut entries: Vec<_> = resource_file.iter_entries().collect();
    entries.sort_by_key(|(offset, _)| *offset);

    assert!(!entries.is_empty());
}
