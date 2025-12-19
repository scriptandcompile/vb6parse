use image::EncodableLayout;
use vb6parse::parsers::ModuleFile;
use vb6parse::SourceFile;

#[test]
fn artificial_life_module_load() {
    let file_path = "./tests/data/vb6-code/Artificial-life/Declarations.bas";
    let module_file_bytes = std::fs::read(file_path).unwrap();

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let Some(module_file) = result.result else {
        if result.has_failures() && !result.failures.is_empty() {
            for failure in result.failures {
                failure.eprint();
            }
        }

        panic!("Failed to parse '{file_path}' module file");
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn game_physics_basic_module_load() {
    let file_path = "./tests/data/vb6-code/Game-physics-basic/Physics_Logic.bas";

    let module_file_bytes = std::fs::read(file_path).unwrap();

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let Some(module_file) = result.result else {
        if result.has_failures() && !result.failures.is_empty() {
            for failure in result.failures {
                failure.eprint();
            }
        }

        panic!("Failed to parse '{file_path}' module file");
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn histograms_advanced_module_load() {
    let file_path = "./tests/data/vb6-code/Histograms-advanced/mod_PublicVars.bas";
    let module_file_bytes = std::fs::read(file_path).unwrap();

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let Some(module_file) = result.result else {
        if result.has_failures() && !result.failures.is_empty() {
            for failure in result.failures {
                failure.eprint();
            }
        }

        panic!("Failed to parse '{file_path}' module file");
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn histograms_basic_module_load() {
    let file_path = "./tests/data/vb6-code/Histograms-basic/mod_PublicVars.bas";
    let module_file_bytes = std::fs::read(file_path).unwrap();

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let Some(module_file) = result.result else {
        if result.has_failures() && !result.failures.is_empty() {
            for failure in result.failures {
                failure.eprint();
            }
        }

        panic!("Failed to parse '{file_path}' module file");
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn levels_effect_module_load() {
    let file_path = "./tests/data/vb6-code/Levels-effect/mod_PublicVars.bas";
    let module_file_bytes = std::fs::read(file_path).unwrap();

    let module_source_file =
        match SourceFile::decode_with_replacement(file_path, module_file_bytes.as_bytes()) {
            Ok(source_file) => source_file,
            Err(e) => {
                e.print();
                panic!("failed to decode module '{file_path}'.");
            }
        };

    let result = ModuleFile::parse(&module_source_file);

    let Some(module_file) = result.result else {
        if result.has_failures() && !result.failures.is_empty() {
            for failure in result.failures {
                failure.eprint();
            }
        }

        panic!("Failed to parse '{file_path}' module file");
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn map_editor_2d_module_load() {
    let subs_file_path = "./tests/data/vb6-code/Map-editor-2D/Subs.bas";

    let subs_module_file_bytes = std::fs::read(subs_file_path).unwrap();

    let subs_module_source_file = match SourceFile::decode_with_replacement(
        subs_file_path,
        subs_module_file_bytes.as_bytes(),
    ) {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("failed to decode module '{subs_file_path}'.");
        }
    };

    let subs_result = ModuleFile::parse(&subs_module_source_file);

    let Some(subs_module_file) = subs_result.result else {
        if subs_result.has_failures() && !subs_result.failures.is_empty() {
            for failure in subs_result.failures {
                failure.eprint();
            }
        }

        panic!("Failed to parse '{subs_file_path}' module file");
    };

    let declaration_file_path = "./tests/data/vb6-code/Map-editor-2D/Declarations.bas";

    let declaration_module_file_bytes = std::fs::read(declaration_file_path).unwrap();

    let declaration_module_source_file = match SourceFile::decode_with_replacement(
        declaration_file_path,
        declaration_module_file_bytes.as_bytes(),
    ) {
        Ok(source_file) => source_file,
        Err(e) => {
            e.print();
            panic!("failed to decode module '{declaration_file_path}'.");
        }
    };

    let declaration_result = ModuleFile::parse(&declaration_module_source_file);

    let Some(declaration_module_file) = declaration_result.result else {
        if declaration_result.has_failures() && !declaration_result.failures.is_empty() {
            for failure in declaration_result.failures {
                failure.eprint();
            }
        }

        panic!("Failed to parse '{declaration_file_path}' module file");
    };

    insta::assert_yaml_snapshot!(subs_module_file);
    insta::assert_yaml_snapshot!(declaration_module_file);
}
