use vb6parse::parsers::VB6ModuleFile;

#[test]
fn artificial_life_module_load() {
    let module_file_bytes = include_bytes!("./data/vb6-code/Artificial-life/Declarations.bas");

    let module_file = match VB6ModuleFile::parse("Declarations.bas".to_owned(), module_file_bytes) {
        Ok(module_file) => module_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Declarations.bas' module file");
        }
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn game_physics_basic_module_load() {
    let module_file_bytes = include_bytes!("./data/vb6-code/Game-physics-basic/Physics_Logic.bas");

    let module_file = match VB6ModuleFile::parse("Physics_Logic.bas".to_owned(), module_file_bytes)
    {
        Ok(module_file) => module_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Physics_Logic.bas' module file");
        }
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn histograms_advanced_module_load() {
    let module_file_bytes =
        include_bytes!("./data/vb6-code/Histograms-advanced/mod_PublicVars.bas");

    let module_file = match VB6ModuleFile::parse("mod_PublicVars.bas".to_owned(), module_file_bytes)
    {
        Ok(module_file) => module_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'mod_PublicVars.bas' module file");
        }
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn histograms_basic_module_load() {
    let module_file_bytes = include_bytes!("./data/vb6-code/Histograms-basic/mod_PublicVars.bas");

    let module_file = match VB6ModuleFile::parse("mod_PublicVars.bas".to_owned(), module_file_bytes)
    {
        Ok(module_file) => module_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'mod_PublicVars.bas' module file");
        }
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn levels_effect_module_load() {
    let module_file_bytes = include_bytes!("./data/vb6-code/Levels-effect/mod_PublicVars.bas");

    let module_file = match VB6ModuleFile::parse("mod_PublicVars.bas".to_owned(), module_file_bytes)
    {
        Ok(module_file) => module_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'mod_PublicVars.bas' module file");
        }
    };

    insta::assert_yaml_snapshot!(module_file);
}

#[test]
fn map_editor_2d_module_load() {
    let module1_file_bytes = include_bytes!("./data/vb6-code/Map-editor-2D/Subs.bas");

    let module1_file = match VB6ModuleFile::parse("Subs.bas".to_owned(), module1_file_bytes) {
        Ok(module1_file) => module1_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Subs.bas' module file");
        }
    };

    let module2_file_bytes = include_bytes!("./data/vb6-code/Map-editor-2D/Declarations.bas");

    let module2_file = match VB6ModuleFile::parse("Declarations.bas".to_owned(), module2_file_bytes)
    {
        Ok(module2_file) => module2_file,
        Err(e) => {
            eprintln!("{e}");
            panic!("Failed to parse 'Declarations.bas' module file");
        }
    };

    insta::assert_yaml_snapshot!(module1_file);
    insta::assert_yaml_snapshot!(module2_file);
}
