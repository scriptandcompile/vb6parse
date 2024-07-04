use vb6parse::{class, project::VB6Project};

#[test]
fn bulk_load_all_projects() {
    let projects = [
        "./tests/data/vb6-code/Artificial-life/Artificial Life.vbp",
        "./tests/data/vb6-code/Blacklight-effect/Blacklight.vbp",
        "./tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp",
        "./tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.vbp",
        "./tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.vbp",
        "./tests/data/vb6-code/Color-shift-effect/ShiftColor.vbp",
        "./tests/data/vb6-code/Colorize-effect/Colorize.vbp",
        "./tests/data/vb6-code/Contrast-effect/Contrast.vbp",
        "./tests/data/vb6-code/Curves-effect/Curves.vbp",
        "./tests/data/vb6-code/Custom-image-filters/CustomFilters.vbp",
        "./tests/data/vb6-code/Diffuse-effect/Diffuse.vbp",
        "./tests/data/vb6-code/Edge-detection/EdgeDetection.vbp",
        "./tests/data/vb6-code/Emboss-engrave-effect/EmbossEngrave.vbp",
        "./tests/data/vb6-code/Fill-image-region/Fill_Region.vbp",
        "./tests/data/vb6-code/Fire-effect/FlameTest.vbp",
        "./tests/data/vb6-code/Game-physics-basic/Physics.vbp",
        "./tests/data/vb6-code/Gradient-2D/Gradient.vbp",
        "./tests/data/vb6-code/Grayscale-effect/Grayscale.vbp",
        "./tests/data/vb6-code/Hidden-Markov-model/HMM.vbp",
        "./tests/data/vb6-code/Histograms-advanced/Advanced Histograms.vbp",
        "./tests/data/vb6-code/Histograms-basic/Basic Histograms.vbp",
        "./tests/data/vb6-code/Levels-effect/Image Levels.vbp",
        "./tests/data/vb6-code/Mandelbrot/Mandelbrot.vbp",
        "./tests/data/vb6-code/Map-editor-2D/Map Editor.vbp",
        "./tests/data/vb6-code/Nature-effects/NatureFilters.vbp",
        "./tests/data/vb6-code/Randomize-effects/RandomizationFX.vbp",
        "./tests/data/vb6-code/Scanner-TWAIN/VB_Scanner_Support.vbp",
        "./tests/data/vb6-code/Screen-capture/ScreenCapture.vbp",
        "./tests/data/vb6-code/Sepia-effect/Sepia.vbp",
        "./tests/data/vb6-code/Threshold-effect/Threshold.vbp",
        "./tests/data/vb6-code/Transparency-2D/Transparency.vbp"
    ];

    println!("Loading projects...");

    for project_path in projects.iter() {
        println!("Loading project: {}", project_path);

        let project_contents = std::fs::read(project_path).unwrap();
        let project = VB6Project::parse(&project_contents).unwrap();

        for class_reference in project.classes {
            //remove filename from path
            let project_directory = std::path::Path::new(project_path).parent().unwrap();

            let class_path = project_directory.join(&class_reference.path.to_string());

            if std::fs::metadata(&class_path).is_err() {
                println!("Class not found: {}", class_path.to_str().unwrap());
                continue;
            }

            println!("Loading class: {}", class_path.to_str().unwrap());

            let class_contents = std::fs::read(&class_path).unwrap();
            let _class = class::VB6ClassFile::parse(&class_contents).unwrap();
        }
    }
}
