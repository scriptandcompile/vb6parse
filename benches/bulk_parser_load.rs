use vb6parse::parsers::VB6ClassFile;
use vb6parse::parsers::VB6ModuleFile;
use vb6parse::parsers::VB6Project;
use vb6parse::parsers::VB6Stream;

use criterion::{black_box, criterion_group, criterion_main, Criterion};

fn project_benchmark(c: &mut Criterion) {
    let project_names = vec![
        "Artificial Life.vbp".to_owned(),
        "Blacklight.vbp".to_owned(),
        "Brightness.vbp".to_owned(),
        "Brightness2.vbp".to_owned(),
        "Brightness3.vbp".to_owned(),
        "ShiftColor.vbp".to_owned(),
        "Colorize.vbp".to_owned(),
        "Contrast.vbp".to_owned(),
        "Curves.vbp".to_owned(),
        "CustomFilters.vbp".to_owned(),
        "Diffuse.vbp".to_owned(),
        "EdgeDetection.vbp".to_owned(),
        "EmbossEngrave.vbp".to_owned(),
        "Fill_Region.vbp".to_owned(),
        "FlameTest.vbp".to_owned(),
        "Physics.vbp".to_owned(),
        "Gradient.vbp".to_owned(),
        "Grayscale.vbp".to_owned(),
        "HMM.vbp".to_owned(),
        "Advanced Histograms.vbp".to_owned(),
        "Basic Histograms.vbp".to_owned(),
        "Image Levels.vbp".to_owned(),
        "Mandelbrot.vbp".to_owned(),
        "Map Editor.vbp".to_owned(),
        "NatureFilters.vbp".to_owned(),
        "RandomizationFX.vbp".to_owned(),
        "VB_Scanner_Support.vbp".to_owned(),
        "ScreenCapture.vbp".to_owned(),
        "Sepia.vbp".to_owned(),
        "Threshold.vbp".to_owned(),
        "Transparency.vbp".to_owned(),
    ];

    let projects = vec![
        include_bytes!("../tests/data/vb6-code/Artificial-life/Artificial Life.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Blacklight-effect/Blacklight.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.vbp")
            .to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Color-shift-effect/ShiftColor.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Colorize-effect/Colorize.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Contrast-effect/Contrast.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Curves-effect/Curves.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Custom-image-filters/CustomFilters.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Diffuse-effect/Diffuse.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Edge-detection/EdgeDetection.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Emboss-engrave-effect/EmbossEngrave.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Fill-image-region/Fill_Region.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Fire-effect/FlameTest.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Game-physics-basic/Physics.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Gradient-2D/Gradient.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Grayscale-effect/Grayscale.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Hidden-Markov-model/HMM.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-advanced/Advanced Histograms.vbp")
            .to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-basic/Basic Histograms.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Levels-effect/Image Levels.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Mandelbrot/Mandelbrot.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Map-editor-2D/Map Editor.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Nature-effects/NatureFilters.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Randomize-effects/RandomizationFX.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Scanner-TWAIN/VB_Scanner_Support.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Screen-capture/ScreenCapture.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Sepia-effect/Sepia.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Threshold-effect/Threshold.vbp").to_vec(),
        include_bytes!("../tests/data/vb6-code/Transparency-2D/Transparency.vbp").to_vec(),
    ];

    let project_pairs: Vec<(_, _)> = project_names.iter().zip(projects.iter()).collect();

    c.bench_function("load multiple projects", |b| {
        b.iter(|| {
            for project_pair in &project_pairs {
                black_box({
                    let mut stream = VB6Stream::new(project_pair.0, project_pair.1.as_slice());
                    let _proj = VB6Project::parse(&mut stream);
                });
            }
        })
    });
}

fn class_benchmark(c: &mut Criterion) {
    let class_names = vec![
        "FastDrawing.cls".to_owned(),
        "pdOpenSaveDialog.cls".to_owned(),
        "cCommonDialog.cls".to_owned(),
        "cSystemColorDialog.cls".to_owned(),
    ];

    let classes = vec![
        include_bytes!("../tests/data/vb6-code/Levels-effect/FastDrawing.cls").to_vec(),
        include_bytes!("../tests/data/vb6-code/Levels-effect/pdOpenSaveDialog.cls").to_vec(),
        include_bytes!("../tests/data/vb6-code/Randomize-effects/cCommonDialog.cls").to_vec(),
        include_bytes!("../tests/data/vb6-code/Emboss-engrave-effect/cSystemColorDialog.cls")
            .to_vec(),
        include_bytes!("../tests/data/vb6-code/Artificial-life/Organism.cls").to_vec(),
    ];

    let class_pairs: Vec<(_, _)> = class_names.iter().zip(classes.iter()).collect();

    c.bench_function("load multiple classes", |b| {
        b.iter(|| {
            for class_pair in &class_pairs {
                black_box({
                    let _class =
                        VB6ClassFile::parse(class_pair.0.to_string(), &mut class_pair.1.as_slice());
                });
            }
        })
    });
}

fn bas_module_benchmark(c: &mut Criterion) {
    let bas_module_names = vec![
        "Physics_Logic.bas".to_owned(),
        "mod_PublicVars.bas".to_owned(),
        "mod_PublicVars.bas".to_owned(),
        "Declarations.bas".to_owned(),
        "Subs.bas".to_owned(),
        "mod_PublicVars.bas".to_owned(),
        "Declarations.bas".to_owned(),
    ];

    let bas_modules = vec![
        include_bytes!("../tests/data/vb6-code/Game-physics-basic/Physics_Logic.bas").to_vec(),
        include_bytes!("../tests/data/vb6-code/Levels-effect/mod_PublicVars.bas").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-advanced/mod_PublicVars.bas").to_vec(),
        include_bytes!("../tests/data/vb6-code/Map-editor-2D/Declarations.bas").to_vec(),
        include_bytes!("../tests/data/vb6-code/Map-editor-2D/Subs.bas").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-basic/mod_PublicVars.bas").to_vec(),
        include_bytes!("../tests/data/vb6-code/Artificial-life/Declarations.bas").to_vec(),
    ];

    let bas_modules_pairs: Vec<(_, _)> = bas_module_names.iter().zip(bas_modules.iter()).collect();

    c.bench_function("load multiple bas modules", |b| {
        b.iter(|| {
            for bas_module_pair in &bas_modules_pairs {
                black_box({
                    let _class = VB6ModuleFile::parse(
                        bas_module_pair.0.to_string(),
                        &mut bas_module_pair.1.as_slice(),
                    );
                });
            }
        })
    });
}

fn form_benchmark(c: &mut Criterion) {
    let form_names = vec![
        "FormPhysics.frm".to_owned(),
        "Histogram.frm".to_owned(),
        "Main.frm".to_owned(),
        "Contrast.frm".to_owned(),
        "Colorize.frm".to_owned(),
        "CustomFilters.frm".to_owned(),
        //"Diffuse.frm".to_owned(),
        "frmHMM.frm".to_owned(),
        "Grayscale.frm".to_owned(),
        "frmScanner.frm".to_owned(),
        "ShiftColors.frm".to_owned(),
        "frmFill.frm".to_owned(),
        //"Threshold.frm".to_owned(),
        "frmTransparency.frm".to_owned(),
        "EdgeDetection.frm".to_owned(),
        "NatureFilters.frm".to_owned(),
        "FormScreenCapture.frm".to_owned(),
        "Histogram.frm".to_owned(),
        "Main.frm".to_owned(),
        "Sepia.frm".to_owned(),
        "frmFire.frm".to_owned(),
        "RandomizationFX.frm".to_owned(),
        "Brightness3.frm".to_owned(),
        "Brightness.frm".to_owned(),
        "Brightness2.frm".to_owned(),
        "Brightness.frm".to_owned(),
        "FrmResize.frm".to_owned(),
        "Main Editor.frm".to_owned(),
        "Histogram.frm".to_owned(),
        "Main.frm".to_owned(),
        "EmbossEngrave.frm".to_owned(),
        "Blacklight.frm".to_owned(),
        "Gradient.frm".to_owned(),
        "Mandelbrot.frm".to_owned(),
        "Curves.frm".to_owned(),
        "frmMain.frm".to_owned(),
    ];

    let forms = vec![
        include_bytes!("../tests/data/vb6-code/Game-physics-basic/FormPhysics.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Levels-effect/Histogram.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Levels-effect/Main.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Contrast-effect/Contrast.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Colorize-effect/Colorize.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Custom-image-filters/CustomFilters.frm").to_vec(),
        //include_bytes!("../tests/data/vb6-code/Diffuse-effect/Diffuse.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Hidden-Markov-model/frmHMM.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Grayscale-effect/Grayscale.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Scanner-TWAIN/frmScanner.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Color-shift-effect/ShiftColors.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Fill-image-region/frmFill.frm").to_vec(),
        //include_bytes!("../tests/data/vb6-code/Threshold-effect/Threshold.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Transparency-2D/frmTransparency.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Edge-detection/EdgeDetection.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Nature-effects/NatureFilters.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Screen-capture/FormScreenCapture.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-advanced/Histogram.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-advanced/Main.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Sepia-effect/Sepia.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Fire-effect/frmFire.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Randomize-effects/RandomizationFX.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 3 - DIBs/Brightness3.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 4 - Even faster DIBs/Brightness.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 2 - API - GetPixel and SetPixel/Brightness2.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Brightness-effect/Part 1 - Pure VB6/Brightness.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Map-editor-2D/FrmResize.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Map-editor-2D/Main Editor.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-basic/Histogram.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Histograms-basic/Main.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Emboss-engrave-effect/EmbossEngrave.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Blacklight-effect/Blacklight.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Gradient-2D/Gradient.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Mandelbrot/Mandelbrot.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Curves-effect/Curves.frm").to_vec(),
        include_bytes!("../tests/data/vb6-code/Artificial-life/frmMain.frm").to_vec(),
        
    ];

    let forms_pairs: Vec<(_, _)> = form_names.iter().zip(forms.iter()).collect();

    c.bench_function("load multiple forms", |b| {
        b.iter(|| {
            for form_pair in &forms_pairs {
                black_box({
                    let _class = VB6ModuleFile::parse(
                        form_pair.0.to_string(),
                        &mut form_pair.1.as_slice(),
                    );
                });
            }
        })
    });
}

criterion_group!(
    benches,
    project_benchmark,
    class_benchmark,
    bas_module_benchmark,
    form_benchmark
);
criterion_main!(benches);
