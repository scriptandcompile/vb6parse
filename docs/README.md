# VB6Parse GitHub Pages

This directory contains the GitHub Pages website for VB6Parse.

## Viewing Locally

To view the site locally, simply open `index.html` in a web browser.

## Deployment

To deploy this site to GitHub Pages:

1. Push the `docs` directory to the repository
2. Go to repository Settings → Pages
3. Set Source to "Deploy from a branch"
4. Select branch: `main` and folder: `/docs`
5. Save

The site will be available at: `https://scriptandcompile.github.io/vb6parse`

## Structure

```
docs/
├── index.html - Homepage
├── getting-started.html - Getting started guide (loads examples from GitHub CDN)
├── documentation.html - API documentation
├── benchmarks.html - Benchmark results page
├── coverage.html - Test coverage page
├── library/ - VB6 Library Reference (generated)
│   ├── index.html - Library index
│   ├── search-index.json - Search data
│   ├── functions/ - Function documentation (generated)
│   └── statements/ - Statement documentation (generated)
├── assets/
│   ├── css/ - Stylesheets
│   │   ├── style.css - Main stylesheet
│   │   ├── docs-style.css - Documentation styles
│   │   ├── benchmarks.css - Benchmark page styles
│   │   └── coverage.css - Coverage page styles
│   ├── js/ - JavaScript files
│   │   ├── theme-switcher.js - Dark/light theme toggle
│   │   ├── benchmarks.js - Benchmark data loading
│   │   └── stats.js - Statistics loading
│   └── data/ - Generated data files
│       ├── benchmarks.json - Benchmark results (generated)
│       ├── coverage.json - Coverage data (generated)
│       └── stats.json - Project statistics (generated)
├── technical/ - Technical documentation
├── _config.yml - Jekyll configuration
└── README.md - This file
```

## Getting Started Examples

The `getting-started.html` page dynamically loads code examples from `examples/docs/` via GitHub CDN. This ensures:
- Documentation always shows current, working code
- Examples are verified by pre-commit hooks (`cargo check --examples`)
- No duplication between docs and repository

## Updating

When updating the website:

1. Edit HTML files for content changes
2. Edit CSS files in `assets/css/` for styling changes  
3. Edit JS files in `assets/js/` for functionality changes
4. Test locally by opening `index.html` in a browser
5. **Getting Started examples:** Edit files in `examples/docs/` (automatically loaded via CDN)
6. Regenerate data files:
   - Run `python3 scripts/generate-library-docs.py` for library reference
   - Run `python3 scripts/generate-benchmarks.py` for benchmark data
   - Run `python3 scripts/generate-coverage.py` for coverage data
7. Commit and push changes
8. GitHub Pages will automatically rebuild

## Future Enhancements

Potential additions:
- More interactive examples
- WebAssembly-based live parser demo
