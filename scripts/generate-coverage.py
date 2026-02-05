#!/usr/bin/env python3
"""
Generate coverage data and test statistics for VB6Parse.
Cross-platform script for Windows and Linux.
"""

import json
import os
import sys
import subprocess
import glob
from pathlib import Path


def run_coverage():
    """Run cargo llvm-cov to generate coverage data."""
    print("Generating coverage data...")
    
    try:
        subprocess.run(
            ["cargo", "llvm-cov", "--lib", "--tests", "--json", "--output-path", "docs/assets/data/coverage.json"],
            check=True
        )
    except subprocess.CalledProcessError as e:
        print(f"Error running coverage: {e}", file=sys.stderr)
        sys.exit(1)


def generate_html_coverage():
    """Generate HTML coverage reports using llvm-cov."""
    print("Generating HTML coverage reports...")
    
    output_dir = Path('docs/assets/coverage')
    output_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        subprocess.run(
            [
                "cargo", 
                "llvm-cov", 
                "--lib",
                "--tests",
                "--html",
                "--output-dir", 
                str(output_dir)
            ],
            check=True
        )
        
        # Restructure the output to use workspace-relative paths
        restructure_coverage_html(output_dir)
        
        print(f"✓ HTML coverage reports generated in {output_dir}")
        return output_dir
    except subprocess.CalledProcessError as e:
        print(f"Error generating HTML coverage: {e}", file=sys.stderr)
        sys.exit(1)


def restructure_coverage_html(output_dir):
    """Restructure llvm-cov HTML output to use workspace-relative paths."""
    import shutil
    import re
    
    html_dir = output_dir / 'html'
    
    if not html_dir.exists():
        print("Warning: html directory not found, skipping restructure")
        return
    
    # Move style.css and control.js to the coverage root (index.html not needed)
    for file in ['style.css', 'control.js']:
        src = html_dir / file
        if src.exists():
            shutil.move(str(src), str(output_dir / file))
    
    # Find the workspace root in the nested structure
    coverage_subdir = html_dir / 'coverage'
    if not coverage_subdir.exists():
        print("Warning: coverage subdirectory not found")
        return
    
    # Get the current working directory to find where the project path starts
    project_root = Path.cwd()
    
    # Navigate through the nested path to find the src directory
    nested_path = coverage_subdir
    for part in project_root.parts:
        nested_path = nested_path / part
        if not nested_path.exists():
            break
    
    # Find where 'src' directory is in the nested structure
    src_path = None
    for root, dirs, files in os.walk(coverage_subdir):
        if root.endswith('/src') or '/src/' in root:
            # Found a src directory, use its parent as the base
            root_path = Path(root)
            # Go up to find the project root (where src, tests, etc. are)
            while root_path.name not in ['coverage']:
                if (root_path / 'src').exists() or root_path.name == str(project_root.name):
                    src_path = root_path
                    break
                root_path = root_path.parent
            if src_path:
                break
    
    if not src_path:
        print("Warning: Could not find src directory in nested structure")
        # Try to find any directory that contains the project name
        for root, dirs, files in os.walk(coverage_subdir):
            if project_root.name in root:
                src_path = Path(root)
                break
    
    if src_path and src_path.exists():
        # Move src directory to output_dir/src
        for item in src_path.iterdir():
            dest = output_dir / item.name
            if dest.exists():
                if dest.is_dir():
                    shutil.rmtree(dest)
                else:
                    dest.unlink()
            shutil.move(str(item), str(dest))
    
    # Clean up the html directory
    if html_dir.exists():
        shutil.rmtree(html_dir)
    
    # Post-process HTML files to fix internal paths
    # Pattern to match: href='coverage/home/.../vb6parse/src/...'
    # Replace with: href='src/...'
    project_root_str = str(project_root)
    # Create pattern that matches the nested path structure
    pattern = re.compile(r"(href|src)='coverage/[^']*?/src/", re.IGNORECASE)
    replacement = r"\1='src/"
    
    # Also handle case where paths might just start with the full absolute path
    abs_pattern = re.compile(r"(href|src)='[^']*?/" + re.escape(project_root.name) + r"/src/", re.IGNORECASE)
    
    # Pattern to fix absolute paths in source file titles and link to GitHub
    # Matches: <div class='source-name-title'><pre>/absolute/path/to/vb6parse/src/file.rs</pre></div>
    source_title_pattern = re.compile(
        r"<div class='source-name-title'><pre>.*?/" + re.escape(project_root.name) + r"/(src/[^<]+)</pre></div>",
        re.IGNORECASE
    )
    github_url = "https://github.com/scriptandcompile/vb6parse/blob/master/"
    
    def source_title_replacement(match):
        """Replace absolute path with relative path linked to GitHub."""
        relative_path = match.group(1)  # e.g., "src/lexer/mod.rs"
        return f"<div class='source-name-title'><a href='{github_url}{relative_path}'>{relative_path}</a></div>"
    
    # Theme synchronization script with toggle functionality to inject into HTML files
    theme_script = """<script>
// Sync theme with main site and setup theme toggle
(function() {
    const THEME_KEY = 'vb6parse-theme';
    function getSystemPreference() {
        return (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
    }
    function getTheme() {
        return localStorage.getItem(THEME_KEY) || getSystemPreference();
    }
    function applyTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        const themeIcon = document.querySelector('.theme-icon');
        if (themeIcon) {
            themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
        }
    }
    function toggleTheme() {
        const currentTheme = getTheme();
        const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
        localStorage.setItem(THEME_KEY, newTheme);
        applyTheme(newTheme);
    }
    // Apply theme immediately
    const theme = getTheme();
    applyTheme(theme);
    // Setup toggle button when DOM is ready
    document.addEventListener('DOMContentLoaded', function() {
        const toggle = document.querySelector('.theme-toggle');
        if (toggle) {
            toggle.addEventListener('click', toggleTheme);
        }
    });
})();
</script>"""
    
    # Fix paths in all HTML files in src/ directory
    src_dir = output_dir / 'src'
    if src_dir.exists():
        for html_file in src_dir.rglob('*.html'):
            # Source files need to adjust their relative paths to CSS/JS
            # Calculate correct depth based on directory nesting
            try:
                content = html_file.read_text(encoding='utf-8')
                
                # Calculate how deep this file is relative to src/ directory
                relative_to_src = html_file.relative_to(src_dir)
                depth_from_src = len(relative_to_src.parts) - 1  # -1 because we don't count the file itself
                
                # Calculate path to CSS (up to docs/assets/, then into css/)
                # From src/ we're already 1 level deep in coverage/, so total depth is depth_from_src + 1
                css_depth = depth_from_src + 2  # +1 for src/, +1 for coverage/
                css_base = '../' * css_depth + 'css/'
                
                # Calculate path to control.js (stays in coverage directory)
                js_prefix = '../' * (depth_from_src + 1)  # +1 to go from src/ to coverage/
                
                # Replace CSS references - load BOTH main site style.css AND llvm-cov.css
                css_links = f"<link rel='stylesheet' type='text/css' href='{css_base}style.css'><link rel='stylesheet' type='text/css' href='{css_base}llvm-cov.css'>"
                content = re.sub(r"<link rel='stylesheet' type='text/css' href='(?:\.\./)+style\.css'>", css_links, content)
                
                # Replace JS references (handles excessive ../ paths from llvm-cov)
                content = re.sub(r"src='(?:\.\./)+control\.js'", f"src='{js_prefix}control.js'", content)
                
                # Inject theme script before closing </head> tag
                content = re.sub(r'</head>', f'{theme_script}</head>', content, count=1)
                
                # Inject coverage header after <body> tag (for source files)
                # Calculate path back to main overview from this depth
                docs_depth = depth_from_src + 2  # +1 for src/, +1 for coverage/
                docs_path = '../' * docs_depth + '../index.html'
                
                # Get back to coverage page (main docs coverage, not llvm-cov index)
                coverage_page_path = '../' * docs_depth + '../coverage.html'
                
                header_html = f"""<header>
    <div class="container">
        <h1>VB6Parse Coverage Report</h1>
        <p class="tagline">Generated from llvm-cov</p>
    </div>
</header>
<nav>
    <div class="container">
        <a href='{coverage_page_path}'>Coverage Report</a>
        <a href='{docs_path}'>Overview</a>
        <button id="theme-toggle" class="theme-toggle" aria-label="Toggle theme">
            <span class="theme-icon">🌙</span>
        </button>
    </div>
</nav>"""
                
                # Inject header and nav after <body> tag
                content = re.sub(r'<body>', '<body>' + header_html, content, count=1)
                
                # Remove the old llvm-cov header elements (<h2>Coverage Report</h2><h4>Created: ...</h4>)
                content = re.sub(r'<h2>Coverage Report</h2><h4>Created: [^<]+</h4>', '', content, count=1)
                
                # Fix absolute path in source title and link to GitHub
                content = source_title_pattern.sub(source_title_replacement, content)
                
                # Fix file links
                content = pattern.sub(replacement, content)
                content = abs_pattern.sub(replacement, content)
                
                # Fix internal links to other coverage files (remove source extensions)
                # Pattern: href='path/to/file.rs.html' -> href='path/to/file.html'
                content = re.sub(r"href='([^']*?)\.(rs|toml|md|txt|json|yml|yaml)\.html'", r"href='\1.html'", content)
                
                html_file.write_text(content, encoding='utf-8')
            except Exception as e:
                print(f"Warning: Could not fix paths in {html_file}: {e}")
        
        # Rename .rs.html files to .html (remove the source extension)
        for html_file in src_dir.rglob('*.html'):
            file_name = html_file.name
            # Check if file has a source extension before .html
            if any(file_name.endswith(ext + '.html') for ext in ['.rs', '.toml', '.md', '.txt', '.json', '.yml', '.yaml']):
                # Remove the source extension
                new_name = re.sub(r'\.(rs|toml|md|txt|json|yml|yaml)\.html$', '.html', file_name)
                if new_name != file_name:
                    new_path = html_file.parent / new_name
                    html_file.rename(new_path)
    
    print("  ✓ Restructured HTML files to use workspace-relative paths")


def count_tests_from_list(args):
    """Run cargo test --list and count tests."""
    try:
        result = subprocess.run(
            args,
            capture_output=True,
            text=True,
            check=True
        )
        return len([line for line in result.stdout.split('\n') if ': test' in line])
    except subprocess.CalledProcessError as e:
        print(f"Warning: Failed to count tests for {args}: {e}", file=sys.stderr)
        return 0


def collect_test_statistics():
    """Collect test count breakdown."""
    print("Collecting test statistics...")
    
    # Get library tests (from src/)
    lib_tests = count_tests_from_list(['cargo', 'test', '--lib', '--', '--list'])
    
    # Get doc tests
    doc_tests = count_tests_from_list(['cargo', 'test', '--doc', '--', '--list'])
    
    # Get integration tests by counting each test file
    integration_tests = 0
    test_files = glob.glob('tests/*.rs')
    for test_file in test_files:
        test_name = Path(test_file).stem  # Remove .rs extension
        integration_tests += count_tests_from_list(
            ['cargo', 'test', '--test', test_name, '--', '--list']
        )
    
    # Total test count
    test_count = lib_tests + doc_tests + integration_tests
    
    # Count fuzz targets
    fuzz_dir = Path('fuzz/fuzz_targets')
    fuzz_targets = 0
    if fuzz_dir.exists():
        fuzz_targets = len(list(fuzz_dir.glob('*.rs')))
    
    return {
        'test_count': test_count,
        'lib_tests': lib_tests,
        'doc_tests': doc_tests,
        'integration_tests': integration_tests,
        'fuzz_targets': fuzz_targets
    }


def extract_coverage_metrics():
    """Extract coverage metrics from coverage.json."""
    try:
        with open('docs/assets/data/coverage.json', 'r') as f:
            coverage = json.load(f)
        
        totals = coverage['data'][0]['totals']
        
        return {
            'line_coverage': round(totals['lines']['percent'], 2),
            'function_coverage': round(totals['functions']['percent'], 2),
            'region_coverage': round(totals['regions']['percent'], 2)
        }
    except (IOError, KeyError, IndexError) as e:
        print(f"Error reading coverage data: {e}", file=sys.stderr)
        return {
            'line_coverage': 0.0,
            'function_coverage': 0.0,
            'region_coverage': 0.0
        }


def write_stats(test_stats, coverage_metrics):
    """Write combined statistics to stats.json."""
    stats = {**test_stats, **coverage_metrics}
    
    stats_path = Path('docs/assets/data/stats.json')
    stats_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(stats_path, 'w') as f:
        json.dump(stats, f, indent=2)
    
    # Print summary
    print(f"\nGenerated coverage statistics:")
    print(f"  Total tests: {stats['test_count']:,}")
    print(f"    - Library tests: {stats['lib_tests']:,}")
    print(f"    - Doc tests: {stats['doc_tests']:,}")
    print(f"    - Integration tests: {stats['integration_tests']:,}")
    print(f"    - Fuzz targets: {stats['fuzz_targets']}")
    print(f"  Line coverage: {stats['line_coverage']}%")
    print(f"  Function coverage: {stats['function_coverage']}%")
    print(f"  Region coverage: {stats['region_coverage']}%")
    
    print(f"\n✓ Coverage data saved to docs/assets/data/coverage.json")
    print(f"✓ Test statistics saved to {stats_path}")


def main():
    """Main execution function."""
    try:
        # Generate both JSON and HTML outputs
        run_coverage()
        html_dir = generate_html_coverage()
        
        # Continue with statistics
        test_stats = collect_test_statistics()
        coverage_metrics = extract_coverage_metrics()
        write_stats(test_stats, coverage_metrics)
        
        print(f"\n✓ HTML coverage reports available at {html_dir}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
