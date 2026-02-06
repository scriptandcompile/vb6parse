#!/usr/bin/env python3
"""
Generate coverage data and test statistics for VB6Parse with historical tracking.
Cross-platform script for Windows and Linux.
"""

import json
import os
import sys
import subprocess
import glob
from pathlib import Path
from datetime import datetime, timezone

# Configuration
HISTORY_FILE = Path("docs/assets/data/coverage-history.json")
COVERAGE_FILE = Path("docs/assets/data/coverage.json")
STATS_FILE = Path("docs/assets/data/stats.json")
RETENTION_DAYS_FULL = 30
RETENTION_DAYS_WEEKLY = 180
RETENTION_DAYS_MONTHLY = 365


def get_git_info():
    """Get current git commit information."""
    try:
        commit_sha = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            capture_output=True,
            text=True,
            check=True
        ).stdout.strip()
        
        commit_msg = subprocess.run(
            ["git", "log", "-1", "--pretty=%B"],
            capture_output=True,
            text=True,
            check=True
        ).stdout.strip().split('\n')[0]
        
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        
        return commit_sha, commit_msg, timestamp
    except subprocess.CalledProcessError:
        # If git is not available or not a git repo
        return "unknown", "No git information available", datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def load_history():
    """Load existing coverage history or create new."""
    if HISTORY_FILE.exists():
        try:
            with open(HISTORY_FILE, 'r') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            print(f"Warning: Failed to load history, starting fresh: {e}", file=sys.stderr)
    
    return {
        "version": "1.0",
        "last_updated": "",
        "snapshots": [],
        "coverage_summary": {}
    }


def apply_retention_policy(snapshots):
    """Apply retention policy to limit history size.
    
    Keeps the most recent snapshot from each time period:
    - All snapshots from last 30 days
    - One snapshot per week for days 31-180 (most recent in each week)
    - One snapshot per month for days 181-365 (most recent in each month)
    - One snapshot per quarter beyond 365 days (most recent in each quarter)
    """
    if not snapshots:
        return [], {"removed": 0, "kept": 0}
    
    now = datetime.now(timezone.utc)
    original_count = len(snapshots)
    
    # Parse all valid snapshots with their timestamps
    parsed_snapshots = []
    for snapshot in snapshots:
        try:
            timestamp_str = snapshot['timestamp'].replace('Z', '+00:00')
            timestamp = datetime.fromisoformat(timestamp_str)
            age_days = (now - timestamp).days
            parsed_snapshots.append((timestamp, age_days, snapshot))
        except (ValueError, KeyError) as e:
            print(f"Warning: Skipping malformed snapshot: {e}", file=sys.stderr)
            continue
    
    # Sort by timestamp (newest first for grouping)
    parsed_snapshots.sort(key=lambda x: x[0], reverse=True)
    
    retained = []
    seen_weeks = set()
    seen_months = set()
    seen_quarters = set()
    
    retention_stats = {
        "full": 0,
        "weekly": 0,
        "monthly": 0,
        "quarterly": 0
    }
    
    for timestamp, age_days, snapshot in parsed_snapshots:
        # Keep all from last 30 days
        if age_days <= RETENTION_DAYS_FULL:
            retained.append(snapshot)
            retention_stats["full"] += 1
        # Keep one per week for 31-180 days (ISO week)
        elif age_days <= RETENTION_DAYS_WEEKLY:
            week_key = (timestamp.year, timestamp.isocalendar()[1])
            if week_key not in seen_weeks:
                retained.append(snapshot)
                seen_weeks.add(week_key)
                retention_stats["weekly"] += 1
        # Keep one per month for 181-365 days
        elif age_days <= RETENTION_DAYS_MONTHLY:
            month_key = (timestamp.year, timestamp.month)
            if month_key not in seen_months:
                retained.append(snapshot)
                seen_months.add(month_key)
                retention_stats["monthly"] += 1
        # Keep one per quarter beyond 365 days
        else:
            quarter = (timestamp.month - 1) // 3 + 1
            quarter_key = (timestamp.year, quarter)
            if quarter_key not in seen_quarters:
                retained.append(snapshot)
                seen_quarters.add(quarter_key)
                retention_stats["quarterly"] += 1
    
    # Return in chronological order (oldest first)
    retained_sorted = sorted(retained, key=lambda s: s['timestamp'])
    
    summary = {
        "removed": original_count - len(retained_sorted),
        "kept": len(retained_sorted),
        "breakdown": retention_stats
    }
    
    return retained_sorted, summary


def run_coverage():
    """Run cargo llvm-cov to generate coverage data."""
    print("Generating coverage data...")
    
    try:
        subprocess.run(
            ["cargo", "llvm-cov", "--lib", "--tests", "--json", "--output-path", str(COVERAGE_FILE)],
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
        with open(COVERAGE_FILE, 'r') as f:
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


def create_coverage_snapshot(commit_sha, commit_msg, timestamp, test_stats, coverage_metrics):
    """Create a coverage snapshot from current data."""
    try:
        with open(COVERAGE_FILE, 'r') as f:
            coverage_data = json.load(f)
        
        totals = coverage_data['data'][0]['totals']
        
        return {
            "timestamp": timestamp,
            "commit_sha": commit_sha,
            "commit_message": commit_msg,
            "coverage": {
                "line_coverage": coverage_metrics['line_coverage'],
                "function_coverage": coverage_metrics['function_coverage'],
                "region_coverage": coverage_metrics['region_coverage']
            },
            "tests": {
                "total": test_stats['test_count'],
                "lib_tests": test_stats['lib_tests'],
                "doc_tests": test_stats['doc_tests'],
                "integration_tests": test_stats['integration_tests'],
                "fuzz_targets": test_stats['fuzz_targets']
            },
            "details": {
                "lines": {
                    "covered": totals['lines']['covered'],
                    "total": totals['lines']['count'],
                    "percent": round(totals['lines']['percent'], 2)
                },
                "functions": {
                    "covered": totals['functions']['covered'],
                    "total": totals['functions']['count'],
                    "percent": round(totals['functions']['percent'], 2)
                },
                "regions": {
                    "covered": totals['regions']['covered'],
                    "total": totals['regions']['count'],
                    "percent": round(totals['regions']['percent'], 2)
                }
            }
        }
    except (IOError, KeyError, IndexError) as e:
        print(f"Error creating snapshot: {e}", file=sys.stderr)
        sys.exit(1)


def update_coverage_summary(history):
    """Calculate summary statistics and trends for all metrics."""
    if not history['snapshots']:
        return {}
    
    summary = {}
    
    for metric in ['line_coverage', 'function_coverage', 'region_coverage']:
        values = [s['coverage'][metric] for s in history['snapshots']]
        
        if len(values) >= 2:
            recent = values[-1]
            previous = values[-2]
            change = recent - previous
            
            # Coverage changes are small, use tight threshold
            if abs(change) < 0.1:
                trend = "stable"
            elif change > 0:
                trend = "improving"
            else:
                trend = "degrading"
        else:
            trend = "no_data"
            change = 0
        
        summary[metric] = {
            "latest": values[-1],
            "trend": trend,
            "change_percent": round(abs(change), 2),
            "best": round(max(values), 2),
            "worst": round(min(values), 2),
            "average": round(sum(values) / len(values), 2)
        }
    
    # Test count tracking
    test_counts = [s['tests']['total'] for s in history['snapshots']]
    if len(test_counts) >= 2:
        change = test_counts[-1] - test_counts[-2]
        growth_rate = (change / test_counts[-2]) * 100 if test_counts[-2] > 0 else 0
        
        trend = "growing" if change > 5 else "stable" if change >= -5 else "shrinking"
        
        summary['test_count'] = {
            "latest": test_counts[-1],
            "trend": trend,
            "change_count": change,
            "growth_rate": round(growth_rate, 2)
        }
    
    return summary


def write_stats(test_stats, coverage_metrics):
    """Write combined statistics to stats.json."""
    stats = {**test_stats, **coverage_metrics}
    
    STATS_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    with open(STATS_FILE, 'w') as f:
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
    
    print(f"\n✓ Coverage data saved to {COVERAGE_FILE}")
    print(f"✓ Test statistics saved to {STATS_FILE}")


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
        
        # Historical tracking
        print("\n📊 Updating coverage history...")
        commit_sha, commit_msg, timestamp = get_git_info()
        
        # Load and update history
        history = load_history()
        snapshot = create_coverage_snapshot(
            commit_sha, commit_msg, timestamp,
            test_stats, coverage_metrics
        )
        
        history['snapshots'].append(snapshot)
        history['last_updated'] = timestamp
        
        # Apply retention policy
        before_count = len(history['snapshots'])
        history['snapshots'], retention_summary = apply_retention_policy(
            history['snapshots']
        )
        
        # Update summary with trends
        history['coverage_summary'] = update_coverage_summary(history)
        
        # Write history file
        HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(HISTORY_FILE, 'w') as f:
            json.dump(history, f, indent=2)
        
        # Print summary
        print(f"   Snapshots before retention: {before_count}")
        print(f"   Snapshots retained: {retention_summary['kept']}")
        print(f"   Snapshots removed: {retention_summary['removed']}")
        print(f"\n✅ Coverage history saved to {HISTORY_FILE}")
        print(f"   {len(history['snapshots'])} total snapshots")
        print(f"   Commit: {commit_sha[:8]} - {commit_msg}")
        
        # Show trends
        summary = history['coverage_summary']
        if summary:
            print(f"\n📈 Coverage Trends:")
            print(f"   Line: {summary['line_coverage']['latest']}% "
                  f"({summary['line_coverage']['trend']})")
            print(f"   Function: {summary['function_coverage']['latest']}% "
                  f"({summary['function_coverage']['trend']})")
            print(f"   Region: {summary['region_coverage']['latest']}% "
                  f"({summary['region_coverage']['trend']})")
            if 'test_count' in summary:
                print(f"   Tests: {summary['test_count']['latest']:,} "
                      f"({summary['test_count']['trend']}, "
                      f"{summary['test_count']['change_count']:+d})")
        
        print(f"\n✓ HTML coverage reports available at {html_dir}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
