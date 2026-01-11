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
            ["cargo", "llvm-cov", "--json", "--output-path", "docs/coverage.json"],
            check=True
        )
    except subprocess.CalledProcessError as e:
        print(f"Error running coverage: {e}", file=sys.stderr)
        sys.exit(1)


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
        with open('docs/coverage.json', 'r') as f:
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
    
    stats_path = Path('docs/stats.json')
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
    
    print(f"\n✓ Coverage data saved to docs/coverage.json")
    print(f"✓ Test statistics saved to {stats_path}")


def main():
    """Main execution function."""
    try:
        run_coverage()
        test_stats = collect_test_statistics()
        coverage_metrics = extract_coverage_metrics()
        write_stats(test_stats, coverage_metrics)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
