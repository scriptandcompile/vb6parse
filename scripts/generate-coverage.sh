#!/bin/bash
# Generate coverage data and test statistics

echo "Generating coverage data..."

# Run coverage
cargo llvm-cov --json --output-path docs/coverage.json

# Generate stats.json with test count breakdown
echo "Collecting test statistics..."
python3 << 'PYTHON'
import json
import subprocess
import os
import glob

# Get library tests (from src/)
lib_result = subprocess.run(['cargo', 'test', '--lib', '--', '--list'], capture_output=True, text=True)
lib_tests = len([line for line in lib_result.stdout.split('\n') if ': test' in line])

# Get doc tests
doc_result = subprocess.run(['cargo', 'test', '--doc', '--', '--list'], capture_output=True, text=True)
doc_tests = len([line for line in doc_result.stdout.split('\n') if ': test' in line])

# Get integration tests by counting each test file
integration_tests = 0
test_files = glob.glob('tests/*.rs')
for test_file in test_files:
    test_name = os.path.basename(test_file)[:-3]  # Remove .rs extension
    result = subprocess.run(
        ['cargo', 'test', '--test', test_name, '--', '--list'],
        capture_output=True, text=True
    )
    integration_tests += len([line for line in result.stdout.split('\n') if ': test' in line])

# Total test count
test_count = lib_tests + doc_tests + integration_tests

# Count fuzz targets
fuzz_dir = 'fuzz/fuzz_targets'
fuzz_targets = 0
if os.path.exists(fuzz_dir):
    fuzz_targets = len([f for f in os.listdir(fuzz_dir) if f.endswith('.rs')])

# Read coverage data
with open('docs/coverage.json', 'r') as f:
    coverage = json.load(f)

totals = coverage['data'][0]['totals']

# Create stats
stats = {
    'test_count': test_count,
    'lib_tests': lib_tests,
    'doc_tests': doc_tests,
    'integration_tests': integration_tests,
    'fuzz_targets': fuzz_targets,
    'line_coverage': round(totals['lines']['percent'], 2),
    'function_coverage': round(totals['functions']['percent'], 2),
    'region_coverage': round(totals['regions']['percent'], 2)
}

# Write stats.json
with open('docs/stats.json', 'w') as f:
    json.dump(stats, f, indent=2)

print(f"\nGenerated coverage statistics:")
print(f"  Total tests: {stats['test_count']:,}")
print(f"    - Library tests: {stats['lib_tests']:,}")
print(f"    - Doc tests: {stats['doc_tests']:,}")
print(f"    - Integration tests: {stats['integration_tests']:,}")
print(f"    - Fuzz targets: {stats['fuzz_targets']}")
print(f"  Line coverage: {stats['line_coverage']}%")
print(f"  Function coverage: {stats['function_coverage']}%")
print(f"  Region coverage: {stats['region_coverage']}%")
PYTHON

echo -e "\n✓ Coverage data saved to docs/coverage.json"
echo "✓ Test statistics saved to docs/stats.json"
