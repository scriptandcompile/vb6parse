#!/bin/bash
# Generate benchmark data for VB6Parse documentation

set -e

echo "Running benchmarks..."
cargo bench --message-format=json > /dev/null 2>&1 || echo "Benchmarks completed with warnings"

echo "Aggregating benchmark data..."
python3 << 'PYTHON'
import json
import os
import glob
from pathlib import Path

# Find all criterion benchmark result files
criterion_dir = "target/criterion"
benchmarks = []

if os.path.exists(criterion_dir):
    # Walk through criterion directory to find estimates.json files
    for estimates_file in glob.glob(f"{criterion_dir}/**/new/estimates.json", recursive=True):
        benchmark_name = Path(estimates_file).parent.parent.name
        
        with open(estimates_file, 'r') as f:
            data = json.load(f)
        
        # Extract key metrics
        mean = data.get('mean', {})
        median = data.get('median', {})
        std_dev = data.get('std_dev', {})
        
        benchmark = {
            'name': benchmark_name,
            'mean': mean.get('point_estimate', 0),
            'median': median.get('point_estimate', 0),
            'std_dev': std_dev.get('point_estimate', 0),
            'unit': 'ns'  # Criterion uses nanoseconds by default
        }
        
        benchmarks.append(benchmark)

# Sort benchmarks by name
benchmarks.sort(key=lambda x: x['name'])

# Write to docs/benchmarks.json
output = {
    'benchmarks': benchmarks,
    'count': len(benchmarks)
}

with open('docs/benchmarks.json', 'w') as f:
    json.dump(output, f, indent=2)

print(f"Generated benchmark data for {len(benchmarks)} benchmarks")
print(json.dumps(output, indent=2))
PYTHON

echo "Benchmark data written to docs/benchmarks.json"
