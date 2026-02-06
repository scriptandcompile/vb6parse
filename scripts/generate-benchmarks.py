#!/usr/bin/env python3
"""
Generate benchmark data for VB6Parse documentation.
Cross-platform script for Windows and Linux.
"""

import json
import os
import sys
import subprocess
import glob
from pathlib import Path


def run_benchmarks():
    """Run cargo benchmarks."""
    print("Running benchmarks...")
    try:
        subprocess.run(
            ["cargo", "bench", "--message-format=json"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False  # Don't fail on warnings
        )
        print("Benchmarks completed")
    except subprocess.CalledProcessError:
        print("Benchmarks completed with warnings")


def aggregate_benchmark_data():
    """Aggregate benchmark data from Criterion output."""
    print("Aggregating benchmark data...")
    
    # Find all criterion benchmark result files
    criterion_dir = Path("target/criterion")
    benchmarks = []
    
    if criterion_dir.exists():
        # Walk through criterion directory to find estimates.json files
        for estimates_file in criterion_dir.rglob("**/new/estimates.json"):
            benchmark_name = estimates_file.parent.parent.name
            
            try:
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
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Failed to read {estimates_file}: {e}", file=sys.stderr)
    
    # Sort benchmarks by name
    benchmarks.sort(key=lambda x: x['name'])
    
    return benchmarks


def write_benchmark_data(benchmarks):
    """Write benchmark data to JSON file."""
    output = {
        'benchmarks': benchmarks,
        'count': len(benchmarks)
    }
    
    output_path = Path('docs/assets/data/benchmarks.json')
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, 'w') as f:
        json.dump(output, f, indent=2)
    
    print(f"Generated benchmark data for {len(benchmarks)} benchmarks")
    print(json.dumps(output, indent=2))
    print(f"\nBenchmark data written to {output_path}")


def main():
    """Main execution function."""
    try:
        run_benchmarks()
        benchmarks = aggregate_benchmark_data()
        write_benchmark_data(benchmarks)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
