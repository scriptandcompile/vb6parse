#!/usr/bin/env python3
"""
Generate benchmark data for VB6Parse documentation with historical tracking.
Cross-platform script for Windows and Linux.
"""

import json
import os
import sys
import subprocess
import glob
from pathlib import Path
from datetime import datetime, timedelta, timezone

# Configuration
HISTORY_FILE = Path("docs/assets/data/benchmarks-history.json")
SNAPSHOT_FILE = Path("docs/assets/data/benchmarks.json")
RETENTION_DAYS_FULL = 30
RETENTION_DAYS_WEEKLY = 180
RETENTION_DAYS_MONTHLY = 365


def run_benchmarks():
    """Run cargo benchmarks with progress updates."""
    print("Running benchmarks...")
    print("This may take several minutes...")
    print()
    
    try:
        # Use unbuffered output
        env = os.environ.copy()
        env['RUST_BACKTRACE'] = '0'
        
        process = subprocess.Popen(
            ["cargo", "bench"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
            env=env
        )
        
        benchmark_count = 0
        test_count = 0
        in_test_phase = False
        
        for line in process.stdout:
            line = line.rstrip()
            if not line:
                continue
            
            # Detect test phase vs benchmark phase
            if "running" in line.lower() and "test" in line.lower():
                if "bench" not in line.lower():
                    in_test_phase = True
                    print(f"🧪 {line}")
                else:
                    in_test_phase = False
                    print(f"🔄 {line}")
            elif in_test_phase and ("test " in line or "running" in line):
                # Show condensed test progress (every 100 tests)
                if "ok" in line.lower():
                    test_count += 1
                    if test_count % 100 == 0:
                        print(f"  ... {test_count} tests completed", end='\r', flush=True)
            elif "test result:" in line.lower():
                if in_test_phase:
                    print(f"\n✅ {line}")
                    in_test_phase = False
                else:
                    print(f"📊 {line}")
            elif "bench:" in line:
                benchmark_count += 1
                print(f"  {line}")
            elif any(keyword in line for keyword in ["Benchmarking", "Collecting", "Analyzing", "Warming up"]):
                # Criterion.rs output
                print(f"  {line}")
        
        process.wait()
        print()
        if benchmark_count > 0:
            print(f"✅ Benchmarks completed ({benchmark_count} results)")
        else:
            print(f"✅ Process completed")
        
    except KeyboardInterrupt:
        print("\n⚠️  Benchmark run interrupted by user")
        sys.exit(130)
    except Exception as e:
        print(f"\n⚠️  Error during benchmark run: {e}")
        sys.exit(130)
    except Exception as e:
        print(f"\n⚠️  Benchmarks completed with errors: {e}")


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
    """Load existing benchmark history or create new."""
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
        "benchmarks_summary": {}
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
        return []
    
    now = datetime.now(timezone.utc)
    
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
    
    for timestamp, age_days, snapshot in parsed_snapshots:
        # Keep all from last 30 days
        if age_days <= RETENTION_DAYS_FULL:
            retained.append(snapshot)
        # Keep one per week for 31-180 days (ISO week)
        elif age_days <= RETENTION_DAYS_WEEKLY:
            week_key = (timestamp.year, timestamp.isocalendar()[1])
            if week_key not in seen_weeks:
                retained.append(snapshot)
                seen_weeks.add(week_key)
        # Keep one per month for 181-365 days
        elif age_days <= RETENTION_DAYS_MONTHLY:
            month_key = (timestamp.year, timestamp.month)
            if month_key not in seen_months:
                retained.append(snapshot)
                seen_months.add(month_key)
        # Keep one per quarter beyond 365 days
        else:
            quarter = (timestamp.month - 1) // 3 + 1
            quarter_key = (timestamp.year, quarter)
            if quarter_key not in seen_quarters:
                retained.append(snapshot)
                seen_quarters.add(quarter_key)
    
    # Return in chronological order (oldest first)
    return sorted(retained, key=lambda s: s['timestamp'])


def calculate_trend(history_data, benchmark_name):
    """Calculate trend for a specific benchmark."""
    if not history_data or len(history_data) < 2:
        return None
    
    # Get last two entries
    latest = history_data[-1]
    previous = history_data[-2]
    
    change_percent = ((latest - previous) / previous) * 100
    
    if abs(change_percent) < 0.5:
        direction = "stable"
    elif change_percent < 0:
        direction = "improving"
    else:
        direction = "degrading"
    
    return {
        "direction": direction,
        "change_percent": round(change_percent, 2)
    }


def update_benchmarks_summary(history, benchmarks):
    """Update the benchmarks summary with historical data and trends."""
    summary = {}
    
    for benchmark in benchmarks:
        name = benchmark['name']
        
        # Collect historical data for this benchmark
        history_data = []
        for snapshot in history['snapshots']:
            for bench in snapshot['benchmarks']:
                if bench['name'] == name:
                    history_data.append({
                        'timestamp': snapshot['timestamp'],
                        'mean': bench['mean'],
                        'median': bench['median'],
                        'std_dev': bench['std_dev']
                    })
                    break
        
        # Calculate trend
        mean_values = [h['mean'] for h in history_data]
        trend = calculate_trend(mean_values, name) if len(mean_values) >= 2 else None
        
        summary[name] = {
            'history': history_data[-10:],  # Keep last 10 data points for quick access
            'trend': trend,
            'latest': {
                'mean': benchmark['mean'],
                'median': benchmark['median'],
                'std_dev': benchmark['std_dev']
            }
        }
    
    return summary


def write_benchmark_data(benchmarks):
    """Write current benchmark data and update history."""
    # Write current snapshot
    output = {
        'benchmarks': benchmarks,
        'count': len(benchmarks)
    }
    
    SNAPSHOT_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    with open(SNAPSHOT_FILE, 'w') as f:
        json.dump(output, f, indent=2)
    
    print(f"✅ Generated benchmark data for {len(benchmarks)} benchmarks")
    print(f"   Written to {SNAPSHOT_FILE}")
    
    # Update history
    commit_sha, commit_msg, timestamp = get_git_info()
    
    history = load_history()
    
    # Add new snapshot
    snapshot = {
        "timestamp": timestamp,
        "commit_sha": commit_sha,
        "commit_message": commit_msg,
        "benchmarks": benchmarks
    }
    
    history["snapshots"].append(snapshot)
    history["last_updated"] = timestamp
    
    # Apply retention policy
    history["snapshots"] = apply_retention_policy(history["snapshots"])
    
    # Update benchmarks summary with trends
    history["benchmarks_summary"] = update_benchmarks_summary(history, benchmarks)
    
    # Write history file
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=2)
    
    snapshot_count = len(history["snapshots"])
    print(f"✅ Updated benchmark history ({snapshot_count} snapshots)")
    print(f"   Written to {HISTORY_FILE}")
    print(f"   Commit: {commit_sha[:8]} - {commit_msg}")


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
