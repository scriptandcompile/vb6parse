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
    
    print(f"\n✅ Generated benchmark data for {len(benchmarks)} benchmarks")
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
    print(f"\n📊 Applying retention policy...")
    before_count = len(history["snapshots"])
    history["snapshots"], retention_summary = apply_retention_policy(history["snapshots"])
    
    print(f"   Snapshots before retention: {before_count}")
    print(f"   Snapshots retained: {retention_summary['kept']}")
    print(f"   Snapshots removed: {retention_summary['removed']}")
    print(f"   Breakdown:")
    print(f"     - Full retention (0-{RETENTION_DAYS_FULL} days): {retention_summary['breakdown']['full']}")
    print(f"     - Weekly ({RETENTION_DAYS_FULL+1}-{RETENTION_DAYS_WEEKLY} days): {retention_summary['breakdown']['weekly']}")
    print(f"     - Monthly ({RETENTION_DAYS_WEEKLY+1}-{RETENTION_DAYS_MONTHLY} days): {retention_summary['breakdown']['monthly']}")
    print(f"     - Quarterly (>{RETENTION_DAYS_MONTHLY} days): {retention_summary['breakdown']['quarterly']}")
    
    # Update benchmarks summary with trends
    history["benchmarks_summary"] = update_benchmarks_summary(history, benchmarks)
    
    # Count trends
    trend_counts = {"improving": 0, "stable": 0, "degrading": 0, "no_data": 0}
    for summary in history["benchmarks_summary"].values():
        if summary["trend"]:
            trend_counts[summary["trend"]["direction"]] += 1
        else:
            trend_counts["no_data"] += 1
    
    # Write history file
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, indent=2)
    
    snapshot_count = len(history["snapshots"])
    print(f"\n✅ Updated benchmark history ({snapshot_count} snapshots)")
    print(f"   Written to {HISTORY_FILE}")
    print(f"   Commit: {commit_sha[:8]} - {commit_msg}")
    
    if trend_counts["improving"] + trend_counts["degrading"] + trend_counts["stable"] > 0:
        print(f"\n📈 Performance Trends:")
        print(f"   Improving: {trend_counts['improving']}")
        print(f"   Stable: {trend_counts['stable']}")
        print(f"   Degrading: {trend_counts['degrading']}")
        if trend_counts["no_data"] > 0:
            print(f"   No historical data: {trend_counts['no_data']}")


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
