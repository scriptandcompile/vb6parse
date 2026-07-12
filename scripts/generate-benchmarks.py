#!/usr/bin/env python3
"""
Generate benchmark data for VB6Parse documentation with historical tracking.
Cross-platform script for Windows and Linux.
"""

import json
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

# Configuration
HISTORY_FILE = Path("docs/assets/data/benchmarks-history.json")
SNAPSHOT_FILE = Path("docs/assets/data/benchmarks.json")
RETENTION_DAYS_FULL = 30
RETENTION_DAYS_WEEKLY = 180
RETENTION_DAYS_MONTHLY = 365

CI_COMMIT_PATTERNS = [
    r"\bci\b",
    r"\[skip\s+ci\]",
    r"github\s*actions",
    r"workflow",
    r"merge",
    r"update\s+benchmark\s+data",
]


def run_benchmarks():
    """Run cargo benchmarks."""
    print("Running benchmarks...")
    try:
        subprocess.run(
            ["cargo", "bench", "--message-format=json"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,  # Don't fail on warnings
        )
        print("Benchmarks completed")
    except subprocess.CalledProcessError:
        print("Benchmarks completed with warnings")


def aggregate_benchmark_data():
    """Aggregate benchmark data from Criterion output."""
    print("Aggregating benchmark data...")

    criterion_dir = Path("target/criterion")
    benchmarks = []

    if criterion_dir.exists():
        # Walk through criterion directory to find estimates.json files
        for estimates_file in criterion_dir.rglob("**/new/estimates.json"):
            benchmark_name = estimates_file.parent.parent.name

            try:
                with open(estimates_file, "r") as f:
                    data = json.load(f)

                mean = data.get("mean", {})
                median = data.get("median", {})
                std_dev = data.get("std_dev", {})

                benchmark = {
                    "name": benchmark_name,
                    "mean": mean.get("point_estimate", 0),
                    "median": median.get("point_estimate", 0),
                    "std_dev": std_dev.get("point_estimate", 0),
                    "unit": "ns",  # Criterion uses nanoseconds by default
                }

                benchmarks.append(benchmark)
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Failed to read {estimates_file}: {e}", file=sys.stderr)

    benchmarks.sort(key=lambda x: x["name"])
    return benchmarks


def is_ci_commit_message(message):
    """Return True if a commit message appears to be CI-related."""
    ci_regex = re.compile("|".join(CI_COMMIT_PATTERNS), re.IGNORECASE)
    return bool(ci_regex.search((message or "").strip()))


def get_current_checkout_ref():
    """Get a ref that can restore the current checkout state."""
    try:
        branch = subprocess.run(
            ["git", "rev-parse", "--abbrev-ref", "HEAD"],
            capture_output=True,
            text=True,
            check=True,
        ).stdout.strip()

        if branch and branch != "HEAD":
            return branch

        return subprocess.run(
            ["git", "rev-parse", "HEAD"],
            capture_output=True,
            text=True,
            check=True,
        ).stdout.strip()
    except subprocess.CalledProcessError:
        return None


def restore_checkout_ref(ref):
    """Restore checkout to the original ref so we don't leave git in an odd state."""
    if not ref:
        return

    result = subprocess.run(
        ["git", "checkout", "--quiet", ref],
        capture_output=True,
        text=True,
        check=False,
    )

    if result.returncode != 0:
        err = (result.stderr or "").strip()
        print(f"Warning: Failed to restore git checkout to '{ref}': {err}", file=sys.stderr)


def get_git_info():
    """Get nearest non-CI git commit information by searching backward from HEAD."""
    try:
        log_output = subprocess.run(
            ["git", "log", "--pretty=format:%H%x1f%s", "-n", "500"],
            capture_output=True,
            text=True,
            check=True,
        ).stdout.splitlines()

        selected_sha = None
        selected_msg = None

        # Seek backwards from HEAD and pick the first non-CI commit.
        for line in log_output:
            parts = line.split("\x1f", 1)
            if len(parts) != 2:
                continue

            commit_sha, commit_msg = parts[0].strip(), parts[1].strip()
            if not commit_sha:
                continue

            if not is_ci_commit_message(commit_msg):
                selected_sha = commit_sha
                selected_msg = commit_msg or "No commit message"
                break

        # Fallback to HEAD if all recent commits look CI-related.
        if not selected_sha:
            selected_sha = subprocess.run(
                ["git", "rev-parse", "HEAD"],
                capture_output=True,
                text=True,
                check=True,
            ).stdout.strip()

            selected_msg = subprocess.run(
                ["git", "log", "-1", "--pretty=%s"],
                capture_output=True,
                text=True,
                check=True,
            ).stdout.strip() or "No commit message"

        if selected_sha and selected_msg:
            print(f"Debug: selected non-CI commit {selected_sha[:8]} - {selected_msg}")

        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        return selected_sha, selected_msg, timestamp
    except subprocess.CalledProcessError:
        return (
            "unknown",
            "No git information available",
            datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        )


def load_history():
    """Load existing benchmark history or create new."""
    if HISTORY_FILE.exists():
        try:
            with open(HISTORY_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            print(f"Warning: Failed to load history, starting fresh: {e}", file=sys.stderr)

    return {
        "version": "1.0",
        "last_updated": "",
        "snapshots": [],
        "benchmarks_summary": {},
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

    parsed_snapshots = []
    for snapshot in snapshots:
        try:
            timestamp_str = snapshot["timestamp"].replace("Z", "+00:00")
            timestamp = datetime.fromisoformat(timestamp_str)
            age_days = (now - timestamp).days
            parsed_snapshots.append((timestamp, age_days, snapshot))
        except (ValueError, KeyError) as e:
            print(f"Warning: Skipping malformed snapshot: {e}", file=sys.stderr)
            continue

    parsed_snapshots.sort(key=lambda x: x[0], reverse=True)

    retained = []
    seen_weeks = set()
    seen_months = set()
    seen_quarters = set()

    retention_stats = {
        "full": 0,
        "weekly": 0,
        "monthly": 0,
        "quarterly": 0,
    }

    for timestamp, age_days, snapshot in parsed_snapshots:
        if age_days <= RETENTION_DAYS_FULL:
            retained.append(snapshot)
            retention_stats["full"] += 1
        elif age_days <= RETENTION_DAYS_WEEKLY:
            week_key = (timestamp.year, timestamp.isocalendar()[1])
            if week_key not in seen_weeks:
                retained.append(snapshot)
                seen_weeks.add(week_key)
                retention_stats["weekly"] += 1
        elif age_days <= RETENTION_DAYS_MONTHLY:
            month_key = (timestamp.year, timestamp.month)
            if month_key not in seen_months:
                retained.append(snapshot)
                seen_months.add(month_key)
                retention_stats["monthly"] += 1
        else:
            quarter = (timestamp.month - 1) // 3 + 1
            quarter_key = (timestamp.year, quarter)
            if quarter_key not in seen_quarters:
                retained.append(snapshot)
                seen_quarters.add(quarter_key)
                retention_stats["quarterly"] += 1

    retained_sorted = sorted(retained, key=lambda s: s["timestamp"])

    summary = {
        "removed": original_count - len(retained_sorted),
        "kept": len(retained_sorted),
        "breakdown": retention_stats,
    }

    return retained_sorted, summary


def calculate_trend(history_data):
    """Calculate trend for a specific benchmark."""
    if not history_data or len(history_data) < 2:
        return None

    latest = history_data[-1]
    previous = history_data[-2]

    if previous == 0:
        return None

    change_percent = ((latest - previous) / previous) * 100

    if abs(change_percent) < 0.5:
        direction = "stable"
    elif change_percent < 0:
        direction = "improving"
    else:
        direction = "degrading"

    return {"direction": direction, "change_percent": round(change_percent, 2)}


def update_benchmarks_summary(history, benchmarks):
    """Update the benchmarks summary with historical data and trends."""
    summary = {}

    for benchmark in benchmarks:
        name = benchmark["name"]

        history_data = []
        for snapshot in history["snapshots"]:
            for bench in snapshot["benchmarks"]:
                if bench["name"] == name:
                    history_data.append(
                        {
                            "timestamp": snapshot["timestamp"],
                            "mean": bench["mean"],
                            "median": bench["median"],
                            "std_dev": bench["std_dev"],
                        }
                    )
                    break

        mean_values = [h["mean"] for h in history_data]
        trend = calculate_trend(mean_values) if len(mean_values) >= 2 else None

        summary[name] = {
            "history": history_data[-10:],
            "trend": trend,
            "latest": {
                "mean": benchmark["mean"],
                "median": benchmark["median"],
                "std_dev": benchmark["std_dev"],
            },
        }

    return summary


def write_benchmark_data(benchmarks):
    """Write current benchmark data and update history."""
    output = {"benchmarks": benchmarks, "count": len(benchmarks)}

    SNAPSHOT_FILE.parent.mkdir(parents=True, exist_ok=True)

    with open(SNAPSHOT_FILE, "w") as f:
        json.dump(output, f, indent=2)

    print(f"\n✅ Generated benchmark data for {len(benchmarks)} benchmarks")
    print(f"   Written to {SNAPSHOT_FILE}")

    commit_sha, commit_msg, timestamp = get_git_info()

    history = load_history()

    snapshot = {
        "timestamp": timestamp,
        "commit_sha": commit_sha,
        "commit_message": commit_msg,
        "benchmarks": benchmarks,
    }

    history["snapshots"].append(snapshot)
    history["last_updated"] = timestamp

    print("\n📊 Applying retention policy...")
    before_count = len(history["snapshots"])
    history["snapshots"], retention_summary = apply_retention_policy(history["snapshots"])

    print(f"   Snapshots before retention: {before_count}")
    print(f"   Snapshots retained: {retention_summary['kept']}")
    print(f"   Snapshots removed: {retention_summary['removed']}")
    print("   Breakdown:")
    print(
        f"     - Full retention (0-{RETENTION_DAYS_FULL} days): "
        f"{retention_summary['breakdown']['full']}"
    )
    print(
        f"     - Weekly ({RETENTION_DAYS_FULL + 1}-{RETENTION_DAYS_WEEKLY} days): "
        f"{retention_summary['breakdown']['weekly']}"
    )
    print(
        f"     - Monthly ({RETENTION_DAYS_WEEKLY + 1}-{RETENTION_DAYS_MONTHLY} days): "
        f"{retention_summary['breakdown']['monthly']}"
    )
    print(
        f"     - Quarterly (>{RETENTION_DAYS_MONTHLY} days): "
        f"{retention_summary['breakdown']['quarterly']}"
    )

    history["benchmarks_summary"] = update_benchmarks_summary(history, benchmarks)

    trend_counts = {"improving": 0, "stable": 0, "degrading": 0, "no_data": 0}
    for entry in history["benchmarks_summary"].values():
        if entry["trend"]:
            trend_counts[entry["trend"]["direction"]] += 1
        else:
            trend_counts["no_data"] += 1

    with open(HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)

    snapshot_count = len(history["snapshots"])
    print(f"\n✅ Updated benchmark history ({snapshot_count} snapshots)")
    print(f"   Written to {HISTORY_FILE}")
    print(f"   Commit: {commit_sha[:8]} - {commit_msg}")

    if trend_counts["improving"] + trend_counts["degrading"] + trend_counts["stable"] > 0:
        print("\n📈 Performance Trends:")
        print(f"   Improving: {trend_counts['improving']}")
        print(f"   Stable: {trend_counts['stable']}")
        print(f"   Degrading: {trend_counts['degrading']}")
        if trend_counts["no_data"] > 0:
            print(f"   No historical data: {trend_counts['no_data']}")


def main():
    """Main execution function."""
    original_ref = get_current_checkout_ref()

    try:
        run_benchmarks()
        benchmarks = aggregate_benchmark_data()
        write_benchmark_data(benchmarks)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback

        traceback.print_exc()
        sys.exit(1)
    finally:
        # Always return checkout to where we started.
        restore_checkout_ref(original_ref)


if __name__ == "__main__":
    main()
