#!/usr/bin/env python3
"""
Backfill benchmark history by running real benchmarks for selected non-CI commits.

Default selection policy:
- Last 30 days: all non-CI commits.
- Days 31-90: one non-CI commit per day (if available).
- Days 91-183 (~6 months): up to 3 non-CI commits per month, evenly spaced.
- Older than 6 months: one non-CI commit per month, closest to month midpoint.

This script executes `cargo bench --message-format=json` for each selected commit
inside an isolated git worktree and aggregates Criterion results from that commit.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path


DEFAULT_OUTPUT_FILE = Path("docs/assets/data/benchmarks-history.json")

CI_COMMIT_PATTERNS = [
    r"\bci\b",
    r"\[skip\s+ci\]",
    r"github\s*actions",
    r"workflow",
    r"merge",
    r"update\s+benchmark\s+data",
]


@dataclass(frozen=True)
class Commit:
    sha: str
    timestamp: datetime
    subject: str


@dataclass(frozen=True)
class SelectedCommit:
    commit: Commit
    bucket: str


@dataclass(frozen=True)
class SnapshotBuildResult:
    status: str
    note: str
    benchmarks: list[dict]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Backfill benchmark history from non-CI commits by running cargo bench "
            "for each selected commit."
        )
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_FILE,
        help="Path to write benchmarks-history.json",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only print selected commits; do not run benchmarks or write files",
    )
    parser.add_argument(
        "--max-commits",
        type=int,
        default=0,
        help="Limit number of selected commits (0 means no limit)",
    )
    parser.add_argument(
        "--keep-worktree",
        action="store_true",
        help="Keep temporary worktree directory for debugging",
    )
    return parser.parse_args()


def run_cmd(cmd: list[str], cwd: Path | None = None, check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, check=check)


def summarize_compiler_errors(output: str, max_items: int = 4) -> str:
    """Extract actionable compiler errors from cargo output.

    cargo with --message-format=json emits JSON lines on stdout. Fall back to
    plain-text parsing if JSON decoding fails.
    """
    if not output:
        return ""

    errors: list[str] = []
    seen: set[str] = set()

    for line in output.splitlines():
        line = line.strip()
        if not line:
            continue

        if line.startswith("{") and '"reason"' in line:
            try:
                msg = json.loads(line)
                if msg.get("reason") != "compiler-message":
                    continue
                message = msg.get("message", {})
                if message.get("level") != "error":
                    continue
                text = (message.get("message") or "").strip()
                if text and text not in seen:
                    seen.add(text)
                    errors.append(text)
            except json.JSONDecodeError:
                pass
            continue

        lower = line.lower()
        if "error:" in lower:
            text = line
            if text not in seen:
                seen.add(text)
                errors.append(text)

    if not errors:
        return ""

    return " | ".join(errors[:max_items])


def get_repo_root() -> Path:
    output = run_cmd(["git", "rev-parse", "--show-toplevel"]).stdout.strip()
    if not output:
        raise RuntimeError("Unable to detect git repository root")
    return Path(output)


def load_non_ci_commits() -> list[Commit]:
    cmd = [
        "git",
        "log",
        "--reverse",
        "--date=iso-strict",
        "--pretty=format:%H%x1f%cI%x1f%s",
    ]
    completed = run_cmd(cmd)
    lines = completed.stdout.splitlines()

    ci_regex = re.compile("|".join(CI_COMMIT_PATTERNS), re.IGNORECASE)
    commits: list[Commit] = []

    for line in lines:
        parts = line.split("\x1f")
        if len(parts) != 3:
            continue

        sha, iso_timestamp, subject = parts
        if ci_regex.search(subject.strip()):
            continue

        try:
            timestamp = datetime.fromisoformat(iso_timestamp)
        except ValueError:
            continue

        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=timezone.utc)

        commits.append(Commit(sha=sha, timestamp=timestamp, subject=subject.strip()))

    if not commits:
        raise RuntimeError("No non-CI commits found in git history")

    return commits


def age_days(commit: Commit, now: datetime) -> int:
    return (now - commit.timestamp).days


def pick_commit_closest_to_time(commits: list[Commit], target_time: datetime) -> Commit:
    return min(commits, key=lambda c: abs((c.timestamp - target_time).total_seconds()))


def select_one_per_day(commits: list[Commit]) -> list[Commit]:
    by_day: dict[tuple[int, int, int], list[Commit]] = {}
    for commit in commits:
        dt = commit.timestamp.astimezone(timezone.utc)
        key = (dt.year, dt.month, dt.day)
        by_day.setdefault(key, []).append(commit)

    selected: list[Commit] = []
    for (year, month, day), day_commits in sorted(by_day.items()):
        target = datetime(year, month, day, 12, 0, 0, tzinfo=timezone.utc)
        selected.append(pick_commit_closest_to_time(day_commits, target))

    return selected


def evenly_spaced(commits: list[Commit], count: int) -> list[Commit]:
    if count <= 0 or not commits:
        return []
    if len(commits) <= count:
        return commits.copy()

    max_index = len(commits) - 1
    indices = []
    for i in range(count):
        idx = round(i * max_index / (count - 1)) if count > 1 else 0
        indices.append(idx)

    unique_indices = sorted(set(indices))
    return [commits[i] for i in unique_indices]


def select_three_per_month_evenly(commits: list[Commit]) -> list[Commit]:
    by_month: dict[tuple[int, int], list[Commit]] = {}
    for commit in commits:
        dt = commit.timestamp.astimezone(timezone.utc)
        key = (dt.year, dt.month)
        by_month.setdefault(key, []).append(commit)

    selected: list[Commit] = []
    for key in sorted(by_month.keys()):
        month_commits = sorted(by_month[key], key=lambda c: c.timestamp)
        selected.extend(evenly_spaced(month_commits, 3))

    return selected


def month_midpoint(year: int, month: int) -> datetime:
    month_start = datetime(year, month, 1, 0, 0, 0, tzinfo=timezone.utc)
    if month == 12:
        next_month = datetime(year + 1, 1, 1, 0, 0, 0, tzinfo=timezone.utc)
    else:
        next_month = datetime(year, month + 1, 1, 0, 0, 0, tzinfo=timezone.utc)
    delta = next_month - month_start
    return month_start + timedelta(seconds=delta.total_seconds() / 2)


def select_one_per_month_midpoint(commits: list[Commit]) -> list[Commit]:
    by_month: dict[tuple[int, int], list[Commit]] = {}
    for commit in commits:
        dt = commit.timestamp.astimezone(timezone.utc)
        key = (dt.year, dt.month)
        by_month.setdefault(key, []).append(commit)

    selected: list[Commit] = []
    for (year, month) in sorted(by_month.keys()):
        target = month_midpoint(year, month)
        selected.append(pick_commit_closest_to_time(by_month[(year, month)], target))

    return selected


def select_commits_by_policy(commits: list[Commit]) -> list[SelectedCommit]:
    now = datetime.now(timezone.utc)

    last_30 = [c for c in commits if 0 <= age_days(c, now) <= 30]
    days_31_90 = [c for c in commits if 31 <= age_days(c, now) <= 90]
    days_91_183 = [c for c in commits if 91 <= age_days(c, now) <= 183]
    older_than_183 = [c for c in commits if age_days(c, now) >= 184]

    selected: list[SelectedCommit] = []

    for commit in sorted(last_30, key=lambda c: c.timestamp):
        selected.append(SelectedCommit(commit=commit, bucket="last_30_days"))

    for commit in select_one_per_day(days_31_90):
        selected.append(SelectedCommit(commit=commit, bucket="days_31_90_daily"))

    for commit in select_three_per_month_evenly(days_91_183):
        selected.append(SelectedCommit(commit=commit, bucket="days_91_183_three_per_month"))

    for commit in select_one_per_month_midpoint(older_than_183):
        selected.append(SelectedCommit(commit=commit, bucket="older_than_6_months_midpoint_monthly"))

    deduped: list[SelectedCommit] = []
    seen: set[str] = set()
    for entry in sorted(selected, key=lambda s: s.commit.timestamp):
        if entry.commit.sha in seen:
            continue
        seen.add(entry.commit.sha)
        deduped.append(entry)

    return deduped


def commit_has_benchmark_support(commit_sha: str) -> bool:
    result = run_cmd(
        ["git", "ls-tree", "-d", "--name-only", commit_sha, "benches"],
        check=False,
    )
    return result.returncode == 0 and "benches" in result.stdout.splitlines()


def aggregate_criterion_data(criterion_dir: Path) -> list[dict]:
    benchmarks: list[dict] = []

    if not criterion_dir.exists():
        return benchmarks

    for estimates_file in criterion_dir.rglob("**/new/estimates.json"):
        benchmark_name = estimates_file.parent.parent.name

        try:
            with estimates_file.open("r", encoding="utf-8") as f:
                data = json.load(f)

            mean = data.get("mean", {})
            median = data.get("median", {})
            std_dev = data.get("std_dev", {})

            benchmarks.append(
                {
                    "name": benchmark_name,
                    "mean": mean.get("point_estimate", 0),
                    "median": median.get("point_estimate", 0),
                    "std_dev": std_dev.get("point_estimate", 0),
                    "unit": "ns",
                }
            )
        except (json.JSONDecodeError, IOError):
            continue

    benchmarks.sort(key=lambda x: x["name"])
    return benchmarks


def run_benchmarks_for_commit(worktree_dir: Path, commit: Commit) -> SnapshotBuildResult:
    if not commit_has_benchmark_support(commit.sha):
        return SnapshotBuildResult(
            status="no_data_pre_benchmark",
            note="No benchmark data available for this commit (pre-benchmark era).",
            benchmarks=[],
        )

    checkout = run_cmd(["git", "checkout", "--quiet", commit.sha], cwd=worktree_dir, check=False)
    if checkout.returncode != 0:
        tail = (checkout.stderr or "").strip().splitlines()
        note = tail[-1] if tail else "Failed to checkout commit in benchmark worktree."
        return SnapshotBuildResult(status="benchmark_failed", note=note, benchmarks=[])

    # Ensure benchmark fixture submodules are present for this commit.
    submodule_sync = run_cmd(["git", "submodule", "sync", "--recursive"], cwd=worktree_dir, check=False)
    if submodule_sync.returncode != 0:
        err = (submodule_sync.stderr or "").strip().splitlines()
        note = err[-1] if err else "Failed to sync submodules for commit."
        return SnapshotBuildResult(status="benchmark_failed", note=note, benchmarks=[])

    submodule_update = run_cmd(
        ["git", "submodule", "update", "--init", "--recursive", "--jobs", str(max(1, os.cpu_count() or 1))],
        cwd=worktree_dir,
        check=False,
    )
    if submodule_update.returncode != 0:
        err = (submodule_update.stderr or "").strip().splitlines()
        note = err[-1] if err else "Failed to update submodules for commit."
        return SnapshotBuildResult(status="benchmark_failed", note=note, benchmarks=[])

    criterion_dir = worktree_dir / "target" / "criterion"
    if criterion_dir.exists():
        shutil.rmtree(criterion_dir, ignore_errors=True)

    bench = run_cmd(
        ["cargo", "bench", "--message-format=json"],
        cwd=worktree_dir,
        check=False,
    )

    if bench.returncode != 0:
        summary = summarize_compiler_errors(bench.stdout)
        if not summary:
            summary = summarize_compiler_errors(bench.stderr)

        if summary:
            note = f"cargo bench failed: {summary}"
        else:
            tail = (bench.stderr or bench.stdout or "").strip().splitlines()
            note = tail[-1] if tail else f"cargo bench exited with code {bench.returncode}."
        return SnapshotBuildResult(status="benchmark_failed", note=note, benchmarks=[])

    benchmarks = aggregate_criterion_data(criterion_dir)
    if not benchmarks:
        return SnapshotBuildResult(
            status="no_benchmark_output",
            note="cargo bench succeeded but no Criterion estimates were found.",
            benchmarks=[],
        )

    return SnapshotBuildResult(status="ok", note="", benchmarks=benchmarks)


def calculate_trend(values: list[float]) -> dict | None:
    if len(values) < 2:
        return None

    previous = values[-2]
    latest = values[-1]
    if previous == 0:
        return None

    change_percent = ((latest - previous) / previous) * 100.0
    if abs(change_percent) < 0.5:
        direction = "stable"
    elif change_percent < 0:
        direction = "improving"
    else:
        direction = "degrading"

    return {"direction": direction, "change_percent": round(change_percent, 2)}


def build_benchmarks_summary(snapshots: list[dict]) -> dict:
    by_name: dict[str, list[dict]] = {}

    for snapshot in snapshots:
        ts = snapshot["timestamp"]
        for benchmark in snapshot.get("benchmarks", []):
            name = benchmark["name"]
            by_name.setdefault(name, []).append(
                {
                    "timestamp": ts,
                    "mean": benchmark["mean"],
                    "median": benchmark["median"],
                    "std_dev": benchmark["std_dev"],
                }
            )

    summary: dict[str, dict] = {}
    for name, history in by_name.items():
        history.sort(key=lambda h: h["timestamp"])
        means = [h["mean"] for h in history]
        trend = calculate_trend(means)
        latest = history[-1]

        summary[name] = {
            "history": history[-10:],
            "trend": trend,
            "latest": {
                "mean": latest["mean"],
                "median": latest["median"],
                "std_dev": latest["std_dev"],
            },
        }

    return summary


def iso_z(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def main() -> int:
    args = parse_args()

    try:
        repo_root = get_repo_root()
        commits = load_non_ci_commits()
        selected = select_commits_by_policy(commits)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    if args.max_commits > 0:
        selected = selected[: args.max_commits]

    if not selected:
        print("Error: No commits selected for backfill", file=sys.stderr)
        return 1

    bucket_counts: dict[str, int] = {}
    for entry in selected:
        bucket_counts[entry.bucket] = bucket_counts.get(entry.bucket, 0) + 1

    print(f"Selected non-CI commits: {len(selected)}")
    for bucket in sorted(bucket_counts.keys()):
        print(f"  - {bucket}: {bucket_counts[bucket]}")
    print(f"Date range: {iso_z(selected[0].commit.timestamp)} -> {iso_z(selected[-1].commit.timestamp)}")

    if args.dry_run:
        print("Dry run complete; no benchmarks executed and no files written.")
        return 0

    tmpdir_obj = tempfile.TemporaryDirectory(prefix="vb6parse-bench-backfill-")
    worktree_path = Path(tmpdir_obj.name)

    snapshots: list[dict] = []
    status_counts: dict[str, int] = {}

    try:
        add = run_cmd(
            ["git", "worktree", "add", "--detach", str(worktree_path), "HEAD"],
            cwd=repo_root,
            check=False,
        )
        if add.returncode != 0:
            err = (add.stderr or "").strip()
            raise RuntimeError(f"Failed to create benchmark worktree: {err}")

        total = len(selected)
        for index, entry in enumerate(selected, start=1):
            commit = entry.commit
            print(
                f"[{index}/{total}] Benchmarking {commit.sha[:8]} "
                f"({entry.bucket}) - {commit.subject}"
            )

            result = run_benchmarks_for_commit(worktree_path, commit)
            status_counts[result.status] = status_counts.get(result.status, 0) + 1

            snapshots.append(
                {
                    "timestamp": iso_z(commit.timestamp),
                    "commit_sha": commit.sha,
                    "commit_message": commit.subject,
                    "selection_bucket": entry.bucket,
                    "benchmark_status": result.status,
                    "benchmark_note": result.note,
                    "benchmarks": result.benchmarks,
                }
            )

            if result.status != "ok":
                print(f"    status={result.status}: {result.note}")
            else:
                print(f"    benchmarks={len(result.benchmarks)}")

    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    finally:
        if not args.keep_worktree:
            rm = run_cmd(
                ["git", "worktree", "remove", "--force", str(worktree_path)],
                cwd=repo_root,
                check=False,
            )
            if rm.returncode != 0:
                err = (rm.stderr or "").strip()
                print(f"Warning: Failed to remove worktree {worktree_path}: {err}", file=sys.stderr)
            tmpdir_obj.cleanup()
        else:
            print(f"Debug: keeping worktree at {worktree_path}")

    history = {
        "version": "1.2",
        "last_updated": iso_z(datetime.now(timezone.utc)),
        "snapshots": snapshots,
        "benchmarks_summary": build_benchmarks_summary(snapshots),
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8") as f:
        json.dump(history, f, indent=2)

    print(f"Wrote backfilled benchmark history: {args.output}")
    for status in sorted(status_counts.keys()):
        print(f"  - {status}: {status_counts[status]}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
