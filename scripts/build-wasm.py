#!/usr/bin/env python3
"""
Cross-platform WASM build script for VB6Parse playground.
Works on Windows, macOS, and Linux, both locally and in GitHub Actions.

Requirements:
- Python 3.7+
- wasm-pack (installed via: cargo install wasm-pack)
- wasm-opt (optional, installed via: cargo install wasm-opt)

Usage:
    python scripts/build-wasm.py [--optimize] [--no-typescript]
"""

import argparse
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


def find_executable(name: str) -> str | None:
    """
    Find an executable in PATH, handling Windows .exe extension.
    Returns the full path to the executable or None if not found.
    """
    # On Windows, check with .exe extension
    if platform.system() == "Windows":
        executable = shutil.which(f"{name}.exe")
        if executable:
            return executable
    
    # Try without extension (works on all platforms)
    executable = shutil.which(name)
    return executable


def check_requirements() -> tuple[str, str | None]:
    """
    Check if required tools are installed.
    Returns tuple of (wasm_pack_path, wasm_opt_path)
    """
    wasm_pack = find_executable("wasm-pack")
    if not wasm_pack:
        print("❌ Error: wasm-pack not found in PATH", file=sys.stderr)
        print("   Install with: cargo install wasm-pack", file=sys.stderr)
        sys.exit(1)
    
    wasm_opt = find_executable("wasm-opt")
    if not wasm_opt:
        print("⚠️  Warning: wasm-opt not found (optional optimization will be skipped)")
        print("   Install with: cargo install wasm-opt", file=sys.stderr)
    
    return wasm_pack, wasm_opt


def run_command(cmd: list[str], description: str) -> None:
    """Run a command and handle errors."""
    print(f"🔨 {description}...")
    try:
        result = subprocess.run(
            cmd,
            check=True,
            capture_output=True,
            text=True
        )
        if result.stdout:
            print(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"❌ Error: {description} failed", file=sys.stderr)
        if e.stdout:
            print(e.stdout, file=sys.stderr)
        if e.stderr:
            print(e.stderr, file=sys.stderr)
        sys.exit(1)


def build_wasm(wasm_pack: str, output_dir: Path, no_typescript: bool) -> None:
    """Build the WASM module using wasm-pack."""
    cmd = [
        wasm_pack,
        "build",
        "--target", "web",
        "--out-dir", str(output_dir),
        "--release"
    ]
    
    # Add --no-typescript flag if requested (reduces output files)
    if no_typescript:
        cmd.append("--no-typescript")
    
    run_command(cmd, "Building WASM module with wasm-pack")


def optimize_wasm(wasm_opt: str | None, wasm_file: Path) -> None:
    """Optimize WASM binary using wasm-opt if available."""
    if not wasm_opt:
        print("⏩ Skipping wasm-opt optimization (not installed)")
        return
    
    if not wasm_file.exists():
        print(f"⚠️  Warning: {wasm_file} not found, skipping optimization")
        return
    
    # Create backup
    backup_file = wasm_file.with_suffix(".wasm.bak")
    shutil.copy2(wasm_file, backup_file)
    
    try:
        cmd = [
            wasm_opt,
            "-Oz",  # Optimize aggressively for size
            "-o", str(wasm_file),
            str(backup_file)
        ]
        run_command(cmd, "Optimizing WASM binary with wasm-opt")
        
        # Show size comparison
        original_size = backup_file.stat().st_size
        optimized_size = wasm_file.stat().st_size
        savings = original_size - optimized_size
        percent = (savings / original_size) * 100
        
        print(f"   Original size: {original_size:,} bytes")
        print(f"   Optimized size: {optimized_size:,} bytes")
        print(f"   Saved: {savings:,} bytes ({percent:.1f}%)")
        
        # Remove backup
        backup_file.unlink()
        
    except Exception as e:
        print(f"⚠️  Warning: Optimization failed: {e}")
        print("   Restoring original file...")
        shutil.move(backup_file, wasm_file)


def main():
    parser = argparse.ArgumentParser(
        description="Build VB6Parse WASM module for playground"
    )
    parser.add_argument(
        "--optimize",
        action="store_true",
        help="Run wasm-opt optimization (requires wasm-opt to be installed)"
    )
    parser.add_argument(
        "--no-typescript",
        action="store_true",
        help="Skip TypeScript definition generation"
    )
    args = parser.parse_args()
    
    # Determine project root (parent of scripts directory)
    script_dir = Path(__file__).parent.resolve()
    project_root = script_dir.parent
    output_dir = project_root / "docs" / "assets" / "wasm"
    
    print("=" * 60)
    print("VB6Parse WASM Build Script")
    print("=" * 60)
    print(f"Platform: {platform.system()} {platform.machine()}")
    print(f"Python: {sys.version.split()[0]}")
    print(f"Project root: {project_root}")
    print(f"Output directory: {output_dir}")
    print("=" * 60)
    
    # Change to project root
    os.chdir(project_root)
    
    # Check requirements
    wasm_pack, wasm_opt = check_requirements()
    print(f"✅ wasm-pack found: {wasm_pack}")
    if wasm_opt:
        print(f"✅ wasm-opt found: {wasm_opt}")
    
    # Ensure output directory exists
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Build WASM
    build_wasm(wasm_pack, output_dir, args.no_typescript)
    
    # Optimize if requested and tool is available
    if args.optimize:
        wasm_file = output_dir / "vb6parse_bg.wasm"
        optimize_wasm(wasm_opt, wasm_file)
    
    # Remove .gitignore created by wasm-pack (we want to commit these files)
    gitignore_file = output_dir / ".gitignore"
    if gitignore_file.exists():
        gitignore_file.unlink()
        print("🗑️  Removed .gitignore from output directory")
    
    print("=" * 60)
    print("✅ WASM build complete!")
    print(f"📦 Output files in: {output_dir}")
    
    # List generated files
    if output_dir.exists():
        files = sorted(output_dir.iterdir())
        if files:
            print("\n📄 Generated files:")
            for file in files:
                size = file.stat().st_size
                print(f"   - {file.name} ({size:,} bytes)")
    
    print("=" * 60)


if __name__ == "__main__":
    main()