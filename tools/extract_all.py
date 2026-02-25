#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = []
# ///
"""
extract_all.py — Discover all scriptable apps and generate markdown dictionaries.

Usage:
    uv run tools/extract_all.py
    uv run tools/extract_all.py --out-dir app-dictionaries
"""

import argparse
import platform
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
SDEF_TO_MD = SCRIPT_DIR / "sdef_to_md.py"

SEARCH_DIRS = [
    Path("/Applications"),
    Path("/System/Applications"),
    Path("/System/Library/CoreServices"),
]


def get_macos_version() -> tuple[int, str]:
    """Return (major_version, full_version) e.g. (15, '15.2')."""
    ver = platform.mac_ver()[0]
    major = int(ver.split(".")[0])
    return major, ver


def find_all_sdef_files() -> list[tuple[str, Path]]:
    """Scan standard directories for .app bundles containing .sdef files.

    Returns list of (app_display_name, sdef_path).
    """
    results = []
    seen_names = set()

    for search_dir in SEARCH_DIRS:
        if not search_dir.exists():
            continue
        # Walk through .app bundles (including nested ones like Utilities/)
        for app_bundle in sorted(search_dir.rglob("*.app")):
            # Skip nested .app bundles inside other .app bundles
            parts = app_bundle.relative_to(search_dir).parts
            app_count = sum(1 for p in parts if p.endswith(".app"))
            if app_count > 1:
                continue

            sdef_files = list(app_bundle.rglob("*.sdef"))
            if not sdef_files:
                continue

            # Prefer the one in Resources/
            sdef = sdef_files[0]
            for s in sdef_files:
                if s.parent.name == "Resources":
                    sdef = s
                    break

            app_name = app_bundle.stem
            if app_name in seen_names:
                continue
            seen_names.add(app_name)
            results.append((app_name, sdef))

    return results


def slug(app_name: str) -> str:
    """Convert app name to filename slug: 'QuickTime Player' -> 'quicktime-player'."""
    return app_name.lower().replace(" ", "-")


def inject_macos_version(md_path: Path, macos_ver: str) -> None:
    """Add macOS version to the auto-generated header in the markdown file."""
    text = md_path.read_text()
    text = text.replace(
        "> Do not edit manually",
        f"> macOS {macos_ver}. Do not edit manually",
    )
    md_path.write_text(text)


def main():
    parser = argparse.ArgumentParser(description="Extract all scriptable app dictionaries.")
    parser.add_argument("--out-dir", default="app-dictionaries", help="Output directory")
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    major, full_ver = get_macos_version()
    print(f"macOS {full_ver} (major: {major})")
    print()

    apps = find_all_sdef_files()
    print(f"Found {len(apps)} scriptable apps")
    print()

    generated = []
    failed = []

    for app_name, sdef_path in apps:
        out_file = out_dir / f"{slug(app_name)}.md"
        print(f"  {app_name} ... ", end="", flush=True)

        try:
            result = subprocess.run(
                [sys.executable, str(SDEF_TO_MD), "--path", str(sdef_path), "--out", str(out_file)],
                capture_output=True, text=True, timeout=30,
            )
            if result.returncode != 0:
                print(f"FAILED ({result.stderr.strip().splitlines()[-1] if result.stderr.strip() else 'unknown error'})")
                failed.append((app_name, result.stderr.strip()))
                continue

            inject_macos_version(out_file, full_ver)
            generated.append(app_name)
            print("OK")

        except subprocess.TimeoutExpired:
            print("TIMEOUT")
            failed.append((app_name, "timed out"))
        except Exception as e:
            print(f"ERROR ({e})")
            failed.append((app_name, str(e)))

    # Summary
    print()
    print("=" * 60)
    print(f"  Found:     {len(apps)} scriptable apps")
    print(f"  Generated: {len(generated)}")
    print(f"  Failed:    {len(failed)}")
    print("=" * 60)

    if failed:
        print()
        print("Failed apps:")
        for name, err in failed:
            print(f"  - {name}: {err}")

    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
