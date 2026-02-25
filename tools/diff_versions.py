#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = []
# ///
"""
diff_versions.py — Compare app dictionaries across macOS versions.

Parses the markdown structure (suites, commands, classes, properties, enums)
and produces a semantic changelog rather than a line-level diff.

Usage:
    # Compare two directories
    uv run tools/diff_versions.py old-dictionaries/ new-dictionaries/ -o changelogs/14-to-15.md

    # Compare git refs (checks out app-dictionaries/ from each ref)
    uv run tools/diff_versions.py --git-old v14 --git-new v15 -o changelogs/14-to-15.md
"""

import argparse
import re
import subprocess
import sys
import tempfile
from pathlib import Path


# ── Markdown parser ──────────────────────────────────────────────────────────

def parse_dictionary(md_text: str) -> dict:
    """Parse a dictionary markdown file into a structured dict.

    Returns: {suite_name: {commands: set, classes: {name: {props: set}}, enums: set}}
    """
    suites = {}
    current_suite = None
    current_section = None  # "commands", "classes", "enums"
    current_class = None

    for line in md_text.splitlines():
        # Suite heading: ## Suite Name
        if m := re.match(r"^## (.+)$", line):
            name = m.group(1).strip()
            if name == "Table of Contents":
                continue
            current_suite = name
            suites[current_suite] = {"commands": set(), "classes": {}, "enums": set()}
            current_section = None
            current_class = None
            continue

        if current_suite is None:
            continue

        suite = suites[current_suite]

        # Section heading: ### Commands / ### Classes / ### Enumerations
        if line.strip() == "### Commands":
            current_section = "commands"
            current_class = None
            continue
        if line.strip() == "### Classes":
            current_section = "classes"
            current_class = None
            continue
        if line.strip() == "### Enumerations":
            current_section = "enums"
            current_class = None
            continue

        # Item heading: ### `name`
        if m := re.match(r"^### `(.+)`$", line):
            item_name = m.group(1)
            if current_section == "commands":
                suite["commands"].add(item_name)
            elif current_section == "classes":
                current_class = item_name
                suite["classes"][item_name] = {"props": set()}
            elif current_section == "enums":
                suite["enums"].add(item_name)
            continue

        # Property row inside a class: | `prop_name` | ...
        if current_section == "classes" and current_class and line.startswith("| `"):
            if m := re.match(r"^\| `([^`]+)`", line):
                suite["classes"][current_class]["props"].add(m.group(1))

    return suites


def load_dictionaries(directory: Path) -> dict[str, dict]:
    """Load all .md dictionaries from a directory. Returns {filename: parsed_dict}."""
    result = {}
    for md_file in sorted(directory.glob("*.md")):
        if md_file.name == "README.md":
            continue
        result[md_file.name] = parse_dictionary(md_file.read_text())
    return result


def load_from_git_ref(ref: str, subdir: str = "app-dictionaries") -> dict[str, dict]:
    """Load dictionaries from a git ref by checking out the directory."""
    with tempfile.TemporaryDirectory() as tmpdir:
        # List files in the ref
        result = subprocess.run(
            ["git", "ls-tree", "--name-only", f"{ref}:{subdir}"],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            print(f"Error: could not list {subdir}/ in git ref '{ref}'", file=sys.stderr)
            sys.exit(1)

        dicts = {}
        for filename in result.stdout.strip().splitlines():
            if not filename.endswith(".md") or filename == "README.md":
                continue
            show = subprocess.run(
                ["git", "show", f"{ref}:{subdir}/{filename}"],
                capture_output=True, text=True,
            )
            if show.returncode == 0:
                dicts[filename] = parse_dictionary(show.stdout)
        return dicts


# ── Diff logic ───────────────────────────────────────────────────────────────

def diff_classes(old_classes: dict, new_classes: dict) -> list[str]:
    """Diff class definitions, returning changelog lines."""
    lines = []
    old_names = set(old_classes)
    new_names = set(new_classes)

    for name in sorted(new_names - old_names):
        lines.append(f"  - Added class `{name}`")
    for name in sorted(old_names - new_names):
        lines.append(f"  - Removed class `{name}`")

    for name in sorted(old_names & new_names):
        old_props = old_classes[name]["props"]
        new_props = new_classes[name]["props"]
        for p in sorted(new_props - old_props):
            lines.append(f"  - `{name}`: added property `{p}`")
        for p in sorted(old_props - new_props):
            lines.append(f"  - `{name}`: removed property `{p}`")

    return lines


def diff_app(old_suites: dict, new_suites: dict) -> list[str]:
    """Diff two parsed dictionaries. Returns changelog lines."""
    lines = []
    all_suites = sorted(set(old_suites) | set(new_suites))

    for suite_name in all_suites:
        if suite_name not in old_suites:
            lines.append(f"- **{suite_name}**: new suite")
            continue
        if suite_name not in new_suites:
            lines.append(f"- **{suite_name}**: removed")
            continue

        old_s = old_suites[suite_name]
        new_s = new_suites[suite_name]
        suite_lines = []

        # Commands
        added_cmds = new_s["commands"] - old_s["commands"]
        removed_cmds = old_s["commands"] - new_s["commands"]
        for c in sorted(added_cmds):
            suite_lines.append(f"  - Added command `{c}`")
        for c in sorted(removed_cmds):
            suite_lines.append(f"  - Removed command `{c}`")

        # Classes
        suite_lines.extend(diff_classes(old_s["classes"], new_s["classes"]))

        # Enums
        added_enums = new_s["enums"] - old_s["enums"]
        removed_enums = old_s["enums"] - new_s["enums"]
        for e in sorted(added_enums):
            suite_lines.append(f"  - Added enum `{e}`")
        for e in sorted(removed_enums):
            suite_lines.append(f"  - Removed enum `{e}`")

        if suite_lines:
            lines.append(f"- **{suite_name}**:")
            lines.extend(suite_lines)

    return lines


# ── Changelog renderer ───────────────────────────────────────────────────────

def render_changelog(old_label: str, new_label: str,
                     old_dicts: dict, new_dicts: dict) -> str:
    """Render a full changelog markdown document."""
    lines = [
        f"# Dictionary Changes: {old_label} to {new_label}",
        "",
    ]

    all_apps = sorted(set(old_dicts) | set(new_dicts))

    added_apps = sorted(set(new_dicts) - set(old_dicts))
    removed_apps = sorted(set(old_dicts) - set(new_dicts))
    common_apps = sorted(set(old_dicts) & set(new_dicts))

    if added_apps:
        lines.append("## New Apps")
        lines.append("")
        for app in added_apps:
            lines.append(f"- {app.removesuffix('.md')}")
        lines.append("")

    if removed_apps:
        lines.append("## Removed Apps")
        lines.append("")
        for app in removed_apps:
            lines.append(f"- {app.removesuffix('.md')}")
        lines.append("")

    changes_found = False
    change_lines = ["## Changed Apps", ""]

    for app in common_apps:
        app_changes = diff_app(old_dicts[app], new_dicts[app])
        if app_changes:
            changes_found = True
            display = app.removesuffix(".md")
            change_lines.append(f"### {display}")
            change_lines.append("")
            change_lines.extend(app_changes)
            change_lines.append("")

    if changes_found:
        lines.extend(change_lines)

    if not added_apps and not removed_apps and not changes_found:
        lines.append("No changes detected.")
        lines.append("")

    return "\n".join(lines)


# ── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Compare app dictionaries across versions.")
    parser.add_argument("old_dir", nargs="?", help="Path to old dictionaries directory")
    parser.add_argument("new_dir", nargs="?", help="Path to new dictionaries directory")
    parser.add_argument("--git-old", help="Git ref for old version")
    parser.add_argument("--git-new", help="Git ref for new version")
    parser.add_argument("-o", "--output", help="Output changelog file (default: stdout)")
    args = parser.parse_args()

    # Determine source of old/new dictionaries
    if args.git_old and args.git_new:
        old_label = args.git_old
        new_label = args.git_new
        old_dicts = load_from_git_ref(args.git_old)
        new_dicts = load_from_git_ref(args.git_new)
    elif args.old_dir and args.new_dir:
        old_label = args.old_dir
        new_label = args.new_dir
        old_dicts = load_dictionaries(Path(args.old_dir))
        new_dicts = load_dictionaries(Path(args.new_dir))
    else:
        parser.error("Provide either two directories or --git-old and --git-new")
        return

    changelog = render_changelog(old_label, new_label, old_dicts, new_dicts)

    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(changelog)
        print(f"Changelog written to {out}", file=sys.stderr)
    else:
        print(changelog)


if __name__ == "__main__":
    main()
