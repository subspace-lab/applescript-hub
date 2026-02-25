#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = []
# ///
"""
sdef_to_md.py — Generate markdown documentation from an app's SDEF dictionary.

Usage:
    uv run tools/sdef_to_md.py "Finder"
    uv run tools/sdef_to_md.py "Mail" --out app-dictionaries/mail.md
    uv run tools/sdef_to_md.py --path /path/to/custom.sdef --out output.md

The SDEF is read directly from the app bundle — never committed to the repo.
"""

import argparse
import subprocess
import sys
from pathlib import Path
from xml.etree import ElementTree as ET


# ── SDEF location helpers ────────────────────────────────────────────────────

def find_app_path(app_name: str) -> Path:
    """Resolve app bundle path via AppleScript (handles Cryptex/SIP locations)."""
    result = subprocess.run(
        ["osascript", "-e", f'POSIX path of (path to application "{app_name}")'],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise FileNotFoundError(f"Could not locate app: {app_name}")
    return Path(result.stdout.strip())


def find_sdef_in_bundle(app_path: Path) -> Path:
    """Find the .sdef file inside an app bundle."""
    matches = list(app_path.rglob("*.sdef"))
    if not matches:
        raise FileNotFoundError(f"No .sdef found in {app_path}")
    # Prefer the one directly in Resources/
    for m in matches:
        if m.parent.name == "Resources":
            return m
    return matches[0]


# ── SDEF parser ──────────────────────────────────────────────────────────────

NS = {"cocoa": "http://www.apple.com/DTDs/sdef.dtd"}


def text(el, attr="description") -> str:
    return (el.get(attr) or "").strip()


def parse_type(el) -> str:
    """Extract type, handling list types."""
    t = el.get("type", "")
    if not t:
        # check for <type> child elements (union types)
        types = [c.get("type", "") for c in el.findall("type")]
        t = " | ".join(filter(None, types)) or "any"
    return t


def parse_command(cmd) -> dict:
    params = []
    dp = cmd.find("direct-parameter")
    if dp is not None:
        params.append({
            "name": "(direct)",
            "type": parse_type(dp),
            "optional": dp.get("optional", "no") == "yes",
            "desc": text(dp),
        })
    for p in cmd.findall("parameter"):
        params.append({
            "name": p.get("name", ""),
            "type": parse_type(p),
            "optional": p.get("optional", "no") == "yes",
            "desc": text(p),
        })
    result = cmd.find("result")
    return {
        "name": cmd.get("name", ""),
        "desc": text(cmd),
        "params": params,
        "result": {"type": parse_type(result), "desc": text(result)} if result is not None else None,
    }


def parse_class(cls) -> dict:
    props = []
    for p in cls.findall("property"):
        access = p.get("access", "rw")
        props.append({
            "name": p.get("name", ""),
            "type": parse_type(p),
            "access": {"r": "r/o", "w": "w/o"}.get(access, "r/w"),
            "desc": text(p),
        })
    elements = [e.get("type", "") for e in cls.findall("element")]
    responds = [r.get("command", "") for r in cls.findall("responds-to")]
    return {
        "name": cls.get("name", ""),
        "desc": text(cls),
        "plural": cls.get("plural", ""),
        "props": props,
        "elements": elements,
        "responds": responds,
    }


def parse_enumeration(en) -> dict:
    return {
        "name": en.get("name", ""),
        "desc": text(en),
        "values": [
            {"name": e.get("name", ""), "desc": text(e)}
            for e in en.findall("enumerator")
        ],
    }


def parse_sdef(sdef_path: Path) -> list[dict]:
    tree = ET.parse(sdef_path)
    root = tree.getroot()
    suites = []
    for suite in root.findall("suite"):
        suites.append({
            "name": suite.get("name", ""),
            "desc": text(suite),
            "commands": [parse_command(c) for c in suite.findall("command")],
            "classes": [parse_class(c) for c in suite.findall("class")],
            "enumerations": [parse_enumeration(e) for e in suite.findall("enumeration")],
        })
    return suites


# ── Markdown renderer ─────────────────────────────────────────────────────────

def render_command(cmd: dict) -> str:
    lines = [f"### `{cmd['name']}`", ""]
    if cmd["desc"]:
        lines += [cmd["desc"], ""]
    if cmd["params"]:
        lines.append("**Parameters:**")
        lines.append("")
        lines.append("| Parameter | Type | Optional | Description |")
        lines.append("|-----------|------|----------|-------------|")
        for p in cmd["params"]:
            opt = "yes" if p["optional"] else "no"
            lines.append(f"| `{p['name']}` | `{p['type']}` | {opt} | {p['desc']} |")
        lines.append("")
    if cmd["result"]:
        r = cmd["result"]
        lines.append(f"**Returns:** `{r['type']}`" + (f" — {r['desc']}" if r["desc"] else ""))
        lines.append("")
    return "\n".join(lines)


def render_class(cls: dict) -> str:
    lines = [f"### `{cls['name']}`", ""]
    if cls["desc"]:
        lines += [cls["desc"], ""]
    if cls["plural"]:
        lines += [f"*Plural:* `{cls['plural']}`", ""]
    if cls["elements"]:
        lines.append(f"**Contains:** {', '.join(f'`{e}`' for e in cls['elements'])}")
        lines.append("")
    if cls["props"]:
        lines.append("**Properties:**")
        lines.append("")
        lines.append("| Property | Type | Access | Description |")
        lines.append("|----------|------|--------|-------------|")
        for p in cls["props"]:
            lines.append(f"| `{p['name']}` | `{p['type']}` | {p['access']} | {p['desc']} |")
        lines.append("")
    if cls["responds"]:
        lines.append(f"**Responds to:** {', '.join(f'`{r}`' for r in cls['responds'])}")
        lines.append("")
    return "\n".join(lines)


def render_enumeration(en: dict) -> str:
    lines = [f"### `{en['name']}`", ""]
    if en["desc"]:
        lines += [en["desc"], ""]
    if en["values"]:
        lines.append("| Value | Description |")
        lines.append("|-------|-------------|")
        for v in en["values"]:
            lines.append(f"| `{v['name']}` | {v['desc']} |")
        lines.append("")
    return "\n".join(lines)


def render_markdown(app_name: str, suites: list[dict], sdef_path: Path) -> str:
    lines = [
        f"# {app_name} AppleScript Dictionary",
        "",
        f"> Auto-generated from `{sdef_path.name}` inside the app bundle.  ",
        f"> Do not edit manually — regenerate with `uv run tools/sdef_to_md.py \"{app_name}\"`",
        "",
        "## Table of Contents",
        "",
    ]

    # TOC
    for suite in suites:
        anchor = suite["name"].lower().replace(" ", "-")
        lines.append(f"- [{suite['name']}](#{anchor})")
    lines.append("")

    # Suites
    for suite in suites:
        lines += [f"---", "", f"## {suite['name']}", ""]
        if suite["desc"]:
            lines += [suite["desc"], ""]

        if suite["commands"]:
            lines += ["### Commands", ""]
            for cmd in suite["commands"]:
                lines.append(render_command(cmd))

        if suite["classes"]:
            lines += ["### Classes", ""]
            for cls in suite["classes"]:
                lines.append(render_class(cls))

        if suite["enumerations"]:
            lines += ["### Enumerations", ""]
            for en in suite["enumerations"]:
                lines.append(render_enumeration(en))

    return "\n".join(lines)


# ── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate markdown docs from an app's SDEF.")
    parser.add_argument("app", nargs="?", help="App name, e.g. 'Finder'")
    parser.add_argument("--path", help="Direct path to a .sdef file")
    parser.add_argument("--out", help="Output .md file (default: stdout)")
    args = parser.parse_args()

    if args.path:
        sdef_path = Path(args.path)
        app_name = sdef_path.stem.capitalize()
    elif args.app:
        app_name = args.app
        print(f"Locating {app_name}...", file=sys.stderr)
        app_path = find_app_path(app_name)
        print(f"  Found: {app_path}", file=sys.stderr)
        sdef_path = find_sdef_in_bundle(app_path)
        print(f"  SDEF:  {sdef_path}", file=sys.stderr)
    else:
        parser.print_help()
        sys.exit(1)

    print(f"Parsing SDEF...", file=sys.stderr)
    suites = parse_sdef(sdef_path)
    total_commands = sum(len(s["commands"]) for s in suites)
    total_classes = sum(len(s["classes"]) for s in suites)
    print(f"  {len(suites)} suites, {total_commands} commands, {total_classes} classes", file=sys.stderr)

    md = render_markdown(app_name, suites, sdef_path)

    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(md)
        print(f"Written to {out}", file=sys.stderr)
    else:
        print(md)


if __name__ == "__main__":
    main()
