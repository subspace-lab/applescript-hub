# Contributing

AppleScript Hub aims to be a ground-truth reference for AppleScript — for humans and AI agents alike. Contributions are welcome!

## App Dictionaries

The `app-dictionaries/` folder contains auto-generated markdown docs from macOS app SDEF files. Only **major macOS versions** matter (e.g. 14, 15) since dictionaries rarely change in point releases.

### Regenerating all dictionaries

```bash
uv run tools/extract_all.py
```

This scans `/Applications`, `/System/Applications`, and `/System/Library/CoreServices` for scriptable apps, auto-detects your macOS version, and generates markdown for each.

### Adding a single app

```bash
uv run tools/sdef_to_md.py "AppName" --out app-dictionaries/appname.md
# or from a .sdef path:
uv run tools/sdef_to_md.py --path /path/to/app.sdef --out app-dictionaries/appname.md
```

### Comparing versions

```bash
uv run tools/diff_versions.py old-dicts/ new-dicts/ -o changelogs/14-to-15.md
```

### Submitting dictionary PRs

1. Run `uv run tools/extract_all.py`
2. Check `git diff` for meaningful changes (ignore whitespace-only diffs)
3. Submit a PR noting your macOS version

## Snippets

Add verified scripts to `snippets/<app>/`. Each file needs a metadata header:

```applescript
-- title: Short description
-- verified: yes
-- macos: 15+
-- app: AppName
-- description: What this script does

tell application "AppName"
    -- your script here
end tell
```

Test with `osascript your_script.applescript` before submitting. Mark `verified: no` if you can't test it.

## General Guidelines

- Keep PRs focused — one app dictionary or a few related snippets
- Don't edit auto-generated dictionary files by hand
- Test snippets on your macOS version and note it in the header
