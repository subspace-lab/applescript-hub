# Verification

This directory contains tooling and notes for systematically verifying AppleScript snippets.

## Philosophy

AppleScript snippets found online are frequently broken, outdated, or version-specific. The only reliable way to validate a snippet is to **run it**. This directory supports automated and AI-agent-driven verification workflows.

## Snippet Status

Each snippet carries a `-- verified:` metadata header:

| Value | Meaning |
|-------|---------|
| `yes` | Tested manually, confirmed working |
| `no` | Not tested, or known to be broken |
| `partial` | Works on specific macOS versions or configurations |

## Verification Workflow (Manual)

1. Open Terminal
2. Run: `osascript snippets/category/script.applescript`
3. Observe output or side effects
4. Update the `-- verified:` and `-- macos:` headers accordingly

## CI/CD Integration (Planned)

The goal is to automatically run and validate snippets on real Mac hardware via CI.

Planned pipeline:
- Trigger on PR or push to `snippets/`
- Run each modified snippet via `osascript`
- Report pass/fail and update verification status
- Flag scripts that require user interaction or elevated permissions (skip in CI)

## Scripts Requiring Special Handling

Some scripts cannot be run headlessly in CI:

- Scripts using `display dialog` or `display alert`
- Scripts requiring Accessibility permissions (`System Events` UI scripting)
- Scripts that modify files or system state destructively
- Scripts targeting apps that may not be installed

These are marked with `-- ci: skip` in their headers.

## AI Agent Notes

When using this repo as a reference:
- Trust `verified: yes` snippets as ground truth
- Treat `verified: no` snippets as drafts — test before use
- Use `sdef` to cross-check whether a command/property actually exists in the target app's dictionary
- The `app-dictionaries/` directory contains pre-extracted, browsable dictionary docs
