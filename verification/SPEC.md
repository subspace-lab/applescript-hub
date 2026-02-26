# Verification Framework Spec

## Snippet Metadata Format

Replace the old `verified: yes/no` with a structured `[tested]` block:

```applescript
-- title: Click a Menu Item
-- app: System Events
-- description: Reusable handler to click any menu item via UI scripting
-- source: macosxautomation.com
-- tags: ui-scripting, rpa, handler
--
-- [tested]
-- 15: pass @ 2026-02-26 by bot
-- 14: pass @ 2026-01-15 by @someuser
-- 13: fail @ 2026-01-15 by @someuser # menu structure changed
-- 12: untested
```

### Fields

| Field | Required | Description |
|-------|----------|-------------|
| `title` | yes | Short descriptive name |
| `app` | yes | Primary app (e.g. `Finder`, `System Events`, `multiple`) |
| `description` | yes | One-line description |
| `source` | yes | Attribution: `original`, `macosxautomation.com`, `github.com/user/repo`, `app dictionary`, etc. |
| `tags` | no | Comma-separated: `ui-scripting`, `rpa`, `file-management`, `basics`, `handler`, `cross-app` |
| `[tested]` | no | Version test results block (see below) |

### `[tested]` Block Format

Each line: `-- <major_version>: <status> @ <YYYY-MM-DD> by <who> # <optional note>`

| Status | Meaning |
|--------|---------|
| `pass` | Runs without error, produces expected result |
| `fail` | Throws error or produces wrong result |
| `skip` | Cannot test (app not installed, requires interaction, etc.) |
| `untested` | No one has tested this version yet (no `@` or `by` needed) |

- `<who>`: GitHub username (with `@`), or `bot` for automated runs
- `# note`: optional, especially useful for `fail` and `skip` (explain why)
- Lines are ordered newest version first
- Only major versions (15, 14, 13...), not minor

## Tools

### `verification/run_tests.py`

Batch runner that:
1. Finds all `.applescript` files in `snippets/`
2. Parses the metadata header
3. Runs each with `osascript` (with timeout, e.g. 10s)
4. Records pass/fail/skip per snippet
5. **Updates the `[tested]` block** in each snippet file with current macOS version + date + `by bot`
6. Prints summary report to stdout
7. Generates `verification/results-<version>.json` for machine consumption

```
Usage:
  uv run verification/run_tests.py                    # test all
  uv run verification/run_tests.py snippets/finder/   # test one dir
  uv run verification/run_tests.py --dry-run           # parse only, don't run
  uv run verification/run_tests.py --skip-interactive   # skip snippets tagged 'interactive'
```

**Skip logic:** Some snippets can't be auto-tested (send-message, UI scripting that needs a specific app open). Use `-- skip-autotest: true` in metadata to mark these. The runner records them as `skip` with note `requires interaction`.

### `verification/gen_compat_matrix.py`

Reads all snippet `[tested]` blocks and generates:
1. A markdown compatibility table for README or `COMPATIBILITY.md`
2. Summary stats (X tested, Y passing, Z failing on version N)

```
| Snippet | 15 | 14 | 13 |
|---------|----|----|-----|
| finder/get-selected-files | ✅ | ✅ | ✅ |
| system-events/click-menu-item | ✅ | ✅ | ❌ |
| messages/send-message | ⏭️ | — | — |
```

### `verification/results-<version>.json`

Machine-readable output per test run:

```json
{
  "macos_version": 15,
  "date": "2026-02-26",
  "by": "bot",
  "results": [
    {
      "snippet": "finder/get-selected-files.applescript",
      "status": "pass",
      "duration_ms": 312,
      "output": "file1.txt, file2.pdf"
    },
    {
      "snippet": "system-events/click-menu-item.applescript",
      "status": "skip",
      "reason": "requires interaction"
    }
  ]
}
```

## Migration

Existing snippets use `-- verified: yes/no`. The runner should:
1. Parse old format gracefully
2. On first run, convert `verified: yes` → add `[tested]` block with current version as `pass`
3. Convert `verified: no` → add `[tested]` block with current version as `untested`
4. Remove the old `-- verified:` line

## Contributor Workflow

1. `uv run verification/run_tests.py` on their Mac
2. Script auto-updates `[tested]` blocks with their macOS version
3. `git diff` shows exactly what changed
4. Submit PR — reviewer sees which version was tested

## Extra Metadata Field

```applescript
-- skip-autotest: true
```

For snippets that:
- Send messages / emails
- Require specific app state (Music playing, Photos open)
- Need user interaction (dialogs, file pickers)
- Modify system state (volume, dark mode)
