# AppleScript Hub

A community-maintained documentation hub for AppleScript — because Apple's official docs are sparse, scattered, and increasingly neglected as attention shifts to Shortcuts.

This repo serves as a reliable reference for **humans and AI agents** alike: verified snippets, extracted app dictionaries, how-to guides, and curated external resources.

## Why This Exists

- Apple's official AppleScript documentation is incomplete and outdated
- Apple has shifted focus toward Shortcuts, leaving AppleScript under-documented
- Many snippets found online are broken or untested
- AI coding assistants often hallucinate AppleScript syntax due to poor training data

This repo aims to be a ground-truth reference by prioritizing **verified, runnable scripts**.

## Repository Structure

```
applescript-hub/
├── snippets/               # Verified AppleScript snippets by app/category
│   ├── finder/
│   ├── safari/
│   ├── mail/
│   ├── calendar/
│   └── system/             # System Events, UI scripting
├── guides/                 # How-to guides
├── resources/              # Curated links to official and reputable sources
├── app-dictionaries/       # Extracted app dictionary documentation (SDEF)
└── verification/           # Notes and scripts for verifying snippet correctness
```

## Snippet Format

Each snippet file includes a metadata header:

```applescript
-- title: Open a file in Finder
-- verified: yes
-- macos: 13+
-- app: Finder
-- description: Opens a specific file using Finder

tell application "Finder"
    open POSIX file "/Users/username/Documents/example.txt"
end tell
```

## Discovering What's Scriptable

Every scriptable app exposes a dictionary. Two ways to browse it:

**Script Editor GUI:**
Open Script Editor → File → Open Dictionary → pick any app

**Terminal (SDEF):**
```bash
sdef /Applications/Safari.app | sdp -fh --basename Safari
# or just dump as XML:
sdef /Applications/Finder.app > finder.sdef
```

**osascript one-liner:**
```bash
osascript -e 'tell application "Finder" to get name of every window'
```

## Verification Approach

Scripts in this repo are marked with their verification status. The `verification/` directory contains tooling to systematically run and validate snippets — designed to support automated CI and AI-agent verification workflows.

| Status | Meaning |
|--------|---------|
| `verified: yes` | Manually tested and confirmed working |
| `verified: no` | Untested or known broken |
| `verified: partial` | Works on specific macOS versions |

## Contributing

1. Add a snippet following the metadata header format
2. Test it with `osascript your_script.applescript`
3. Mark verification status honestly
4. Submit a PR

## Resources

See [`resources/00-curated-links.md`](resources/00-curated-links.md) for official Apple docs, reputable community sources, and tools.

---

> AppleScript isn't dead — it's just under-documented. This repo fixes that.
