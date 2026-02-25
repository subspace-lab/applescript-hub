# Curated Resources

## Official Apple Documentation

| Resource | URL | Notes |
|----------|-----|-------|
| AppleScript Language Guide | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/ | Archived, but still the most complete official reference |
| AppleScript Overview | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptX/AppleScriptX.html | High-level introduction |
| Technical Note TN2065: do shell script | https://developer.apple.com/library/archive/technotes/tn2065/_index.html | Shell integration patterns |
| Scripting Interface Guidelines | https://developer.apple.com/library/archive/documentation/Cocoa/Conceptual/ScriptableCocoaApplications/ | For developers making apps scriptable |
| Script Editor User Guide | https://support.apple.com/guide/script-editor/ | Beginner-friendly Apple support doc |

## App Dictionary Access

| Method | Command/Location |
|--------|-----------------|
| Script Editor GUI | File → Open Dictionary |
| SDEF dump (XML) | `sdef /Applications/AppName.app` |
| List all scriptable apps | `osascript -e 'tell application "System Events" to get name of every application process'` |

## Community & Reference Sites

| Resource | URL | Notes |
|----------|-----|-------|
| MacScripter Forums | https://macscripter.net | Longest-running AppleScript community |
| Late Night Software (Script Debugger) | https://latenightsw.com | Best IDE for AppleScript; excellent blog |
| Sal Soghoian's Resources | https://macosxautomation.com | Former Apple automation product manager |
| AppleScript at Stack Overflow | https://stackoverflow.com/questions/tagged/applescript | Q&A, many verified answers |
| Keyboard Maestro Wiki | https://wiki.keyboardmaestro.com/AppleScript | AppleScript integration with KM |

## Books

| Title | Author | Notes |
|-------|--------|-------|
| *AppleScript: The Definitive Guide* | Matt Neuburg | O'Reilly; older but comprehensive |
| *AppleScript 1-2-3* | Sal Soghoian & Bill Cheeseman | Beginner-friendly, practical |

## Tools

| Tool | Purpose |
|------|---------|
| Script Debugger (Late Night Software) | Full IDE with dictionary browser, step debugger |
| Script Editor (built-in) | Basic editor, good dictionary browser |
| `osascript` | Run scripts from Terminal |
| `sdef` | Extract app dictionaries as SDEF XML |
| `sdp` | Convert SDEF to header files |
| Alfred / Keyboard Maestro | Workflow tools with AppleScript integration |
