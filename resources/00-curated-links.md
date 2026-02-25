# Curated Resources

## Official Apple Documentation

### AppleScript Language Guide
Archived but the most complete official reference. Direct chapter links:

| Chapter | URL |
|---------|-----|
| Introduction | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/introduction/ASLR_intro.html |
| Lexical Conventions | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/conceptual/ASLR_lexical_conventions.html |
| Fundamentals | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/conceptual/ASLR_fundamentals.html |
| Variables and Properties | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/conceptual/ASLR_variables.html |
| Script Objects | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/conceptual/ASLR_script_objects.html |
| About Handlers | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/conceptual/ASLR_about_handlers.html |
| Class Reference | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_classes.html |
| Commands Reference | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_cmds.html |
| Reference Forms | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_reference_forms.html |
| Operators Reference | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_operators.html |
| Control Statements | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_control_statements.html |
| Handler Reference | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_handlers.html |
| Folder Actions Reference | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_folder_actions.html |
| Keywords | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_keywords.html |
| Error Codes | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_error_codes.html |
| Working with Errors | https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_error_xmpls.html |

### Other Official Docs

| Resource | URL | Notes |
|----------|-----|-------|
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
