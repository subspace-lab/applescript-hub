# Discovering App Dictionaries

Every scriptable app ships with a **dictionary** that describes exactly what commands and objects it supports. This is your primary reference — more reliable than any website.

## Method 1: Script Editor (GUI)

1. Open **Script Editor** (`/Applications/Utilities/Script Editor.app`)
2. Go to **File → Open Dictionary** (⌘Shift O)
3. Pick any installed app
4. Browse its classes, commands, and properties

This is the best starting point when exploring an unfamiliar app.

## Method 2: sdef (Terminal)

`sdef` dumps an app's dictionary as XML (SDEF format):

```bash
# Dump to stdout
sdef /Applications/Safari.app

# Save to file for processing
sdef /Applications/Finder.app > finder.sdef

# Pretty-print with xmllint
sdef /Applications/Mail.app | xmllint --format - | less
```

## Method 3: List All Scriptable Apps

```applescript
tell application "System Events"
    set scriptableApps to name of every application process whose has scripting is true
    return scriptableApps
end tell
```

Or from Terminal:
```bash
osascript -e 'tell application "System Events" to get name of every application process whose has scripting is true'
```

## Method 4: Check if an App is Scriptable

```bash
# Returns dictionary XML if scriptable, error if not
sdef /Applications/AppName.app 2>/dev/null && echo "Scriptable" || echo "Not scriptable"
```

## Reading the Dictionary

A dictionary contains:

| Element | Meaning |
|---------|---------|
| **Suite** | A group of related commands/classes (e.g. "Standard Suite") |
| **Command** | An action you can tell the app to do (e.g. `open`, `close`, `quit`) |
| **Class** | An object type (e.g. `window`, `document`, `tab`) |
| **Property** | An attribute of a class (e.g. `name`, `url`, `bounds`) |
| **Element** | A contained object (e.g. a `window` contains `tabs`) |

## Example: Reading Safari's Dictionary

```bash
sdef /Applications/Safari.app | grep -A5 '<command name='
```

Then try it:
```applescript
tell application "Safari"
    get URL of current tab of front window
end tell
```

## Notes

- Not all apps are scriptable — many modern apps ignore AppleScript entirely
- Some apps only expose a minimal "Standard Suite" (quit, activate, etc.)
- App Store sandboxed apps may have restricted scripting access
- Dictionaries don't change unless the app is updated — safe to cache
