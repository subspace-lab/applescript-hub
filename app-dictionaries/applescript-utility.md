# Applescriptutility AppleScript Dictionary

> Auto-generated from `AppleScriptUtility.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Applescriptutility"`

## Table of Contents

- [AppleScript Utility Suite](#applescript-utility-suite)

---

## AppleScript Utility Suite

Terms and Events for controlling the AppleScript Utility application

### Classes

### `application`

the AppleScript Utility application

*Plural:* `applications`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `default script editor` | `file` | r/w | the editor to be used to open scripts |
| `GUI Scripting enabled` | `boolean` | r/o | Are GUI Scripting events currently being processed? |
| `application scripts position` | `ApplicationScriptsPositions` | r/w | the position in the Script menu at which the application scripts are displayed |
| `Script menu enabled` | `boolean` | r/w | Is the Script menu installed in the menu bar? |
| `show Computer scripts` | `boolean` | r/w | Are the Computer scripts shown in the Script menu? |
| `UI elements enabled` | `boolean` | r/o | Are UI element events currently being processed? |

### Enumerations

### `ApplicationScriptsPositions`

| Value | Description |
|-------|-------------|
| `top` | top |
| `bottom` | bottom |
