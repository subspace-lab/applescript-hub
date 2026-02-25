# Shortcuts AppleScript Dictionary

> Auto-generated from `Shortcuts.sdef` inside the app bundle.  
> Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Shortcuts"`

## Table of Contents

- [Shortcuts Suite](#shortcuts-suite)

---

## Shortcuts Suite

Classes and Commands for working with Shortcuts

### Commands

### `run`

Run a shortcut. To run a shortcut in the background, without opening the Shortcuts app, tell 'Shortcuts Events' instead of 'Shortcuts'.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shortcut` | no | the shortcut to run |
| `with input` | `any` | yes | the input to provide to the shortcut |

**Returns:** `any` — the result of the shortcut

### Classes

### `shortcut`

a shortcut in the Shortcuts application

*Plural:* `shortcuts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | the name of the shortcut |
| `subtitle` | `text` | r/o | the shortcut's subtitle |
| `id` | `text` | r/o | the unique identifier of the shortcut |
| `folder` | `folder` | r/w | the folder containing this shortcut |
| `color` | `RGB color` | r/o | the shortcut's color |
| `icon` | `TIFF image` | r/o | the shortcut's icon |
| `accepts input` | `boolean` | r/o | indicates whether or not the shortcut accepts input data |
| `action count` | `integer` | r/o | the number of actions in the shortcut |

**Responds to:** `run`

### `folder`

a folder containing shortcuts

*Plural:* `folders`

**Contains:** `shortcut`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | the name of the folder |
| `id` | `text` | r/o | the unique identifier of the folder |
