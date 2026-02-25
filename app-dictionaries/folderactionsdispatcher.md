# Folderactionsdispatcher AppleScript Dictionary

> Auto-generated from `FolderActionsDispatcher.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Folderactionsdispatcher"`

## Table of Contents

- [Hidden Suite](#hidden-suite)

---

## Hidden Suite

Hidden Terms and Events for controlling the Folder Actions Dispatcher application

### Commands

### `do folder action`

Send a folder action code to a folder action script

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file` | no | the path to the folder |
| `folder action code` | `actn` | no | the folder action message to process |
| `with item list` | `text` | yes | a list of items for the folder action message to process |
| `with window size` | `rectangle` | yes | the new window size for the folder action message to process |

### Enumerations

### `actn`

| Value | Description |
|-------|-------------|
| `items added` | items added |
| `items removed` | items removed |
| `window closed` | window closed |
| `window moved` | window moved |
| `window opened` | window opened |
