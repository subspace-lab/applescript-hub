# Systempreferences AppleScript Dictionary

> Auto-generated from `SystemPreferences.sdef` inside the app bundle.  
> Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Systempreferences"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [System Settings](#system-settings)

---

## Standard Suite

---

## System Settings

Classes and Commands specific to System Settings

### Commands

### `reveal`

Reveals a settings pane or an anchor within a pane.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pane | anchor` | no |  |

**Returns:** `pane | anchor`

### `authorize`

Prompt for authorization for a settings pane. Deprecated: no longer does anything.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pane` | no |  |

**Returns:** `pane`

### `timedLoad`

Times and loads given settings pane and returns load time. Deprecated: no longer does anything.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pane` | no |  |

**Returns:** `real`

### Classes

### `pane`

A settings pane.

**Contains:** `anchor`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The id of the settings pane. |
| `name` | `text` | r/o | The name of the settings pane. |

**Responds to:** `reveal`, `authorize`

### `anchor`

An anchor within a settings pane.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the anchor. |

**Responds to:** `reveal`
