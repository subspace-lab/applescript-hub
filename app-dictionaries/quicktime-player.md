# Quicktimeplayerx AppleScript Dictionary

> Auto-generated from `QuickTimePlayerX.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Quicktimeplayerx"`

## Table of Contents

- [Internet Suite](#internet-suite)
- [QuickTime Player Suite](#quicktime-player-suite)

---

## Internet Suite

Common URL related functionality

### Commands

### `open URL`

Open a URL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | the URL |

---

## QuickTime Player Suite

Classes and Commands for working with QuickTime Player

### Commands

### `play`

Play the movie.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the movie to play |

### `start`

Start the movie recording.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the recording to start |

### `pause`

Pause the recording.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the recording to pause |

### `resume`

Resume the recording.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the recording to resume |

### `stop`

Stop the movie or recording.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the movie or recording to stop |

### `step backward`

Step the movie backward the specified number of steps (default is 1).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the movie to step |
| `by` | `integer` | yes | number of steps |

### `step forward`

Step the movie forward the specified number of steps (default is 1).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the movie to step |
| `by` | `integer` | yes | number of steps |

### `trim`

Trim the movie.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the movie to trim |
| `from` | `real` | no | start time in seconds |
| `to` | `real` | no | end time in seconds |

### `present`

Present the document full screen.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the document to present |

### `new movie recording`

Create a new movie recording document.

**Returns:** `document` — The new movie recording document.

### `new audio recording`

Create a new audio recording document.

**Returns:** `document` — The new audio recording document.

### `new screen recording`

Create a new screen recording document.

### `export`

Export a movie to another file

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the movie to export |
| `in` | `file` | no | the destination file |
| `using settings preset` | `text` | no | the name of the export settings preset to use |

### `show remote hud`

Show the document's Remote HUD

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### Classes

### `video recording device`

A video recording device

*Plural:* `video recording devices`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the device. |
| `id` | `text` | r/o | The unique identifier of the device. |

### `audio recording device`

An audio recording device

*Plural:* `audio recording devices`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the device. |
| `id` | `text` | r/o | The unique identifier of the device. |

### `audio compression preset`

An audio recording compression preset

*Plural:* `audio compression presets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the preset. |
| `id` | `text` | r/o | The unique identifier of the preset. |

### `movie compression preset`

A movie recording compression preset

*Plural:* `movie compression presets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the preset. |
| `id` | `text` | r/o | The unique identifier of the preset. |

### `screen compression preset`

A screen recording compression preset

*Plural:* `screen compression presets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the preset. |
| `id` | `text` | r/o | The unique identifier of the preset. |
