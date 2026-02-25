# Reminders AppleScript Dictionary

> Auto-generated from `Reminders.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Reminders"`

## Table of Contents

- [Reminders Suite](#reminders-suite)

---

## Reminders Suite

Terms and Events for controlling the Reminders application

### Commands

### `show`

Show an object in the Reminders UI

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list | reminder` | no | The object to be shown |

**Returns:** `list | reminder` — The object shown

### Classes

### `account`

An account in the Reminders application

**Contains:** `list`, `reminder`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The unique identifier of the account |
| `name` | `text` | r/o | The name of the account |

**Responds to:** `show`

### `list`

A list in the Reminders application

*Plural:* `lists`

**Contains:** `reminder`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The unique identifier of the list |
| `name` | `text` | r/w | The name of the list |
| `container` | `account | list` | r/o | The container of the list |
| `color` | `text` | r/w | The color of the list |
| `emblem` | `text` | r/w | The emblem icon name of the list |

**Responds to:** `show`

### `reminder`

A reminder in the Reminders application

*Plural:* `reminders`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | The name of the reminder |
| `id` | `text` | r/o | The unique identifier of the reminder |
| `container` | `list | reminder` | r/o | The container of the reminder |
| `creation date` | `date` | r/o | The creation date of the reminder |
| `modification date` | `date` | r/o | The modification date of the reminder |
| `body` | `text` | r/w | The notes attached to the reminder |
| `completed` | `boolean` | r/w | Whether the reminder is completed |
| `completion date` | `date` | r/w | The completion date of the reminder |
| `due date` | `date` | r/w | The due date of the reminder; will set both date and time |
| `allday due date` | `date` | r/w | The all-day due date of the reminder; will only set a date |
| `remind me date` | `date` | r/w | The remind date of the reminder |
| `priority` | `integer` | r/w | The priority of the reminder; 0: no priority, 1–4: high, 5: medium, 6–9: low |
| `flagged` | `boolean` | r/w | Whether the reminder is flagged |

**Responds to:** `show`

### Enumerations

### `saveable file format`

| Value | Description |
|-------|-------------|
| `text` | Text File Format |
