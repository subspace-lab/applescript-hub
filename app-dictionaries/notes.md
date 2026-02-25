# Notes AppleScript Dictionary

> Auto-generated from `Notes.sdef` inside the app bundle.  
> Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Notes"`

## Table of Contents

- [Notes Suite](#notes-suite)

---

## Notes Suite

Terms and Events for controlling the Notes application

### Commands

### `open note location`

Open a note URL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The URL to be opened. |

### `show`

Show an object in the UI

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `account | folder | note | attachment` | no | The object to be shown |
| `separately` | `boolean` | yes |  |

**Returns:** `account | folder | note | attachment` — The object shown.

### Classes

### `account`

a Notes account

**Contains:** `folder`, `note`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `default folder` | `folder` | r/w | the default folder for creating notes |
| `name` | `text` | r/w | the name of the account |
| `upgraded` | `boolean` | r/o | Is the account upgraded? |
| `id` | `text` | r/o | the unique identifier of the account |

**Responds to:** `show`

### `folder`

a folder containing notes

**Contains:** `folder`, `note`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | the name of the folder |
| `id` | `text` | r/o | the unique identifier of the folder |
| `shared` | `boolean` | r/o | Is the folder shared? |
| `container` | `account | folder` | r/o | the container of the folder |

**Responds to:** `show`

### `note`

a note in the Notes application

**Contains:** `attachment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | the name of the note (normally the first line of the body) |
| `id` | `text` | r/o | the unique identifier of the note |
| `container` | `folder` | r/o | the folder of the note |
| `body` | `text` | r/w | the HTML content of the note |
| `plaintext` | `text` | r/o | the plaintext content of the note |
| `creation date` | `date` | r/o | the creation date of the note |
| `modification date` | `date` | r/o | the modification date of the note |
| `password protected` | `boolean` | r/o | Is the note password protected? |
| `shared` | `boolean` | r/o | Is the note shared? |

**Responds to:** `show`

### `attachment`

an attachment in a note

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | the name of the attachment |
| `id` | `text` | r/o | the unique identifier of the attachment |
| `container` | `note` | r/o | the note containing the attachment |
| `content identifier` | `text` | r/o | the content-id URL in the note's HTML |
| `creation date` | `date` | r/o | the creation date of the attachment |
| `modification date` | `date` | r/o | the modification date of the attachment |
| `URL` | `text` | r/o | for URL attachments, the URL the attachment represents |
| `shared` | `boolean` | r/o | Is the attachment shared? |

**Responds to:** `show`, `save`

### Enumerations

### `saveable file format`

| Value | Description |
|-------|-------------|
| `native format` | Native format |
