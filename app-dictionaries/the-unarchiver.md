# Theunarchiversuite AppleScript Dictionary

> Auto-generated from `TheUnarchiverSuite.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Theunarchiversuite"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [The Unarchiver Suite](#the-unarchiver-suite)

---

## Standard Suite

Common classes and commands for all applications.

### Commands

### `open`

Open a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file | file` | no | The file(s) to be opened. |

### `close`

Close a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the document(s) or window(s) to close. |
| `saving` | `save options` | yes | Should changes be saved before closing? |
| `saving in` | `file` | yes | The file in which to save the document, if so. |

### `quit`

Quit the application.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `saving` | `save options` | yes | Should changes be saved before quitting? |

### `count`

Return the number of elements of a particular class within an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The objects to be counted. |
| `each` | `type` | yes | The class of objects to be counted. |

**Returns:** `integer` — The count.

### `delete`

Delete an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object(s) to delete. |

### `exists`

Verify that an object exists.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `any` | no | The object(s) to check. |

**Returns:** `boolean` — Did the object(s) exist?

### `make`

Create a new object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `new` | `type` | no | The class of the new object. |
| `at` | `location specifier` | yes | The location at which to insert the object. |
| `with data` | `any` | yes | The initial contents of the object. |
| `with properties` | `record` | yes | The initial values for properties of the object. |

**Returns:** `specifier` — The new object.

### Classes

### `application`

The application's top-level scripting object.

**Contains:** `window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the application. |
| `frontmost` | `boolean` | r/o | Is this the active application? |
| `version` | `text` | r/o | The version number of the application. |
| `isRunningExtractions` | `boolean` | r/o |  |

**Responds to:** `open`, `quit`

### `window`

A window.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The title of the window. |
| `id` | `integer` | r/o | The unique identifier of the window. |
| `index` | `integer` | r/w | The index of the window, ordered front to back. |
| `bounds` | `rectangle` | r/w | The bounding rectangle of the window. |
| `closeable` | `boolean` | r/o | Does the window have a close button? |
| `miniaturizable` | `boolean` | r/o | Does the window have a minimize button? |
| `miniaturized` | `boolean` | r/w | Is the window minimized right now? |
| `resizable` | `boolean` | r/o | Can the window be resized? |
| `visible` | `boolean` | r/w | Is the window visible right now? |
| `zoomable` | `boolean` | r/o | Does the window have a zoom button? |
| `zoomed` | `boolean` | r/w | Is the window zoomed right now? |

**Responds to:** `close`

### Enumerations

### `save options`

| Value | Description |
|-------|-------------|
| `yes` | Save the file. |
| `no` | Do not save the file. |
| `ask` | Ask the user whether or not to save the file. |

---

## The Unarchiver Suite

Commands and definitions for unarchaving tasks

### Commands

### `unarchive`

Unarchive the archives to the destination. If an option is not given, the user default settings will be used. Raise an error with number 1 and desription:'The file /the/path/ doesn't exist.' if a file doesn't exist.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The files to unarchive, always use POSIX paths (i.e. '/the/path/to/the/archive' ) |
| `to` | `text | File Destination` | yes | Where to unarchive as an absolute path or a constant |
| `deleting Original` | `boolean` | yes | If the archives should be deleted when the operation is finished |
| `creating Folder` | `Create new folder` | yes | Create a new folder for every archive. |
| `wait Until Finished` | `boolean` | yes | The script execution is stopped until all the unarhiving tasks are completed. Default is 'True' |

### Enumerations

### `File Destination`

Where to unarchive

| Value | Description |
|-------|-------------|
| `Original` | The same folder as the archive |
| `Ask User` | Ask the user |
| `Desktop` | The user desktop |
| `User Default` | The folder selected by the user on the preferences panel |

### `Create new folder`

| Value | Description |
|-------|-------------|
| `Never` |  |
| `Only` | Only if there is more than one element on the top level |
| `Always` |  |
