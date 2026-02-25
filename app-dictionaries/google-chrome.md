# Scripting AppleScript Dictionary

> Auto-generated from `scripting.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Scripting"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Chromium Suite](#chromium-suite)

---

## Standard Suite

Common classes and commands for all applications.

### Commands

### `save`

Save an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object to save, usually a document or window |
| `in` | `file` | yes | The file in which to save the object. |
| `as` | `text` | yes | The file type in which to save the data. Can be 'only html', 'complete html', or 'single file'; default is 'complete html'. |

### `open`

Open a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file` | no | The file(s) to be opened. |

### `close`

Close a window.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the document(s) or window(s) to close. |

### `quit`

Quit the application.

### `count`

Return the number of elements of a particular class within an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object whose elements are to be counted |
| `each` | `type` | yes | The class of objects to be counted. |

**Returns:** `integer` — the number of elements

### `delete`

Delete an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object to delete |

### `duplicate`

Copy object(s) and put the copies at a new location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object(s) to duplicate |
| `to` | `location specifier` | yes | The location for the new object(s). |
| `with properties` | `record` | yes | Properties to be set in the new duplicated object(s). |

**Returns:** `specifier` — the duplicated object(s)

### `exists`

Verify if an object exists.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `any` | no | the object in question |

**Returns:** `boolean` — true if it exists, false if not

### `make`

Make a new object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `new` | `type` | no | The class of the new object. |
| `at` | `location specifier` | yes | The location at which to insert the object. |
| `with data` | `any` | yes | The initial contents of the object. |
| `with properties` | `record` | yes | The initial values for properties of the object. |

**Returns:** `specifier` — to the new object

### `move`

Move object(s) to a new location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object(s) to move |
| `to` | `location specifier` | no | The new location for the object(s). |

**Returns:** `specifier` — the moved object(s)

### `print`

Print an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The file(s) or document(s) to be printed. |

### Classes

### `application`

The application's top-level scripting object.

**Contains:** `window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the application. |
| `frontmost` | `boolean` | r/o | Is this the frontmost (active) application? |
| `version` | `text` | r/o | The version of the application. |

**Responds to:** `open`, `print`, `quit`

### `window`

A window.

**Contains:** `tab`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `given name` | `text` | r/w | The given name of the window. |
| `name` | `text` | r/o | The full title of the window. |
| `id` | `text` | r/o | The unique identifier of the window. |
| `index` | `integer` | r/w | The index of the window, ordered front to back. |
| `bounds` | `rectangle` | r/w | The bounding rectangle of the window. |
| `closeable` | `boolean` | r/o | Whether the window has a close box. |
| `minimizable` | `boolean` | r/o | Whether the window can be minimized. |
| `minimized` | `boolean` | r/w | Whether the window is currently minimized. |
| `resizable` | `boolean` | r/o | Whether the window can be resized. |
| `visible` | `boolean` | r/w | Whether the window is currently visible. |
| `zoomable` | `boolean` | r/o | Whether the window can be zoomed. |
| `zoomed` | `boolean` | r/w | Whether the window is currently zoomed. |
| `active tab` | `tab` | r/o | Returns the currently selected tab |
| `mode` | `text` | r/w | Represents the mode of the window which can be 'normal' or 'incognito', can be set only once during creation of the window. |
| `active tab index` | `integer` | r/w | The index of the active tab. |

**Responds to:** `close`

---

## Chromium Suite

Common classes and commands for Chrome.

### Commands

### `reload`

Reload a tab.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `go back`

Go Back (If Possible).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `go forward`

Go Forward (If Possible).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `select all`

Select all.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `cut selection`

Cut selected text (If Possible).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `copy selection`

Copy text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `paste selection`

Paste text (If Possible).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `undo`

Undo the last change.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `redo`

Redo the last change.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `stop`

Stop the current tab from loading.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `view source`

View the HTML source of the tab.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |

### `execute`

Execute a piece of javascript.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The tab to execute the command in. |
| `javascript` | `text` | no | The javascript code to execute. |

**Returns:** `any`

### Classes

### `tab`

A tab.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | Unique ID of the tab. |
| `title` | `text` | r/o | The title of the tab. |
| `URL` | `text` | r/w | The url visible to the user. |
| `loading` | `boolean` | r/o | Is loading? |

**Responds to:** `undo`, `redo`, `cut selection`, `copy selection`, `paste selection`, `select all`, `go back`, `go forward`, `reload`, `stop`, `print`, `view source`, `save`, `close`, `execute`

### `bookmark folder`

A bookmarks folder that contains other bookmarks folder and bookmark items.

**Contains:** `bookmark folder`, `bookmark item`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | Unique ID of the bookmark folder. |
| `title` | `text` | r/w | The title of the folder. |
| `index` | `number` | r/o | Returns the index with respect to its parent bookmark folder. |

### `bookmark item`

An item consists of an URL and the title of a bookmark

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | Unique ID of the bookmark item. |
| `title` | `text` | r/w | The title of the bookmark item. |
| `URL` | `text` | r/w | The URL of the bookmark. |
| `index` | `number` | r/o | Returns the index with respect to its parent bookmark folder. |
