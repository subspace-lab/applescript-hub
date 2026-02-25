# Terminal AppleScript Dictionary

> Auto-generated from `Terminal.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Terminal"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Terminal Suite](#terminal-suite)

---

## Standard Suite

Common classes and commands for all applications.

### Commands

### `open`

Open a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file` | no | The file(s) to be opened. |

### `close`

Close a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the document(s) or window(s) to close. |
| `saving` | `save options` | yes | Whether or not changes should be saved before closing. |
| `saving in` | `file` | yes | The file in which to save the document. |

### `save`

Save a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The document(s) or window(s) to save. |
| `in` | `file` | yes | The file in which to save the document. |

### `print`

Print a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file | specifier` | no | The file(s), document(s), or window(s) to be printed. |
| `with properties` | `print settings` | yes | The print settings to use. |
| `print dialog` | `boolean` | yes | Should the application show the print dialog? |

### `quit`

Quit the application.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `saving` | `save options` | yes | Whether or not changed documents should be saved before closing. |

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
| `to` | `location specifier` | no | The location for the new object(s). |
| `with properties` | `record` | yes | Properties to be set in the new duplicated object(s). |

### `exists`

Verify if an object exists.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object in question |

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

### Classes

### `application`

The application‘s top-level scripting object.

**Contains:** `window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the application. |
| `frontmost` | `boolean` | r/o | Is this the frontmost (active) application? |
| `version` | `text` | r/o | The version of the application. |

**Responds to:** `open`, `print`, `quit`, `get URL`

### `window`

A window.

**Contains:** `tab`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The full title of the window. |
| `id` | `integer` | r/o | The unique identifier of the window. |
| `index` | `integer` | r/w | The index of the window, ordered front to back. |
| `bounds` | `rectangle` | r/w | The bounding rectangle of the window. |
| `closeable` | `boolean` | r/o | Whether the window has a close box. |
| `miniaturizable` | `boolean` | r/o | Whether the window can be minimized. |
| `miniaturized` | `boolean` | r/w | Whether the window is currently minimized. |
| `resizable` | `boolean` | r/o | Whether the window can be resized. |
| `visible` | `boolean` | r/w | Whether the window is currently visible. |
| `zoomable` | `boolean` | r/o | Whether the window can be zoomed. |
| `zoomed` | `boolean` | r/w | Whether the window is currently zoomed. |
| `frontmost` | `boolean` | r/w | Whether the window is currently the frontmost Terminal window. |
| `position` | `point` | r/w | The position of the window, relative to the upper left corner of the screen. |
| `origin` | `point` | r/w | The position of the window, relative to the lower left corner of the screen. |
| `size` | `point` | r/w | The width and height of the window |
| `frame` | `rectangle` | r/w | The bounding rectangle, relative to the lower left corner of the screen. |

**Responds to:** `close`, `print`, `save`

### Enumerations

### `save options`

| Value | Description |
|-------|-------------|
| `yes` | Save the file. |
| `no` | Do not save the file. |
| `ask` | Ask the user whether or not to save the file. |

### `printing error handling`

| Value | Description |
|-------|-------------|
| `standard` | Standard PostScript error handling |
| `detailed` | print a detailed report of PostScript errors |

---

## Terminal Suite

Terminal specific classes.

### Commands

### `do script`

Runs a UNIX shell script or command.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | yes | The command to execute. |
| `with command` | `text | any` | yes | Data to be passed to the Terminal application as the command line. Deprecated; use direct parameter instead. |
| `in` | `tab | window | any` | yes | The tab in which to execute the command |

**Returns:** `tab` — The tab the command was executed in.

### `get URL`

Open a command an ssh, telnet, or x-man-page URL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The URL to open. |

### Classes

### `settings set`

A set of settings.

*Plural:* `settings sets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `integer` | r/o | The unique identifier of the settings set. |
| `name` | `text` | r/w | The name of the settings set. |
| `number of rows` | `integer` | r/w | The number of rows displayed in the tab. |
| `number of columns` | `integer` | r/w | The number of columns displayed in the tab. |
| `cursor color` | `color` | r/w | The cursor color for the tab. |
| `background color` | `color` | r/w | The background color for the tab. |
| `normal text color` | `color` | r/w | The normal text color for the tab. |
| `bold text color` | `color` | r/w | The bold text color for the tab. |
| `font name` | `text` | r/w | The name of the font used to display the tab’s contents. |
| `font size` | `integer` | r/w | The size of the font used to display the tab’s contents. |
| `font antialiasing` | `boolean` | r/w | Whether the font used to display the tab’s contents is antialiased. |
| `clean commands` | `text` | r/w | The processes which will be ignored when checking whether a tab can be closed without showing a prompt. |
| `title displays device name` | `boolean` | r/w | Whether the title contains the device name. |
| `title displays shell path` | `boolean` | r/w | Whether the title contains the shell path. |
| `title displays window size` | `boolean` | r/w | Whether the title contains the tab’s size, in rows and columns. |
| `title displays settings name` | `boolean` | r/w | Whether the title contains the settings name. |
| `title displays custom title` | `boolean` | r/w | Whether the title contains a custom title. |
| `custom title` | `text` | r/w | The tab’s custom title. |

### `tab`

A tab.

*Plural:* `tabs`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `number of rows` | `integer` | r/w | The number of rows displayed in the tab. |
| `number of columns` | `integer` | r/w | The number of columns displayed in the tab. |
| `contents` | `text` | r/o | The currently visible contents of the tab. |
| `history` | `text` | r/o | The contents of the entire scrolling buffer of the tab. |
| `busy` | `boolean` | r/o | Whether the tab is busy running a process. |
| `processes` | `text` | r/o | The processes currently running in the tab. |
| `selected` | `boolean` | r/w | Whether the tab is selected. |
| `title displays custom title` | `boolean` | r/w | Whether the title contains a custom title. |
| `custom title` | `text` | r/w | The tab’s custom title. |
| `tty` | `text` | r/o | The tab’s TTY device. |
| `current settings` | `settings set` | r/w | The set of settings which control the tab’s behavior and appearance. |
| `cursor color` | `color` | r/w | The cursor color for the tab. |
| `background color` | `color` | r/w | The background color for the tab. |
| `normal text color` | `color` | r/w | The normal text color for the tab. |
| `bold text color` | `color` | r/w | The bold text color for the tab. |
| `clean commands` | `text` | r/w | The processes which will be ignored when checking whether a tab can be closed without showing a prompt. |
| `title displays device name` | `boolean` | r/w | Whether the title contains the device name. |
| `title displays shell path` | `boolean` | r/w | Whether the title contains the shell path. |
| `title displays window size` | `boolean` | r/w | Whether the title contains the tab’s size, in rows and columns. |
| `title displays file name` | `boolean` | r/w | Whether the title contains the file name. |
| `font name` | `text` | r/w | The name of the font used to display the tab’s contents. |
| `font size` | `integer` | r/w | The size of the font used to display the tab’s contents. |
| `font antialiasing` | `boolean` | r/w | Whether the font used to display the tab’s contents is antialiased. |
