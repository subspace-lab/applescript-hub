# Iterm2 AppleScript Dictionary

> Auto-generated from `iTerm2.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Iterm2"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [iTerm2 Suite](#iterm2-suite)

---

## Standard Suite

Common classes and commands for all applications.

### Commands

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

The application's top-level scripting object.

**Contains:** `window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `current window` | `window` | r/w | The frontmost window |
| `name` | `text` | r/o | The name of the application. |
| `frontmost` | `boolean` | r/o | Is this the frontmost (active) application? |
| `version` | `text` | r/o | The version of the application. |

### `window`

A window.

**Contains:** `tab`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `integer` | r/o | The unique identifier of the session. |
| `alternate identifier` | `text` | r/o | The alternate unique identifier of the session. |
| `name` | `text` | r/o | The full title of the window. |
| `index` | `integer` | r/w | The index of the window, ordered front to back. |
| `bounds` | `rectangle` | r/w | The bounding rectangle of the window. |
| `closeable` | `boolean` | r/o | Whether the window has a close box. |
| `miniaturizable` | `boolean` | r/o | Whether the window can be minimized. |
| `miniaturized` | `boolean` | r/w | Whether the window is currently minimized. |
| `resizable` | `boolean` | r/o | Whether the window can be resized. |
| `visible` | `boolean` | r/w | Whether the window is currently visible. |
| `zoomable` | `boolean` | r/o | Whether the window can be zoomed. |
| `zoomed` | `boolean` | r/w | Whether the window is currently zoomed. |
| `frontmost` | `boolean` | r/w | Whether the window is currently the frontmost window. |
| `current tab` | `tab` | r/w | The currently selected tab |
| `current session` | `session` | r/w | The current session in a window |
| `position` | `point` | r/w | The position of the window, relative to the upper left corner of the screen. |
| `origin` | `point` | r/w | The position of the window, relative to the lower left corner of the screen. |
| `size` | `point` | r/w | The width and height of the window |
| `frame` | `rectangle` | r/w | The bounding rectangle, relative to the lower left corner of the screen. |

**Responds to:** `close`, `select`, `create tab with default profile`, `create tab`

### Enumerations

### `save options`

| Value | Description |
|-------|-------------|
| `yes` | Save the file. |
| `no` | Do not save the file. |
| `ask` | Ask the user whether or not to save the file. |

---

## iTerm2 Suite

Classes just for the iTerm2 application.

### Commands

### `close`

Close a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the session, tab, or window to close. |

### `create tab`

Create a new tab

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the session, tab, or window to close. |
| `with profile` | `text` | no | The profile name |
| `command` | `text` | yes | Shell command to run |

**Returns:** `tab` — The tab that was created

### `create tab with default profile`

Create a new tab with the default profile

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The window in which to create a new tab |
| `command` | `text` | yes | Shell command to run |

**Returns:** `tab` — The tab that was created

### `create window with profile`

Create a new window

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The profile name |
| `command` | `text` | yes | Shell command to run |

**Returns:** `window` — The window that was created

### `create window with default profile`

Create a new window with the default profile

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `command` | `text` | yes | Shell command to run |

**Returns:** `window` — The window that was created

### `write`

Send text as though it was typed.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The session to send to |
| `contents of file` | `file` | yes | Filename to send the contents of |
| `text` | `text` | yes | Text to send |
| `newline` | `boolean` | yes | If newline should be added to end of text (default: yes) |

### `select`

Make receiver visible and selected.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |

### `split vertically`

Split a session vertically.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |
| `with profile` | `text` | no | Name of profile for new session. |
| `command` | `text` | yes | Shell command to run |

**Returns:** `session` — The session that was created

### `split vertically with default profile`

Split a session vertically, using the default profile for the new session

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |
| `command` | `text` | yes | Shell command to run |

**Returns:** `session` — The session that was created

### `split vertically with same profile`

Split a session vertically, using the original session's profile for the new session

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |
| `command` | `text` | yes | Shell command to run |

**Returns:** `session` — The session that was created

### `split horizontally`

Split a session horizontally.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |
| `with profile` | `text` | no | Name of profile for new session. |
| `command` | `text` | yes | Shell command to run |

**Returns:** `session` — The session that was created

### `split horizontally with default profile`

Split a session horizontally, using the default profile for the new session

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |
| `command` | `text` | yes | Shell command to run |

**Returns:** `session` — The session that was created

### `split horizontally with same profile`

Split a session horizontally, using the original session's profile for the new session

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |
| `command` | `text` | yes | Shell command to run |

**Returns:** `session` — The session that was created

### `variable`

Returns the value of a session variable with the given name

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |
| `named` | `text` | no | Name of variable |

**Returns:** `text` — The variable's value

### `set variable`

Sets the value of a session variable

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object to send to |
| `named` | `text` | no | Name of variable |
| `to` | `text` | no | New value |

**Returns:** `text` — The variable's value

### Classes

### `tab`

A terminal tab

**Contains:** `session`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `current session` | `session` | r/w | The current session in a tab |
| `index` | `integer` | r/w | Index of tab in parent tab view control |

**Responds to:** `close`, `select`

### `session`

A terminal session

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The unique identifier of the session. |
| `is processing` | `boolean` | r/w | The session has received output recently. |
| `is at shell prompt` | `boolean` | r/w | The terminal is at the shell prompt. Requires shell integration. |
| `columns` | `integer` | r/w |  |
| `rows` | `integer` | r/w |  |
| `tty` | `text` | r/o |  |
| `contents` | `text` | r/w | The currently visible contents of the session. |
| `text` | `text` | r/o | The currently visible contents of the session. |
| `background color` | `RGB color` | r/w |  |
| `bold color` | `RGB color` | r/w |  |
| `cursor color` | `RGB color` | r/w |  |
| `cursor text color` | `RGB color` | r/w |  |
| `foreground color` | `RGB color` | r/w |  |
| `selected text color` | `RGB color` | r/w |  |
| `selection color` | `RGB color` | r/w |  |
| `ANSI black color` | `RGB color` | r/w |  |
| `ANSI red color` | `RGB color` | r/w |  |
| `ANSI green color` | `RGB color` | r/w |  |
| `ANSI yellow color` | `RGB color` | r/w |  |
| `ANSI blue color` | `RGB color` | r/w |  |
| `ANSI magenta color` | `RGB color` | r/w |  |
| `ANSI cyan color` | `RGB color` | r/w |  |
| `ANSI white color` | `RGB color` | r/w |  |
| `ANSI bright black color` | `RGB color` | r/w |  |
| `ANSI bright red color` | `RGB color` | r/w |  |
| `ANSI bright green color` | `RGB color` | r/w |  |
| `ANSI bright yellow color` | `RGB color` | r/w |  |
| `ANSI bright blue color` | `RGB color` | r/w |  |
| `ANSI bright magenta color` | `RGB color` | r/w |  |
| `ANSI bright cyan color` | `RGB color` | r/w |  |
| `ANSI bright white color` | `RGB color` | r/w |  |
| `background image` | `text` | r/w |  |
| `name` | `text` | r/w |  |
| `transparency` | `real` | r/w |  |
| `unique ID` | `text` | r/o |  |
| `profile name` | `text` | r/o | The session's profile name |
| `answerback string` | `text` | r/w | ENQ Answerback string |

**Responds to:** `close`, `write`, `select`, `split vertically`, `split vertically with default profile`, `split vertically with same profile`, `split horizontally`, `split horizontally with default profile`, `split horizontally with same profile`, `variable`, `set variable`
