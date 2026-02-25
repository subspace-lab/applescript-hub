# Safari AppleScript Dictionary

> Auto-generated from `Safari.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Safari"`

## Table of Contents

- [Safari suite](#safari-suite)

---

## Safari suite

Safari specific classes

### Commands

### `add reading list item`

Add a new Reading List item with the given URL. Allows a custom title and preview text to be specified.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | URL of the Reading List item |
| `and preview text` | `text` | yes | Preview text for the Reading List item, usually the first few sentences of the article |
| `with title` | `text` | yes | Title of the Reading List item |

### `do JavaScript`

Applies a string of JavaScript code to a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The JavaScript code to evaluate. |
| `in` | `document | tab` | yes | The tab that the JavaScript should be evaluated in. |

**Returns:** `any`

### `email contents`

Emails the contents of a tab.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `of` | `tab | document` | yes | The tab to send. |

### `search the web`

Searches the web using Safari's current search provider.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `in` | `document | tab` | yes | The tab that the search results should shown in. |
| `for` | `text` | no | The query to search for. |

### `show bookmarks`

Shows Safari's bookmarks.

### `show extensions preferences`

Show Safari Extensions preferences.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The identifier of the extension to select. |

### `dispatch message to extension`

Dispatch a message to a Safari Extension.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `any` | no | A dictionary describing the message |

### `sync all plist to disk`

Make sure that all in-memory structures are in-sync with their on-disk counterparts.

### `show privacy report`

Show Safari's Privacy Report

### `show credit card settings`

Show Safari Credit Card Settings.

### Classes

### `tab`

A Safari window tab.

*Plural:* `tabs`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `source` | `text` | r/o | The HTML source of the web page currently loaded in the tab. |
| `URL` | `text` | r/w | The current URL of the tab. |
| `index` | `number` | r/o | The index of the tab, ordered left to right. |
| `text` | `text` | r/o | The text of the web page currently loaded in the tab. Modifications to text aren't reflected on the web page. |
| `visible` | `boolean` | r/o | Whether the tab is currently visible. |
| `name` | `text` | r/o | The name of the tab. |
| `pid` | `number` | r/o | The pid of the WebContent process backing the tab, if it exists. |

**Responds to:** `do JavaScript`, `email contents`, `close`, `search the web`

### `sourceProvider`

**Responds to:** `get`

### `contentsProvider`

**Responds to:** `get`
