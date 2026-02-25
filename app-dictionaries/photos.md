# Photos AppleScript Dictionary

> Auto-generated from `Photos.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Photos"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Photos Suite](#photos-suite)

---

## Standard Suite

Common classes and commands for all applications.

### Commands

### `count`

Return the number of elements of a particular class within an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The objects to be counted. |
| `each` | `type` | yes | The class of objects to be counted. |

**Returns:** `integer` — The count.

### `exists`

Verify that an object exists.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `any` | no | The object(s) to check. |

**Returns:** `boolean` — Did the object(s) exist?

### `open`

Open a photo library

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file` | no | The photo library to be opened. |

### `quit`

Quit the application.

### Classes

### `application`

The application's top-level scripting object.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the application. |
| `frontmost` | `boolean` | r/o | Is this the active application? |
| `version` | `text` | r/o | The version number of the application. |

**Responds to:** `open`, `quit`

---

## Photos Suite

Classes and commands for Photos

### Commands

### `import`

Import files into the library

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file` | no | The list of files to copy. |
| `into` | `album` | yes | The album to import into. |
| `skip check duplicates` | `boolean` | yes | Skip duplicate checking and import everything, defaults to false. |

**Returns:** `media item` — The imported media items in an array

### `export`

Export media items to the specified location as files

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `media item` | no | The list of media items to export. |
| `to` | `file` | no | The destination of the export. |
| `using originals` | `boolean` | yes | Export the original files if true, otherwise export rendered jpgs. defaults to false. |

### `duplicate`

Duplicate an object.  Only media items can be duplicated

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `media item` | no | The media item to duplicate |

**Returns:** `media item` — The duplicated media item

### `make`

Create a new object.  Only new albums and folders can be created.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `new` | `type` | no | The class of the new object, allowed values are album or folder |
| `named` | `text` | yes | The name of the new object. |
| `at` | `folder` | yes | The parent folder for the new object. |

**Returns:** `album | folder` — The new object.

### `delete`

Delete an object.  Only albums and folders can be deleted.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `album | folder` | no | The album or folder to delete. |

### `add`

Add media items to an album.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `media item` | no | The list of media items to add. |
| `to` | `album` | no | The album to add to. |

### `start slideshow`

Display an ad-hoc slide show from a list of media items, an album, or a folder.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `using` | `media item` | no | The media items to show. |

### `stop slideshow`

End the currently-playing slideshow.

### `next slide`

Skip to next slide in currently-playing slideshow.

### `previous slide`

Skip to previous slide in currently-playing slideshow.

### `pause slideshow`

Pause the currently-playing slideshow.

### `resume slideshow`

Resume the currently-playing slideshow.

### `spotlight`

Show the image at path in the application, used to show spotlight search results

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text | media item | container` | no | The full path to the image |

### `search`

search for items matching the search string. Identical to entering search text in the Search field in Photos

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `for` | `text` | no | The text to search for |

**Returns:** `media item` — reference(s) to found media item(s)

### Classes

### `media item`

A media item, such as a photo or video.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `keywords` | `text` | r/w | A list of keywords to associate with a media item |
| `name` | `text` | r/w | The name (title) of the media item. |
| `description` | `text` | r/w | A description of the media item. |
| `favorite` | `boolean` | r/w | Whether the media item has been favorited. |
| `date` | `date` | r/w | The date of the media item |
| `id` | `text` | r/o | The unique ID of the media item |
| `height` | `integer` | r/o | The height of the media item in pixels. |
| `width` | `integer` | r/o | The width of the media item in pixels. |
| `filename` | `text` | r/o | The name of the file on disk. |
| `altitude` | `real` | r/o | The GPS altitude in meters. |
| `size` | `integer` | r/w | The selected media item file size. |
| `location` | `real | missing value` | r/w | The GPS latitude and longitude, in an ordered list of 2 numbers or missing values.  Latitude in range -90.0 to 90.0, longitude in range -180.0 to 180.0. |

**Responds to:** `duplicate`, `spotlight`

### `container`

Base class for collections that contains other items, such as albums and folders

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The unique ID of this container. |
| `name` | `text` | r/w | The name of this container. |
| `parent` | `folder` | r/o | This container's parent folder, if any. |

**Responds to:** `spotlight`

### `album`

An album. A container that holds media items

**Contains:** `media item`

### `folder`

A folder. A container that holds albums and other folders, but not media items

**Contains:** `container`, `album`, `folder`

### `moment`

A set of media items that represents a Moment.

**Contains:** `media item`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The unique ID of the Moment. |
| `name` | `text` | r/o | The name of the Moment. |
