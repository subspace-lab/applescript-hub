# Pages AppleScript Dictionary

> Auto-generated from `Pages.sdef` inside the app bundle.  
> Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Pages"`

## Table of Contents

- [iWork Text Suite](#iwork-text-suite)
- [iWork Suite](#iwork-suite)
- [Pages Suite](#pages-suite)

---

## iWork Text Suite

Classes and commands for iWorks text objects.

### Classes

### `rich text`

This provides the base rich text class for all iWork applications.

*Plural:* `rich text`

**Contains:** `character`, `paragraph`, `word`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `color` | r/w | The color of the font. Expressed as an RGB value consisting of a list of three color values from 0 to 65535. ex: Blue = {0, 0, 65535}. |
| `font` | `text` | r/w | The name of the font.  Can be the PostScript name, such as: “TimesNewRomanPS-ItalicMT”, or display name: “Times New Roman Italic”. TIP: Use the Font Book application get the information about a typeface. |
| `size` | `real` | r/w | The size of the font. |

### `character`

One of some text's characters.

### `paragraph`

One of some text's paragraphs.

**Contains:** `character`, `word`

### `word`

One of some text's words.

**Contains:** `character`

---

## iWork Suite

A set of basic classes for iWork applications.

### Commands

### `set`

Sets the value of the specified object(s).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `to` | `any` | no | The new value. |

### `delete`

Delete an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

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

### `clear`

Clear the contents of a specified range of cells, including formatting and style.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `merge`

Merge a specified range of cells.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `sort`

Sort the rows of the table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `table` | no |  |
| `by` | `column` | no | The column to sort by. |
| `direction` | `NMSD` | yes |  |
| `in rows` | `range` | yes | Limit the sort to the specified rows. |

### `unmerge`

Unmerge all merged cells in a specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `set password`

Set a password to an unencrypted document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no |  |
| `to` | `document` | yes | The document to set a password to. If unspecified, the current target is assumed. |
| `hint` | `text` | yes | A hint for the password. |
| `saving in keychain` | `boolean` | yes | Whether to save the password in keychain or not. By default, the password is not saved in the keychain. |

### `remove password`

Remove the password from the document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The current password of the document. |
| `from` | `document` | yes | The document from which to remove the password. If unspecified, the current target is assumed. |

### Classes

### `iWork container`

A container for iWork items

**Contains:** `audio clip`, `chart`, `image`, `iWork item`, `group`, `line`, `movie`, `shape`, `table`, `text item`

### `iWork item`

An item which supports formatting

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `height` | `integer` | r/w | The height of the iWork item. |
| `locked` | `boolean` | r/w | Whether the object is locked. |
| `parent` | `iWork container` | r/o | The iWork container containing this iWork item. |
| `position` | `point` | r/w | The horizontal and vertical coordinates of the top left point of the iWork item. |
| `width` | `integer` | r/w | The width of the iWork item. |

### `audio clip`

An audio clip

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `file | text` | r/w | The name of the audio file. |
| `clip volume` | `integer` | r/w | The volume setting for the audio clip, from 0 (none) to 100 (full volume). |
| `repetition method` | `playback repetition method` | r/w | If or how the audio clip repeats. |

### `shape`

A shape container

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `background fill type` | `item fill options` | r/o | The background, if any, for the shape. |
| `object text` | `rich text` | r/w | The text contained within the shape. |
| `reflection showing` | `boolean` | r/w | Is the iWork item displaying a reflection? |
| `reflection value` | `integer` | r/w | The percentage of reflection of the iWork item, from 0 (none) to 100 (full). |
| `rotation` | `integer` | r/w | The rotation of the iWork item, in degrees from 0 to 359. |
| `opacity` | `integer` | r/w | The opacity of the object, in percent. |

### `chart`

A chart

### `image`

An image container

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `description` | `text` | r/w | Text associated with the image, read aloud by VoiceOver. |
| `file` | `file` | r/o | The image file. |
| `file name` | `file | text` | r/w | The name of the image file. |
| `opacity` | `integer` | r/w | The opacity of the object, in percent. |
| `reflection showing` | `boolean` | r/w | Is the iWork item displaying a reflection? |
| `reflection value` | `integer` | r/w | The percentage of reflection of the iWork item, from 0 (none) to 100 (full). |
| `rotation` | `integer` | r/w | The rotation of the iWork item, in degrees from 0 to 359. |

### `group`

A group container

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `height` | `integer` | r/w | The height of the iWork item. |
| `parent` | `iWork container` | r/o | The iWork container containing this iWork item. |
| `position` | `point` | r/w | The horizontal and vertical coordinates of the top left point of the iWork item. |
| `width` | `integer` | r/w | The width of the iWork item. |
| `rotation` | `integer` | r/w | The rotation of the iWork item, in degrees from 0 to 359. |

### `line`

A line

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `end point` | `point` | r/w | A list of two numbers indicating the horizontal and vertical position of the line ending point. |
| `reflection showing` | `boolean` | r/w | Is the iWork item displaying a reflection? |
| `reflection value` | `integer` | r/w | The percentage of reflection of the iWork item, from 0 (none) to 100 (full). |
| `rotation` | `integer` | r/w | The rotation of the iWork item, in degrees from 0 to 359. |
| `start point` | `point` | r/w | A list of two numbers indicating the horizontal and vertical position of the line starting point. |

### `movie`

A movie container

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `file | text` | r/w | The name of the movie file. |
| `movie volume` | `integer` | r/w | The volume setting for the movie, from 0 (none) to 100 (full volume). |
| `opacity` | `integer` | r/w | The opacity of the object, in percent. |
| `reflection showing` | `boolean` | r/w | Is the iWork item displaying a reflection? |
| `reflection value` | `integer` | r/w | The percentage of reflection of the iWork item, from 0 (none) to 100 (full). |
| `repetition method` | `playback repetition method` | r/w | If or how the movie repeats. |
| `rotation` | `integer` | r/w | The rotation of the iWork item, in degrees from 0 to 359. |

### `table`

A table

**Contains:** `cell`, `row`, `column`, `range`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | The item's name. |
| `cell range` | `range` | r/o | The range describing every cell in the table. |
| `selection range` | `range` | r/w | The cells currently selected in the table. |
| `row count` | `integer` | r/w | The number of rows in the table. |
| `column count` | `integer` | r/w | The number of columns in the table. |
| `header row count` | `integer` | r/w | The number of header rows in the table. |
| `header column count` | `integer` | r/w | The number of header columns in the table. |
| `footer row count` | `integer` | r/w | The number of footer rows in the table. |

**Responds to:** `sort`

### `text item`

A text container

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `background fill type` | `item fill options` | r/o | The background, if any, for the text item. |
| `object text` | `rich text` | r/w | The text contained within the text item. |
| `opacity` | `integer` | r/w | The opacity of the object, in percent. |
| `reflection showing` | `boolean` | r/w | Is the iWork item displaying a reflection? |
| `reflection value` | `integer` | r/w | The percentage of reflection of the iWork item, from 0 (none) to 100 (full). |
| `rotation` | `integer` | r/w | The rotation of the iWork item, in degrees from 0 to 359. |

### `range`

A range of cells in a table

**Contains:** `cell`, `column`, `row`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `font name` | `text` | r/w | The font of the range's cells. |
| `font size` | `real` | r/w | The font size of the range's cells. |
| `format` | `NMCT | any` | r/w | The format of the range's cells. |
| `alignment` | `tAHT` | r/w | The horizontal alignment of content in the range's cells. |
| `name` | `text` | r/o | The range's coordinates. |
| `text color` | `color` | r/w | The text color of the range's cells. |
| `text wrap` | `boolean` | r/w | Whether text should wrap in the range's cells. |
| `background color` | `color` | r/w | The background color of the range's cells. |
| `vertical alignment` | `tAVT` | r/w | The vertical alignment of content in the range's cells. |

**Responds to:** `clear`, `merge`, `unmerge`

### `cell`

A cell in a table

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `column` | `column` | r/o | The cell's column. |
| `row` | `row` | r/o | The cell's row. |
| `value` | `number | date | text | boolean | missing value` | r/w | The actual value in the cell, or missing value if the cell is empty. |
| `formatted value` | `text` | r/o | The formatted value in the cell, or missing value if the cell is empty. |
| `formula` | `text` | r/o | The formula in the cell, as text, e.g. =SUM(40+2). If the cell does not contain a formula, returns missing value. To set the value of a cell to a formula as text, use the value property. |

### `row`

A row of cells in a table

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `address` | `integer` | r/o | The row's index in the table (e.g., the second row has address 2). |
| `height` | `real` | r/w | The height of the row. |

### `column`

A column of cells in a table

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `address` | `integer` | r/o | The column's index in the table (e.g., the second column has address 2). |
| `width` | `real` | r/w | The width of the column. |

### Enumerations

### `tAVT`

| Value | Description |
|-------|-------------|
| `bottom` | Right-align content. |
| `center` | Center-align content. |
| `top` | Top-align content. |

### `tAHT`

| Value | Description |
|-------|-------------|
| `auto align` | Auto-align based on content type. |
| `center` | Center-align content. |
| `justify` | Fully justify (left and right) content. |
| `left` | Left-align content. |
| `right` | Right-align content. |

### `NMSD`

| Value | Description |
|-------|-------------|
| `ascending` | Sort in increasing value order |
| `descending` | Sort in decreasing value order |

### `NMCT`

| Value | Description |
|-------|-------------|
| `automatic` | Automatic format |
| `checkbox` | Checkbox control format (Numbers only) |
| `currency` | Currency number format |
| `date and time` | Date and time format |
| `fraction` | Fraction number format |
| `number` | Decimal number format |
| `percent` | Percentage number format |
| `pop up menu` | Pop-up menu control format (Numbers only) |
| `scientific` | Scientific notation format |
| `slider` | Slider control format (Numbers only) |
| `stepper` | Stepper control format (Numbers only) |
| `text` | Text format |
| `duration` | Duration format |
| `rating` | Rating format. (Numbers only) |
| `numeral system` | Numeral System |

### `item fill options`

| Value | Description |
|-------|-------------|
| `no fill` |  |
| `color fill` |  |
| `gradient fill` |  |
| `advanced gradient fill` |  |
| `image fill` |  |
| `advanced image fill` |  |

### `playback repetition method`

| Value | Description |
|-------|-------------|
| `none` |  |
| `loop` |  |
| `loop back and forth` |  |

---

## Pages Suite

Classes and commands for Pages.

### Commands

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

### `export`

Export a document to another file

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | The document to export |
| `to` | `file` | no | the destination file |
| `as` | `export format` | no | The format to use. |
| `with properties` | `export options` | yes | Optional export settings. |

### Classes

### `template`

A styled document layout.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The identifier used by the application. |
| `name` | `text` | r/o | The localized name displayed to the user. |

### `section`

A section within a document.

*Plural:* `sections`

**Contains:** `audio clip`, `chart`, `group`, `image`, `iWork item`, `line`, `movie`, `page`, `shape`, `table`, `text item`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `body text` | `rich text` | r/w | The section body text. |

### `page`

A page of the document.

*Plural:* `pages`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `body text` | `rich text` | r/w | The page body text. |

### `placeholder text`

One of some text's placeholders.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `tag` | `text` | r/w | Its script tag. |

### Enumerations

### `saveable file format`

| Value | Description |
|-------|-------------|
| `Pages Format` | The Pages native file format |

### `export format`

| Value | Description |
|-------|-------------|
| `EPUB` | EPUB |
| `unformatted text` | Plain text |
| `PDF` | PDF |
| `Microsoft Word` | Microsoft Word |
| `Pages 09` | Pages 09 |
| `formatted text` | RTF |

### `image quality`

| Value | Description |
|-------|-------------|
| `Good` | good quality |
| `Better` | better quality |
| `Best` | best quality |
