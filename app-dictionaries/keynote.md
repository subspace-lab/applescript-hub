# Keynote AppleScript Dictionary

> Auto-generated from `Keynote.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Keynote"`

## Table of Contents

- [Keynote Suite](#keynote-suite)
- [iWork Text Suite](#iwork-text-suite)
- [iWork Suite](#iwork-suite)
- [Compatibility Suite](#compatibility-suite)

---

## Keynote Suite

Classes and commands for Keynote.

### Commands

### `export`

Export a slideshow to another file

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | The slideshow to export |
| `to` | `file` | no | the destination file |
| `as` | `export format` | no | The format to use. |
| `with properties` | `export options` | yes | Optional export settings. |

### `duplicate`

Copy an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object(s) to copy. |
| `to` | `location specifier` | yes | The location for the new copy or copies. |
| `with properties` | `record` | yes | Properties to set in the new copy or copies right away. |

### `get`

Returns the value of the specified object(s).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

**Returns:** `any`

### Classes

### `slide layout`

A slide layout in a theme or slideshow document

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the slide layout |

### `slide`

A slide in a slideshow document

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `base layout` | `slide layout` | r/w | The slide layout this slide is based upon |
| `body showing` | `boolean` | r/w | Is the default body text displayed? |
| `skipped` | `boolean` | r/w | Is the slide skipped? |
| `slide number` | `integer` | r/o | index of the slide in the document |
| `title showing` | `boolean` | r/w | Is the default slide title displayed? |
| `default body item` | `shape` | r/o | The default body container of the slide |
| `default title item` | `shape` | r/o | The default title container of the slide |
| `presenter notes` | `rich text` | r/w | The presenter notes for the slide |
| `transition properties` | `transition settings` | r/w | The transition settings to apply to the slide. |

### `theme`

A collection of slide layouts, with shared design intents and elements.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The identifier used by the application. |
| `name` | `text` | r/o |  |

### Enumerations

### `saveable file format`

| Value | Description |
|-------|-------------|
| `Keynote` | The Keynote native file format |

### `export format`

| Value | Description |
|-------|-------------|
| `HTML` | HTML |
| `QuickTime movie` | QuickTime movie |
| `PDF` | PDF |
| `slide images` | image |
| `Microsoft PowerPoint` | Microsoft PowerPoint |
| `Keynote 09` | Keynote 09 |

### `image export formats`

| Value | Description |
|-------|-------------|
| `JPEG` | JPEG |
| `PNG` | PNG |
| `TIFF` | TIFF |

### `movie export formats`

| Value | Description |
|-------|-------------|
| `format360p` | 360p |
| `format540p` | 540p |
| `format720p` | 720p |
| `format1080p` | 1080p |
| `format2160p` | DCI 4K (4096x2160) |
| `native size` | Exported movie will have the same dimensions as the document, up to 4096x2160 |

### `movie codecs`

| Value | Description |
|-------|-------------|
| `h264` | H.264 |
| `AppleProRes422` | Apple ProRes 422 |
| `AppleProRes4444` | Apple ProRes 4444 |
| `AppleProRes422LT` | Apple ProRes 422LT |
| `AppleProRes422HQ` | Apple ProRes 422HQ |
| `AppleProRes422Proxy` | Apple ProRes 422Proxy |
| `HEVC` | HEVC |

### `movie framerates`

| Value | Description |
|-------|-------------|
| `FPS12` | 12 FPS |
| `FPS2398` | 23.98 FPS |
| `FPS24` | 24 FPS |
| `FPS25` | 25 FPS |
| `FPS2997` | 29.97 FPS |
| `FPS30` | 30 FPS |
| `FPS50` | 50 FPS |
| `FPS5994` | 59.94 FPS |
| `FPS60` | 60 FPS |

### `print what`

| Value | Description |
|-------|-------------|
| `IndividualSlides` | individual slides |
| `SlideWithNotes` | slides with notes |
| `Handouts` | handouts |

### `PDF image quality`

| Value | Description |
|-------|-------------|
| `Good` | good quality |
| `Better` | better quality |
| `Best` | best quality |

### `transition effects`

| Value | Description |
|-------|-------------|
| `no transition effect` |  |
| `magic move` |  |
| `shimmer` |  |
| `sparkle` |  |
| `swing` |  |
| `object cube` |  |
| `object flip` |  |
| `object pop` |  |
| `object push` |  |
| `object revolve` |  |
| `object zoom` |  |
| `perspective` |  |
| `clothesline` |  |
| `confetti` |  |
| `dissolve` |  |
| `drop` |  |
| `droplet` |  |
| `fade through color` |  |
| `grid` |  |
| `iris` |  |
| `move in` |  |
| `push` |  |
| `reveal` |  |
| `switch` |  |
| `wipe` |  |
| `blinds` |  |
| `color planes` |  |
| `cube` |  |
| `doorway` |  |
| `fall` |  |
| `flip` |  |
| `flop` |  |
| `mosaic` |  |
| `page flip` |  |
| `pivot` |  |
| `reflection` |  |
| `revolving door` |  |
| `scale` |  |
| `swap` |  |
| `swoosh` |  |
| `twirl` |  |
| `twist` |  |
| `fade and move` |  |

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

## Compatibility Suite

A set of basic classes for compatibility with prior releases.

### Commands

### `add chart`

Add a chart to a slide

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide` | no | the slide to add the chart to |
| `row names` | `text` | yes |  |
| `column names` | `text` | yes |  |
| `data` | `any` | yes |  |
| `type` | `legacy chart type` | yes |  |
| `group by` | `legacy chart grouping` | yes |  |

### `start`

Start playing the presentation.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | The presentation to play |
| `from` | `slide` | yes | slide at which to start playing |

### `make image slides`

Make a series of slides from a list of files.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no | the document to add the slides to |
| `files` | `file` | no | a list of image files to add |
| `set titles` | `boolean` | yes |  |
| `slide layout` | `slide layout` | yes |  |

### `stop`

Stop the presentation.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `show next`

Advance one build or slide.

### `show previous`

Go to the previous slide.

### `show slide switcher`

Show the slide switcher in play mode

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `hide slide switcher`

Hide the slide switcher in play mode

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `move slide switcher forward`

Move the slide switcher forward one slide

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `move slide switcher backward`

Move the slide switcher backward one slide

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `cancel slide switcher`

Hide the slide switcher without changing slides

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `accept slide switcher`

Hide the slide switcher, going to the slide it has selected

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `start slideshow`

### `start from`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide` | no | the object for the command |

### `stop slideshow`

### `show`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide` | no | the object for the command |

### Enumerations

### `legacy chart type`

Visual style of chart

| Value | Description |
|-------|-------------|
| `pie_2d` | two-dimensional pie chart |
| `vertical_bar_2d` | two-dimensional vertical bar chart |
| `stacked_vertical_bar_2d` | two-dimensional stacked vertical bar chart |
| `horizontal_bar_2d` | two-dimensional horizontal bar chart |
| `stacked_horizontal_bar_2d` | two-dimensional stacked horizontal bar chart |
| `pie_3d` | three-dimensional pie chart. |
| `vertical_bar_3d` | three-dimensional vertical bar chart |
| `stacked_vertical_bar_3d` | three-dimensional stacked bar chart |
| `horizontal_bar_3d` | three-dimensional horizontal bar chart |
| `stacked_horizontal_bar_3d` | three-dimensional stacked horizontal bar chart |
| `area_2d` | two-dimensional area chart. |
| `stacked_area_2d` | two-dimensional stacked area chart |
| `line_2d` | two-dimensional line chart. |
| `line_3d` | three-dimensional line chart |
| `area_3d` | three-dimensional area chart |
| `stacked_area_3d` | three-dimensional stacked area chart |
| `scatterplot_2d` | two-dimensional scatterplot chart |

### `legacy chart grouping`

Grouping for chart data

| Value | Description |
|-------|-------------|
| `chart row` | group by row |
| `chart column` | group by column |
