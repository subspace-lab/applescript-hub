# Word AppleScript Dictionary

> Auto-generated from `Word.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Word"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Miscellaneous Standards](#miscellaneous-standards)
- [Microsoft Office Suite](#microsoft-office-suite)
- [Microsoft Word Suite](#microsoft-word-suite)
- [Drawing Suite](#drawing-suite)
- [Text Suite](#text-suite)
- [Proofing Suite](#proofing-suite)
- [Table Suite](#table-suite)

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
| `file name` | `text` | yes | name of the file to open |
| `confirm conversions` | `boolean` | yes | Set to true to display the convert file dialog box if the file is not in Microsoft Word format |
| `read only` | `boolean` | yes | Set to true to open the file as read only |
| `add to recent files` | `boolean` | yes |  |
| `repair` | `boolean` | yes | Set to true to open the file in repair mode |
| `showing repairs` | `boolean` | yes | Set to false to suppress showing repairs when opening in repair mode |
| `password document` | `text` | yes | The password for opening the document |
| `password template` | `text` | yes | The password for opening the template |
| `revert` | `boolean` | yes | Controls what happens if the file is an open document.  Set to true to discard any unsaved changes.  False will activate the document |
| `write password` | `text` | yes | The password for saving changes to the document |
| `write password template` | `text` | yes | The password for saving changes to the template |
| `file converter` | `WdOpenFormat` | yes | The format for the file converter to use when opening this file. You can specify one of the enumerated constants or use the open format property (an integer) of the file converter object you want to use |

**Returns:** `document | document` — The opened document(s).

### Classes

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
| `document` | `document` | r/o | The document whose contents are displayed in the window. |
| `entry index` | `integer` | r/o | the number of the window |
| `position` | `point` | r/w | upper left coordinates of the window |
| `titled` | `boolean` | r/o | Does the window have a title bar? |
| `floating` | `boolean` | r/o | Does the window float? |
| `modal` | `boolean` | r/o | Is the window modal? |
| `collapsable` | `boolean` | r/o | Is the window collapasable? |
| `collapsed` | `boolean` | r/w | Is the window collapsed? |
| `sheet` | `boolean` | r/o | Is this window a sheet window? |

**Responds to:** `close`, `print`, `save`

---

## Miscellaneous Standards

Miscellaneous standard events and classes

### Commands

### `do script`

Execute a script

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The script to execute |

**Returns:** `any` — the result of the script

---

## Microsoft Office Suite

Common classes and commands used throughout all Office applications

### Commands

### `add item to combobox`

Add a new string to a custom combobox control.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `command bar combobox` | no |  |
| `combobox item` | `text` | no | The string value to be added. |
| `entry_index` | `integer` | yes | The index where the string is to be added. |

### `clear combobox`

Clear all of the strings form a custom combobox.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `command bar combobox` | no |  |

### `execute`

Runs the procedure or built-in command assigned to the specified command bar control.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `command bar control` | no |  |

### `get combobox item`

Return the string at the given index within a combobox.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `command bar combobox` | no |  |
| `entry_index` | `integer` | no | The index in the combobox where the string to be retrieved is located. |

**Returns:** `text` — The string associated with the referenced combobox item.

### `get count of combobox items`

Return the number of strings within a combobox.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `command bar combobox` | no |  |

**Returns:** `integer` — The number of strings in the combobox.

### `remove an item from combobox`

Remove a string from a custom combobox.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `command bar combobox` | no |  |
| `entry_index` | `integer` | no | The index in the combobox where the string to be removed is located. |

### `reset`

Resets a built-in command bar or command bar control to its default configuration.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4000` | no |  |

### `set combobox item`

Set the string an a given index for a custom combobox.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `command bar combobox` | no |  |
| `entry_index` | `integer` | no | The index where the string is to be set. |
| `combobox item` | `text` | no | The string value to be set. |

### Classes

### `command bar button`

A button control within a command bar.

*Plural:* `command bar buttons`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `button face is default` | `boolean` | r/o | Returns if the face of a command bar button control is the original built-in face. |
| `button state` | `MsoButtonState` | r/w | Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons. |
| `button style` | `MsoButtonStyle` | r/w | Returns or sets the way a command button control is displayed. |
| `face id` | `integer` | r/w | Returns or sets the Id number for the face of the command bar button control. |

### `command bar combobox`

A combobox menu control within a command bar.

*Plural:* `command bar comboboxes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `combobox style` | `MsoComboStyle` | r/w | Returns or sets the way a command bar combobox control is displayed. |
| `combobox text` | `text` | r/w | Returns or sets the text in the display or edit portion of the command bar combobox control. |
| `drop down lines` | `integer` | r/w | Returns or sets the number of lines in a command bar control combobox control.  The combobox control must be a custom control. |
| `drop down width` | `integer` | r/w | Returns or sets the width in pixels of the list for the specified command bar combobox control.  An error occurs if you attempt to set this property for a built-in combobox control. |
| `list index` | `integer` | r/w |  |

### `command bar control`

A control within a command bar.

*Plural:* `command bar controls`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `begin group` | `boolean` | r/w | Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar. |
| `built in` | `boolean` | r/o | Returns true if the command bar control is a built-in command bar control. |
| `control type` | `MsoControlType` | r/o | Returns the type of the command bar control. |
| `description text` | `text` | r/w | Returns or sets the description for a command bar control.  The description is not displayed to the user, but it can be useful for documenting the behavior of a control. |
| `enabled` | `boolean` | r/w | Returns or sets if the command bar control is enabled. |
| `entry_index` | `integer` | r/o | Returns the index number for this command bar control. |
| `height` | `integer` | r/w | Returns or sets the height of a command bar control. |
| `help context ID` | `integer` | r/w | Returns or sets the help context ID number for the Help topic attached to the command bar control. |
| `help file` | `text` | r/w | Returns or sets the file name for the help topic attached to the command bar.  To use this property, you must also set the help context ID property. |
| `id` | `integer` | r/o | Returns the id for a built-in command bar control. |
| `left position` | `integer` | r/o | Returns the left position of the command bar control. |
| `name` | `text` | r/w | Returns or sets the caption text for a command bar control. |
| `parameter` | `text` | r/w | Returns or sets a string that is used to execute a command. |
| `priority` | `integer` | r/w | Returns or sets the priority of a command bar control. A controls priority determines whether the control can be dropped from a docked command bar if the command bar controls can not fit in a single row.  Valid priority number are 0 through 7. |
| `tag` | `text` | r/w | Returns or sets information about the command bar control, such as data that can be used as an argument in procedures, or information that identifies the control. |
| `tooltip text` | `text` | r/w | Returns or sets the text displayed in a command bar controls tooltip. |
| `top` | `integer` | r/o | Returns the top position of a command bar control. |
| `visible` | `boolean` | r/w | Returns or sets if the command bar control is visible. |
| `width` | `integer` | r/w | Returns or sets the width in pixels of the command bar control. |

### `command bar popup`

A popup menu control within a command bar.

*Plural:* `command bar popups`

**Contains:** `command bar control`

### `command bar`

Toolbars used in all of the Office applications.

*Plural:* `command bars`

**Contains:** `command bar control`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bar position` | `MsoBarPosition` | r/w | Returns or sets the position of the command bar. |
| `bar type` | `MsoBarType` | r/o | Returns the type of this command bar. |
| `built in` | `boolean` | r/o | True if the command bar is built-in. |
| `context` | `text` | r/o | Returns or sets a string that determines where a command bar will be saved. |
| `enabled` | `boolean` | r/w | Returns or set if the command bar is enabled. |
| `entry_index` | `integer` | r/o | The index of the command bar. |
| `height` | `integer` | r/w | Returns or sets the height of the command bar. |
| `left position` | `integer` | r/w | Returns or sets the left position of the command bar. |
| `local name` | `text` | r/w | Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar. |
| `name` | `text` | r/w | Returns or sets the name of the command bar. |
| `protection` | `MsoBarProtection` | r/w | Returns or sets the way a command bar is protected from user customization.  It accepts a list of the following items: no protection, no customize, no resize, no move, no change visible, no change dock, no vertical dock, no horizontal dock. |
| `row index` | `integer` | r/w | Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero. |
| `top` | `integer` | r/w | Returns or sets the top position of a command bar. |
| `visible` | `boolean` | r/w | Returns or sets if the command bar is visible. |
| `width` | `integer` | r/w | Returns or sets the width in pixels of the command bar. |

### `custom document property`

*Plural:* `custom document properties`

### `document property`

*Plural:* `document properties`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `document property type` | `type` | r/w | Returns or sets the document property type. |
| `link source` | `text` | r/w | Returns or sets the source of a lined custom document property. |
| `link to content` | `boolean` | r/w | True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false. |
| `name` | `text` | r/w | Returns or sets the name of the document property. |
| `value` | `text` | r/w | Returns or sets the value of the document property. |

### `web page font`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `fixed width font` | `text` | r/w | Returns or sets the fixed-width font setting. |
| `fixed width font size` | `real` | r/w | Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point. |
| `proportional font` | `text` | r/w | Returns or sets the proportional font setting. |
| `proportional font size` | `real` | r/w | Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point. |

### Enumerations

### `MsoLineDashStyle`

| Value | Description |
|-------|-------------|
| `line dash style unset` |  |
| `line dash style solid` |  |
| `line dash style square dot` |  |
| `line dash style round dot` |  |
| `line dash style dash` |  |
| `line dash style dash dot` |  |
| `line dash style dash dot dot` |  |
| `line dash style long dash` |  |
| `line dash style long dash dot` |  |
| `line dash style long dash dot dot` |  |
| `line dash style system dash` |  |
| `line dash style system dot` |  |
| `line dash style system dash dot` |  |

### `MsoLineStyle`

| Value | Description |
|-------|-------------|
| `line style unset` |  |
| `single line` |  |
| `thin thin line` |  |
| `thin thick line` |  |
| `thick thin line` |  |
| `thick between thin line` |  |

### `MsoArrowheadStyle`

| Value | Description |
|-------|-------------|
| `arrowhead style unset` |  |
| `no arrowhead` |  |
| `triangle arrowhead` |  |
| `open_arrowhead` |  |
| `stealth arrowhead` |  |
| `diamond arrowhead` |  |
| `oval arrowhead` |  |

### `MsoArrowheadWidth`

| Value | Description |
|-------|-------------|
| `arrowhead width unset` |  |
| `narrow width arrowhead` |  |
| `medium width arrowhead` |  |
| `wide arrowhead` |  |

### `MsoArrowheadLength`

| Value | Description |
|-------|-------------|
| `arrowhead length unset` |  |
| `short arrowhead` |  |
| `medium arrowhead` |  |
| `long arrowhead` |  |

### `MsoFillType`

| Value | Description |
|-------|-------------|
| `fill unset` |  |
| `fill solid` |  |
| `fill patterned` |  |
| `fill gradient` |  |
| `fill textured` |  |
| `fill background` |  |
| `fill picture` |  |

### `MsoGradientStyle`

| Value | Description |
|-------|-------------|
| `gradient unset` |  |
| `horizontal gradient` |  |
| `vertical gradient` |  |
| `diagonal up gradient` |  |
| `diagonal down gradient` |  |
| `from corner gradient` |  |
| `from title gradient` |  |
| `from center gradient` |  |

### `MsoGradientColorType`

| Value | Description |
|-------|-------------|
| `gradient type unset` |  |
| `single shade gradient type` |  |
| `two colors gradient type` |  |
| `preset colors gradient type` |  |
| `multi colors gradient type` |  |

### `MsoTextureType`

| Value | Description |
|-------|-------------|
| `texture type texture type unset` |  |
| `texture type preset texture` |  |
| `texture type user defined texture` |  |

### `MsoPresetTexture`

| Value | Description |
|-------|-------------|
| `preset texture unset` |  |
| `texture papyrus` |  |
| `texture canvas` |  |
| `texture denim` |  |
| `texture woven mat` |  |
| `texture water droplets` |  |
| `texture paper bag` |  |
| `texture fish fossil` |  |
| `texture sand` |  |
| `texture green marble` |  |
| `texture white marble` |  |
| `texture brown marble` |  |
| `texture granite` |  |
| `texture newsprint` |  |
| `texture recycled paper` |  |
| `texture parchment` |  |
| `texture stationery` |  |
| `texture blue tissue paper` |  |
| `texture pink tissue paper` |  |
| `texture purple mesh` |  |
| `texture bouquet` |  |
| `texture cork` |  |
| `texture walnut` |  |
| `texture oak` |  |
| `texture medium wood` |  |

### `MsoPatternType`

| Value | Description |
|-------|-------------|
| `pattern unset` |  |
| `five percent pattern` |  |
| `ten percent pattern` |  |
| `twenty percent pattern` |  |
| `twenty five percent pattern` |  |
| `thirty percent pattern` |  |
| `forty percent pattern` |  |
| `fifty percent pattern` |  |
| `sixty percent pattern` |  |
| `seventy percent pattern` |  |
| `seventy five percent pattern` |  |
| `eighty percent pattern` |  |
| `ninety percent pattern` |  |
| `dark horizontal pattern` |  |
| `dark vertical pattern` |  |
| `dark downward diagonal pattern` |  |
| `dark upward diagonal pattern` |  |
| `small checker board pattern` |  |
| `trellis pattern` |  |
| `light horizontal pattern` |  |
| `light vertical pattern` |  |
| `light downward diagonal pattern` |  |
| `light upward diagonal pattern` |  |
| `small grid pattern` |  |
| `dotted diamond pattern` |  |
| `wide downward diagonal` |  |
| `wide upward diagonal pattern` |  |
| `dashed upward diagonal pattern` |  |
| `dashed downward diagonal pattern` |  |
| `narrow vertical pattern` |  |
| `narrow horizontal pattern` |  |
| `dashed vertical pattern` |  |
| `dashed horizontal pattern` |  |
| `large confetti pattern` |  |
| `large grid pattern` |  |
| `horizontal brick pattern` |  |
| `large checker board pattern` |  |
| `small confetti pattern` |  |
| `zig zag pattern` |  |
| `solid diamond pattern` |  |
| `diagonal brick pattern` |  |
| `outlined diamond pattern` |  |
| `plaid pattern` |  |
| `sphere pattern` |  |
| `weave pattern` |  |
| `dotted grid pattern` |  |
| `divot pattern` |  |
| `shingle pattern` |  |
| `wave pattern` |  |
| `horizontal pattern` |  |
| `vertical pattern` |  |
| `cross pattern` |  |
| `downward diagonal pattern` |  |
| `upward diagonal pattern` |  |
| `diagonal cross pattern` |  |

### `MsoPresetGradientType`

| Value | Description |
|-------|-------------|
| `preset gradient unset` |  |
| `gradient early sunset` |  |
| `gradient late sunset` |  |
| `gradient nightfall` |  |
| `gradient daybreak` |  |
| `gradient horizon` |  |
| `gradient desert` |  |
| `gradient ocean` |  |
| `gradient calm water` |  |
| `gradient fire` |  |
| `gradient fog` |  |
| `gradient moss` |  |
| `gradient peacock` |  |
| `gradient wheat` |  |
| `gradient parchment` |  |
| `gradient mahogany` |  |
| `gradient rainbow` |  |
| `gradient rainbow2` |  |
| `gradient gold` |  |
| `gradient gold2` |  |
| `gradient brass` |  |
| `gradient chrome` |  |
| `gradient chrome2` |  |
| `gradient silver` |  |
| `gradient sapphire` |  |

### `MsoShadowType`

| Value | Description |
|-------|-------------|
| `shadow unset` |  |
| `shadow1` |  |
| `shadow2` |  |
| `shadow3` |  |
| `shadow4` |  |
| `shadow5` |  |
| `shadow6` |  |
| `shadow7` |  |
| `shadow8` |  |
| `shadow9` |  |
| `shadow10` |  |
| `shadow11` |  |
| `shadow12` |  |
| `shadow13` |  |
| `shadow14` |  |
| `shadow15` |  |
| `shadow16` |  |
| `shadow17` |  |
| `shadow18` |  |
| `shadow19` |  |
| `shadow20` |  |
| `shadow21` |  |
| `shadow22` |  |
| `shadow23` |  |
| `shadow24` |  |
| `shadow25` |  |
| `shadow26` |  |
| `shadow27` |  |
| `shadow28` |  |
| `shadow29` |  |
| `shadow30` |  |
| `shadow31` |  |
| `shadow32` |  |
| `shadow33` |  |
| `shadow34` |  |
| `shadow35` |  |
| `shadow36` |  |
| `shadow37` |  |
| `shadow38` |  |
| `shadow39` |  |
| `shadow40` |  |
| `shadow41` |  |
| `shadow42` |  |
| `shadow43` |  |

### `MsoPresetTextEffect`

| Value | Description |
|-------|-------------|
| `wordart format unset` |  |
| `wordart format1` |  |
| `wordart format2` |  |
| `wordart format3` |  |
| `wordart format4` |  |
| `wordart format5` |  |
| `wordart format6` |  |
| `wordart format7` |  |
| `wordart format8` |  |
| `wordart format9` |  |
| `wordart format10` |  |
| `wordart format11` |  |
| `wordart format12` |  |
| `wordart format13` |  |
| `wordart format14` |  |
| `wordart format15` |  |
| `wordart format16` |  |
| `wordart format17` |  |
| `wordart format18` |  |
| `wordart format19` |  |
| `wordart format20` |  |
| `wordart format21` |  |
| `wordart format22` |  |
| `wordart format23` |  |
| `wordart format24` |  |
| `wordart format25` |  |
| `wordart format26` |  |
| `wordart format27` |  |
| `wordart format28` |  |
| `wordart format29` |  |
| `wordart format30` |  |
| `wordart format31` |  |
| `wordart format32` |  |
| `wordart format33` |  |
| `wordart format34` |  |
| `wordart format35` |  |
| `wordart format36` |  |
| `wordart format37` |  |
| `wordart format38` |  |
| `wordart format39` |  |
| `wordart format40` |  |
| `wordart format41` |  |
| `wordart format42` |  |
| `wordart format43` |  |
| `wordart format44` |  |
| `wordart format45` |  |
| `wordart format46` |  |
| `wordart format47` |  |
| `wordart format48` |  |
| `wordart format49` |  |
| `wordart format50` |  |

### `MsoPresetTextEffectShape`

| Value | Description |
|-------|-------------|
| `text effect shape unset` |  |
| `plain text` |  |
| `stop` |  |
| `triangle up` |  |
| `triangle down` |  |
| `chevron up` |  |
| `chevron down` |  |
| `ring inside` |  |
| `ring outside` |  |
| `arch up curve` |  |
| `arch down curve` |  |
| `circle curve` |  |
| `button curve` |  |
| `arch up pour` |  |
| `arch down pour` |  |
| `circle pour` |  |
| `button pour` |  |
| `curve up` |  |
| `curve down` |  |
| `can up` |  |
| `can down` |  |
| `wave1` |  |
| `wave2` |  |
| `double wave1` |  |
| `double wave2` |  |
| `inflate` |  |
| `deflate` |  |
| `inflate bottom` |  |
| `deflate bottom` |  |
| `inflate top` |  |
| `deflate top` |  |
| `deflate inflate` |  |
| `deflate inflate deflate` |  |
| `fade right` |  |
| `fade left` |  |
| `fade up` |  |
| `fade down` |  |
| `slant up` |  |
| `slant down` |  |
| `cascade up` |  |
| `cascade down` |  |

### `MsoTextEffectAlignment`

| Value | Description |
|-------|-------------|
| `text effect alignment unset` |  |
| `left text effect alignment` |  |
| `centered text effect alignment` |  |
| `right text effect alignment` |  |
| `justify text effect alignment` |  |
| `word justify text effect alignment` |  |
| `stretch justify text effect alignment` |  |

### `MsoPresetLightingDirection`

| Value | Description |
|-------|-------------|
| `preset lighting direction unset` |  |
| `light from top left` |  |
| `light from top` |  |
| `light from top right` |  |
| `light from left` |  |
| `light from none` |  |
| `light from right` |  |
| `light from bottom left` |  |
| `light from bottom` |  |
| `light from bottom right` |  |

### `MsoPresetLightingSoftness`

| Value | Description |
|-------|-------------|
| `lighting softness unset` |  |
| `lighting dim` |  |
| `lighting normal` |  |
| `lighting bright` |  |

### `MsoPresetMaterial`

| Value | Description |
|-------|-------------|
| `preset material unset` |  |
| `matte` |  |
| `plastic` |  |
| `metal` |  |
| `wireframe` |  |
| `matte2` |  |
| `plastic2` |  |
| `metal2` |  |
| `warm matte` |  |
| `translucent powder` |  |
| `powder` |  |
| `dark edge` |  |
| `soft edge` |  |
| `material clear` |  |
| `flat` |  |
| `soft metal` |  |

### `MsoPresetExtrusionDirection`

| Value | Description |
|-------|-------------|
| `preset extrusion direction unset` |  |
| `extrude bottom right` |  |
| `extrude bottom` |  |
| `extrude bottom left` |  |
| `extrude right` |  |
| `extrude none` |  |
| `extrude left` |  |
| `extrude top right` |  |
| `extrude top` |  |
| `extrude top left` |  |

### `MsoPresetThreeDFormat`

| Value | Description |
|-------|-------------|
| `preset threeD format unset` |  |
| `format1` |  |
| `format2` |  |
| `format3` |  |
| `format4` |  |
| `format5` |  |
| `format6` |  |
| `format7` |  |
| `format8` |  |
| `format9` |  |
| `format10` |  |
| `format11` |  |
| `format12` |  |
| `format13` |  |
| `format14` |  |
| `format15` |  |
| `format16` |  |
| `format17` |  |
| `format18` |  |
| `format19` |  |
| `format20` |  |

### `MsoExtrusionColorType`

| Value | Description |
|-------|-------------|
| `extrusion color unset` |  |
| `extrusion color automatic` |  |
| `extrusion color custom` |  |

### `MsoConnectorType`

| Value | Description |
|-------|-------------|
| `connector type unset` |  |
| `straight` |  |
| `elbow` |  |
| `curve` |  |

### `MsoHorizontalAnchor`

| Value | Description |
|-------|-------------|
| `horizontal anchor unset` |  |
| `horizontal anchor none` |  |
| `horizontal anchor center` |  |

### `MsoVerticalAnchor`

| Value | Description |
|-------|-------------|
| `vertical anchor unset` |  |
| `anchor top` |  |
| `anchor top baseline` |  |
| `anchor middle` |  |
| `anchor bottom` |  |
| `anchor bottom baseline` |  |

### `MsoAutoShapeType`

| Value | Description |
|-------|-------------|
| `autoshape shape type unset` |  |
| `autoshape rectangle` |  |
| `autoshape parallelogram` |  |
| `autoshape trapezoid` |  |
| `autoshape diamond` |  |
| `autoshape rounded rectangle` |  |
| `autoshape octagon` |  |
| `autoshape isosceles triangle` |  |
| `autoshape right triangle` |  |
| `autoshape oval` |  |
| `autoshape hexagon` |  |
| `autoshape cross` |  |
| `autoshape regular pentagon` |  |
| `autoshape can` |  |
| `autoshape cube` |  |
| `autoshape bevel` |  |
| `autoshape folded corner` |  |
| `autoshape smiley face` |  |
| `autoshape donut` |  |
| `autoshape no symbol` |  |
| `autoshape block arc` |  |
| `autoshape heart` |  |
| `autoshape lightning bolt` |  |
| `autoshape sun` |  |
| `autoshape moon` |  |
| `autoshape arc` |  |
| `autoshape double bracket` |  |
| `autoshape double brace` |  |
| `autoshape plaque` |  |
| `autoshape left bracket` |  |
| `autoshape right bracket` |  |
| `autoshape left brace` |  |
| `autoshape right brace` |  |
| `autoshape right arrow` |  |
| `autoshape left arrow` |  |
| `autoshape up arrow` |  |
| `autoshape down arrow` |  |
| `autoshape left right arrow` |  |
| `autoshape up down arrow` |  |
| `autoshape quad arrow` |  |
| `autoshape left right up arrow` |  |
| `autoshape bent arrow` |  |
| `autoshape U turn arrow` |  |
| `autoshape left up arrow` |  |
| `autoshape bent up arrow` |  |
| `autoshape curved right arrow` |  |
| `autoshape curved left arrow` |  |
| `autoshape curved up arrow` |  |
| `autoshape curved down arrow` |  |
| `autoshape striped right arrow` |  |
| `autoshape notched right arrow` |  |
| `autoshape pentagon` |  |
| `autoshape chevron` |  |
| `autoshape right arrow callout` |  |
| `autoshape left arrow callout` |  |
| `autoshape up arrow callout` |  |
| `autoshape down arrow callout` |  |
| `autoshape left right arrow callout` |  |
| `autoshape up down arrow callout` |  |
| `autoshape quad arrow callout` |  |
| `autoshape circular arrow` |  |
| `autoshape flowchart process` |  |
| `autoshape flowchart alternate process` |  |
| `autoshape flowchart decision` |  |
| `autoshape flowchart data` |  |
| `autoshape flowchart predefined process` |  |
| `autoshape flowchart internal storage` |  |
| `autoshape flowchart document` |  |
| `autoshape flowchart multi document` |  |
| `autoshape flowchart terminator` |  |
| `autoshape flowchart preparation` |  |
| `autoshape flowchart manual input` |  |
| `autoshape flowchart manual operation` |  |
| `autoshape flowchart connector` |  |
| `autoshape flowchart offpage connector` |  |
| `autoshape flowchart card` |  |
| `autoshape flowchart punched tape` |  |
| `autoshape flowchart summing junction` |  |
| `autoshape flowchart or` |  |
| `autoshape flowchart collate` |  |
| `autoshape flowchart sort` |  |
| `autoshape flowchart extract` |  |
| `autoshape flowchart merge` |  |
| `autoshape flowchart stored data` |  |
| `autoshape flowchart delay` |  |
| `autoshape flowchart sequential access storage` |  |
| `autoshape flowchart magnetic disk` |  |
| `autoshape flowchart direct access storage` |  |
| `autoshape flowchart display` |  |
| `autoshape explosion one` |  |
| `autoshape explosion two` |  |
| `autoshape four point star` |  |
| `autoshape five point star` |  |
| `autoshape eight point star` |  |
| `autoshape sixteen point star` |  |
| `autoshape twenty four point star` |  |
| `autoshape thirty two point star` |  |
| `autoshape up ribbon` |  |
| `autoshape down ribbon` |  |
| `autoshape curved up ribbon` |  |
| `autoshape curved down ribbon` |  |
| `autoshape vertical scroll` |  |
| `autoshape horizontal scroll` |  |
| `autoshape wave` |  |
| `autoshape double wave` |  |
| `autoshape rectangular callout` |  |
| `autoshape rounded rectangular callout` |  |
| `autoshape oval callout` |  |
| `autoshape cloud callout` |  |
| `autoshape line callout one` |  |
| `autoshape line callout two` |  |
| `autoshape line callout three` |  |
| `autoshape line callout four` |  |
| `autoshape line callout one accent bar` |  |
| `autoshape line callout two accent bar` |  |
| `autoshape line callout three accent bar` |  |
| `autoshape line callout four accent bar` |  |
| `autoshape line callout one no border` |  |
| `autoshape line callout two no border` |  |
| `autoshape line callout three no border` |  |
| `autoshape line callout four no border` |  |
| `autoshape callout one border and accent bar` |  |
| `autoshape callout two border and accent bar` |  |
| `autoshape callout three border and accent bar` |  |
| `autoshape callout four border and accent bar` |  |
| `autoshape action button custom` |  |
| `autoshape action button home` |  |
| `autoshape action button help` |  |
| `autoshape action button information` |  |
| `autoshape action button back or previous` |  |
| `autoshape action button forward or next` |  |
| `autoshape action button beginning` |  |
| `autoshape action button end` |  |
| `autoshape action button return` |  |
| `autoshape action button document` |  |
| `autoshape action button sound` |  |
| `autoshape action button movie` |  |
| `autoshape balloon` |  |
| `autoshape not primitive` |  |
| `autoshape flowchart offline storage` |  |
| `autoshape left right ribbon` |  |
| `autoshape diagonal stripe` |  |
| `autoshape pie` |  |
| `autoshape non isosceles trapezoid` |  |
| `autoshape Decagon` |  |
| `autoshape Heptagon` |  |
| `autoshape Dodecagon` |  |
| `autoshape six points star` |  |
| `autoshape seven points star` |  |
| `autoshape ten points star` |  |
| `autoshape twelve points star` |  |
| `autoshape round one rectangle` |  |
| `autoshape round two same rectangle` |  |
| `autoshape round two diagonal rectangle` |  |
| `autoshape snip round rectangle` |  |
| `autoshape snip one rectangle` |  |
| `autoshape snip two same rectangle` |  |
| `autoshape snip two diagonal rectangle` |  |
| `autoshape frame` |  |
| `autoshape half frame` |  |
| `autoshape tear` |  |
| `autoshape chord` |  |
| `autoshape corner` |  |
| `autoshape math plus` |  |
| `autoshape math minus` |  |
| `autoshape math multiply` |  |
| `autoshape math divide` |  |
| `autoshape math equal` |  |
| `autoshape math not equal` |  |
| `autoshape corner tabs` |  |
| `autoshape square tabs` |  |
| `autoshape plaque tabs` |  |
| `autoshape gear six` |  |
| `autoshape gear nine` |  |
| `autoshape funnel` |  |
| `autoshape pie wedge` |  |
| `autoshape left circular arrow` |  |
| `autoshape left right circular arrow` |  |
| `autoshape swoosh arrow` |  |
| `autoshape cloud` |  |
| `autoshape chart x` |  |
| `autoshape chart star` |  |
| `autoshape chart plus` |  |
| `autoshape line inverse` |  |

### `MsoShapeType`

| Value | Description |
|-------|-------------|
| `shape type unset` |  |
| `shape type auto` |  |
| `shape type callout` |  |
| `shape type chart` |  |
| `shape type comment` |  |
| `shape type free form` |  |
| `shape type group` |  |
| `shape type embedded OLE control` |  |
| `shape type form control` |  |
| `shape type line` |  |
| `shape type linked OLE object` |  |
| `shape type linked picture` |  |
| `shape type OLE control` |  |
| `shape type picture` |  |
| `shape type place holder` |  |
| `shape type word art` |  |
| `shape type media` |  |
| `shape type text box` |  |
| `shape type script anchor` |  |
| `shape type table` |  |
| `shape type canvas` |  |
| `shape type diagram` |  |
| `shape type ink` |  |
| `shape type ink comment` |  |
| `shape type smartart graphic` |  |
| `shape type slicer` |  |
| `shape type web video` |  |
| `shape type content application` |  |
| `shape type graphic` |  |
| `shape type linked graphic` |  |
| `shape type 3d model` |  |
| `shape type linked 3d model` |  |

### `MsoColorType`

| Value | Description |
|-------|-------------|
| `color type unset` |  |
| `RGB` |  |
| `Scheme` |  |

### `MsoPictureColorType`

| Value | Description |
|-------|-------------|
| `picture color type unset` |  |
| `picture color automatic` |  |
| `picture color gray scale` |  |
| `picture color black and white` |  |
| `picture color watermark` |  |

### `MsoCalloutAngleType`

| Value | Description |
|-------|-------------|
| `angle unset` |  |
| `angle automatic` |  |
| `angle30` |  |
| `angle45` |  |
| `angle60` |  |
| `angle90` |  |

### `MsoCalloutDropType`

| Value | Description |
|-------|-------------|
| `drop unset` |  |
| `drop custom` |  |
| `drop top` |  |
| `drop center` |  |
| `drop bottom` |  |

### `MsoCalloutType`

| Value | Description |
|-------|-------------|
| `callout unset` |  |
| `callout one` |  |
| `callout two` |  |
| `callout three` |  |
| `callout four` |  |

### `MsoTextOrientation`

| Value | Description |
|-------|-------------|
| `text orientation unset` |  |
| `horizontal` |  |
| `upward` |  |
| `downward` |  |
| `vertical east asian` |  |
| `vertical` |  |
| `horizontal rotated east asian` |  |

### `MsoScaleFrom`

| Value | Description |
|-------|-------------|
| `scale from top left` |  |
| `scale from middle` |  |
| `scale from bottom right` |  |

### `MsoPresetCamera`

| Value | Description |
|-------|-------------|
| `preset camera unset` |  |
| `camera legacy oblique from top left` |  |
| `camera legacy oblique from top` |  |
| `camera legacy oblique from topright` |  |
| `camera legacy oblique from left` |  |
| `camera legacy oblique from front` |  |
| `camera legacy oblique from right` |  |
| `camera legacy oblique from bottom left` |  |
| `camera legacy oblique from bottom` |  |
| `camera legacy oblique from bottom right` |  |
| `camera legacy perspective from top left` |  |
| `camera legacy perspective from top` |  |
| `camera legacy perspective from top right` |  |
| `camera legacy perspective from left` |  |
| `camera legacy perspective from front` |  |
| `camera legacy perspective from right` |  |
| `camera legacy perspective from bottom left` |  |
| `camera legacy perspective from bottom` |  |
| `camera legacy perspective from bottom right` |  |
| `camera orthographic` |  |
| `camera isometric from top up` |  |
| `camera isometric from top down` |  |
| `camera isometric from bottom up` |  |
| `camera isometric from bottom down` |  |
| `camera isometric from left up` |  |
| `camera isometric from left down` |  |
| `camera isometric from right up` |  |
| `camera isometric from right down` |  |
| `camera isometric off axis1 from left` |  |
| `camera isometric off axis1 from right` |  |
| `camera isometric off axis1 from top` |  |
| `camera isometric off axis2 from left` |  |
| `camera isometric off axis2 from right` |  |
| `camera isometric off axis2 from top` |  |
| `camera isometric off axis3 from left` |  |
| `camera isometric off axis3 from right` |  |
| `camera isometric off axis3 from bottom` |  |
| `camera isometric off axis4 from left` |  |
| `camera isometric off axis4 from right` |  |
| `camera isometric off axis4 from bottom` |  |
| `camera oblique from top left` |  |
| `camera oblique from top` |  |
| `camera oblique from top right` |  |
| `camera oblique from left` |  |
| `camera oblique from right` |  |
| `camera oblique from bottom left` |  |
| `camera oblique from bottom` |  |
| `camera oblique from bottom right` |  |
| `camera perspective from front` |  |
| `camera perspective from left` |  |
| `camera perspective from right` |  |
| `camera perspective from above` |  |
| `camera perspective from below` |  |
| `camera perspective from above facing left` |  |
| `camera perspective from above facing right` |  |
| `camera perspective contrasting facing left` |  |
| `camera perspective contrasting facing right` |  |
| `camera perspective heroic facing left` |  |
| `camera perspective heroic facing right` |  |
| `camera perspective heroic extreme facing left` |  |
| `camera perspective heroic extreme facing right` |  |
| `camera perspective relaxed` |  |
| `camera perspective relaxed moderately` |  |

### `MsoLightRigType`

| Value | Description |
|-------|-------------|
| `light rig unset` |  |
| `light rig flat1` |  |
| `light rig flat2` |  |
| `light rig flat3` |  |
| `light rig flat4` |  |
| `light rig Normal1` |  |
| `light rig Normal2` |  |
| `light rig Normal3` |  |
| `light rig Normal4` |  |
| `light rig Harsh1` |  |
| `light rig Harsh2` |  |
| `light rig Harsh3` |  |
| `light rig Harsh4` |  |
| `light rig three point` |  |
| `light rig balanced` |  |
| `light rig soft` |  |
| `light rig harsh` |  |
| `light rig flood` |  |
| `light rig contrasting` |  |
| `light rig morning` |  |
| `light rig sunrise` |  |
| `light rig sunset` |  |
| `light rig chilly` |  |
| `light rig freezing` |  |
| `light rig flat` |  |
| `light rig two point` |  |
| `light rig glow` |  |
| `light rig bright room` |  |

### `MsoBevelType`

| Value | Description |
|-------|-------------|
| `bevel type unset` |  |
| `bevel none` |  |
| `bevel relaxed inset` |  |
| `bevel circle` |  |
| `bevel slope` |  |
| `bevel cross` |  |
| `bevel angle` |  |
| `bevel soft round` |  |
| `bevel convex` |  |
| `bevel cool slant` |  |
| `bevel divot` |  |
| `bevel riblet` |  |
| `bevel hard edge` |  |
| `bevel art deco` |  |

### `MsoShadowStyle`

| Value | Description |
|-------|-------------|
| `shadow style unset` |  |
| `shadow style inner` |  |
| `shadow style outer` |  |

### `MsoParagraphAlignment`

| Value | Description |
|-------|-------------|
| `paragraph alignment unset` |  |
| `paragraph align left` |  |
| `paragraph align center` |  |
| `paragraph align right` |  |
| `paragraph align justify` |  |
| `paragraph align distribute` |  |
| `paragraph align Thai` |  |
| `paragraph align justify low` |  |

### `MsoTextStrike`

| Value | Description |
|-------|-------------|
| `strike unset` |  |
| `no strike` |  |
| `single strike` |  |
| `double strike` |  |

### `MsoTextCaps`

| Value | Description |
|-------|-------------|
| `caps unset` |  |
| `no caps` |  |
| `small caps` |  |
| `all caps` |  |

### `MsoTextUnderlineType`

| Value | Description |
|-------|-------------|
| `underline unset` |  |
| `no underline` |  |
| `underline words only` |  |
| `underline single line` |  |
| `underline double line` |  |
| `underline heavy line` |  |
| `underline dotted line` |  |
| `underline heavy dotted line` |  |
| `underline dash line` |  |
| `underline heavy dash line` |  |
| `underline long dash line` |  |
| `underline heavy long dash line` |  |
| `underline dot dash line` |  |
| `underline heavy dot dash line` |  |
| `underline dot dot dash line` |  |
| `underline heavy dot dot dash line` |  |
| `underline wavy line` |  |
| `underline heavy wavy line` |  |
| `underline wavy double line` |  |

### `MsoTextTabAlign`

| Value | Description |
|-------|-------------|
| `tab unset` |  |
| `left tab` |  |
| `center tab` |  |
| `right tab` |  |
| `decimal tab` |  |

### `MsoTextCharWrap`

| Value | Description |
|-------|-------------|
| `character wrap unset` |  |
| `no character wrap` |  |
| `standard character wrap` |  |
| `strict character wrap` |  |
| `custom character wrap` |  |

### `MsoTextFontAlign`

| Value | Description |
|-------|-------------|
| `font alignment unset` |  |
| `automatic alignment` |  |
| `top alignment` |  |
| `center alignment` |  |
| `baseline alignment` |  |
| `bottom alignment` |  |

### `MsoAutoSize`

| Value | Description |
|-------|-------------|
| `auto size unset` |  |
| `auto size none` |  |
| `shape to fit text` |  |
| `text to fit shape` |  |

### `MsoPathFormat`

| Value | Description |
|-------|-------------|
| `path type unset` |  |
| `no path type` |  |
| `path type1` |  |
| `path type2` |  |
| `path type3` |  |
| `path type4` |  |

### `MsoWarpFormat`

| Value | Description |
|-------|-------------|
| `warp format unset` |  |
| `warp format1` |  |
| `warp format2` |  |
| `warp format3` |  |
| `warp format4` |  |
| `warp format5` |  |
| `warp format6` |  |
| `warp format7` |  |
| `warp format8` |  |
| `warp format9` |  |
| `warp format10` |  |
| `warp format11` |  |
| `warp format12` |  |
| `warp format13` |  |
| `warp format14` |  |
| `warp format15` |  |
| `warp format16` |  |
| `warp format17` |  |
| `warp format18` |  |
| `warp format19` |  |
| `warp format20` |  |
| `warp format21` |  |
| `warp format22` |  |
| `warp format23` |  |
| `warp format24` |  |
| `warp format25` |  |
| `warp format26` |  |
| `warp format27` |  |
| `warp format28` |  |
| `warp format29` |  |
| `warp format30` |  |
| `warp format31` |  |
| `warp format32` |  |
| `warp format33` |  |
| `warp format34` |  |
| `warp format35` |  |
| `warp format36` |  |
| `warp format37` |  |

### `MsoTextChangeCase`

| Value | Description |
|-------|-------------|
| `case sentence` |  |
| `case lower` |  |
| `case upper` |  |
| `case title` |  |
| `case toggle` |  |

### `MsoDateTimeFormat`

| Value | Description |
|-------|-------------|
| `date time format unset` |  |
| `date time format Mdyy` |  |
| `date time format ddddMMMMddyyyy` |  |
| `date time format dMMMMyyyy` |  |
| `date time format MMMMdyyyy` |  |
| `date time format dMMMyy` |  |
| `date time format MMMMyy` |  |
| `date time format MMyy` |  |
| `date time format MMddyyHmm` |  |
| `date time format MMddyyhmmAMPM` |  |
| `date time format Hmm` |  |
| `date time format Hmmss` |  |
| `date time format hmmAMPM` |  |
| `date time format hmmssAMPM` |  |
| `date time format figure out` |  |

### `MsoSoftEdgeType`

| Value | Description |
|-------|-------------|
| `soft edge unset` |  |
| `no soft edge` |  |
| `soft edge type1` |  |
| `soft edge type2` |  |
| `soft edge type3` |  |
| `soft edge type4` |  |
| `soft edge type5` |  |
| `soft edge type6` |  |

### `MsoThemeColorSchemeIndex`

| Value | Description |
|-------|-------------|
| `first dark scheme color` |  |
| `first light scheme color` |  |
| `second dark scheme color` |  |
| `second light scheme color` |  |
| `first accent scheme color` |  |
| `second accent scheme color` |  |
| `third accent scheme color` |  |
| `fourth accent scheme color` |  |
| `fifth accent scheme color` |  |
| `sixth accent scheme color` |  |
| `hyperlink scheme color` |  |
| `followed hyperlink scheme color` |  |

### `MsoThemeColorIndex`

| Value | Description |
|-------|-------------|
| `theme color unset` |  |
| `no theme color` |  |
| `first dark theme color` |  |
| `first light theme color` |  |
| `second dark theme color` |  |
| `second light theme color` |  |
| `first accent theme color` |  |
| `second accent theme color` |  |
| `third accent theme color` |  |
| `fourth accent theme color` |  |
| `fifth accent theme color` |  |
| `sixth accent theme color` |  |
| `hyperlink theme color` |  |
| `followed hyperlink theme color` |  |
| `first text theme color` |  |
| `first background theme color` |  |
| `second text theme color` |  |
| `second background theme color` |  |

### `MsoFontLanguageIndex`

| Value | Description |
|-------|-------------|
| `theme font latin` |  |
| `theme font complex script` |  |
| `theme font high ansi` |  |
| `theme font east asian` |  |

### `MsoShapeStyleIndex`

| Value | Description |
|-------|-------------|
| `shape style unset` |  |
| `shape not a preset` |  |
| `shape preset1` |  |
| `shape preset2` |  |
| `shape preset3` |  |
| `shape preset4` |  |
| `shape preset5` |  |
| `shape preset6` |  |
| `shape preset7` |  |
| `shape preset8` |  |
| `shape preset9` |  |
| `shape preset10` |  |
| `shape preset11` |  |
| `shape preset12` |  |
| `shape preset13` |  |
| `shape preset14` |  |
| `shape preset15` |  |
| `shape preset16` |  |
| `shape preset17` |  |
| `shape preset18` |  |
| `shape preset19` |  |
| `shape preset20` |  |
| `shape preset21` |  |
| `shape preset22` |  |
| `shape preset23` |  |
| `shape preset24` |  |
| `shape preset25` |  |
| `shape preset26` |  |
| `shape preset27` |  |
| `shape preset28` |  |
| `shape preset29` |  |
| `shape preset30` |  |
| `shape preset31` |  |
| `shape preset32` |  |
| `shape preset33` |  |
| `shape preset34` |  |
| `shape preset35` |  |
| `shape preset36` |  |
| `shape preset37` |  |
| `shape preset38` |  |
| `shape preset39` |  |
| `shape preset40` |  |
| `shape preset41` |  |
| `shape preset42` |  |
| `line preset1` |  |
| `line preset2` |  |
| `line preset3` |  |
| `line preset4` |  |
| `line preset5` |  |
| `line preset6` |  |
| `line preset7` |  |
| `line preset8` |  |
| `line preset9` |  |
| `line preset10` |  |
| `line preset11` |  |
| `line preset12` |  |
| `line preset13` |  |
| `line preset14` |  |
| `line preset15` |  |
| `line preset16` |  |
| `line preset17` |  |
| `line preset18` |  |
| `line preset19` |  |
| `line preset20` |  |
| `line preset21` |  |

### `MsoBackgroundStyleIndex`

| Value | Description |
|-------|-------------|
| `background unset` |  |
| `background not a preset` |  |
| `background preset1` |  |
| `background preset2` |  |
| `background preset3` |  |
| `background preset4` |  |
| `background preset5` |  |
| `background preset6` |  |
| `background preset7` |  |
| `background preset8` |  |
| `background preset9` |  |
| `background preset10` |  |
| `background preset11` |  |
| `background preset12` |  |

### `MsoTextDirection`

| Value | Description |
|-------|-------------|
| `text direction unset` |  |
| `left to right` |  |
| `right to left` |  |

### `MsoBulletType`

| Value | Description |
|-------|-------------|
| `bullet type unset` |  |
| `bullet type none` |  |
| `bullet type unnumbered` |  |
| `bullet type numbered` |  |
| `picture bullet type` |  |

### `MsoNumberedBulletStyle`

| Value | Description |
|-------|-------------|
| `numbered bullet style unset` |  |
| `numbered bullet style alpha lowercase period` |  |
| `numbered bullet style alpha uppercase period` |  |
| `numbered bullet style arabic right paren` |  |
| `numbered bullet style arabic period` |  |
| `numbered bullet style roman lowercase paren both` |  |
| `numbered bullet style roman lowercase paren right` |  |
| `numbered bullet style roman lowercase period` |  |
| `numbered bullet style roman uppercase period` |  |
| `numbered bullet style alpha lowercase paren both` |  |
| `numbered bullet style alpha lowercase paren right` |  |
| `numbered bullet style alpha uppercase paren both` |  |
| `numbered bullet style alpha uppercase paren right` |  |
| `numbered bullet style arabic paren both` |  |
| `numbered bullet style arabic plain` |  |
| `numbered bullet style roman uppercase paren both` |  |
| `numbered bullet style roman uppercase paren right` |  |
| `numbered bullet style simplified chinese plain` |  |
| `numbered bullet style simplified chinese period` |  |
| `numbered bullet style circle number plain` |  |
| `numbered bullet style circle number white plain` |  |
| `numbered bullet style circle number black plain` |  |
| `numbered bullet style traditional chinese plain` |  |
| `numbered bullet style traditional chinese period` |  |
| `numbered bullet style arabic alpha dash` |  |
| `numbered bullet style arabic abjad dash` |  |
| `numbered bullet style hebrew alpha dash` |  |
| `numbered bullet style kanji korean plain` |  |
| `numbered bullet style kanji korean period` |  |
| `numbered bullet style arabic DB plain` |  |
| `numbered bullet style arabic DB period` |  |
| `numbered bullet style thai alpha period` |  |
| `numbered bullet style thai alpha paren right` |  |
| `numbered bullet style thai alpha paren both` |  |
| `numbered bullet style thai number period` |  |
| `numbered bullet style thai number paren right` |  |
| `numbered bullet style thai paren both` |  |
| `numbered bullet style hindi alpha period` |  |
| `numbered bullet style hindi number period` |  |
| `numbered bullet style kanji simpified chinese DB period` |  |
| `numbered bullet style hindi number paren right` |  |
| `numbered bullet style hindi alpha1 period` |  |

### `MsoTabStopType`

| Value | Description |
|-------|-------------|
| `tabstop unset` |  |
| `tabstop left` |  |
| `tabstop center` |  |
| `tabstop right` |  |
| `tabstop decimal` |  |

### `MsoReflectionType`

| Value | Description |
|-------|-------------|
| `reflection unset` |  |
| `reflection type none` |  |
| `reflection type1` |  |
| `reflection type2` |  |
| `reflection type3` |  |
| `reflection type4` |  |
| `reflection type5` |  |
| `reflection type6` |  |
| `reflection type7` |  |
| `reflection type8` |  |
| `reflection type9` |  |

### `MsoTextureAlignment`

| Value | Description |
|-------|-------------|
| `texture unset` |  |
| `texture top left` |  |
| `texture top` |  |
| `texture top right` |  |
| `texture left` |  |
| `texture center` |  |
| `texture right` |  |
| `texture bottom left` |  |
| `texture botton` |  |
| `texture bottom right` |  |

### `MsoBaselineAlignment`

| Value | Description |
|-------|-------------|
| `text baseline alignment unset` |  |
| `text baseline align baseline` |  |
| `text baseline align top` |  |
| `text baseline align center` |  |
| `text baseline align east asian50` |  |
| `text baseline align automatic` |  |

### `MsoClipboardFormat`

| Value | Description |
|-------|-------------|
| `clipboard format unset` |  |
| `native clipboard format` |  |
| `HTMl clipboard format` |  |
| `RTF clipboard format` |  |
| `plain text clipboard format` |  |

### `MsoTextRangeInsertPosition`

| Value | Description |
|-------|-------------|
| `insert before` |  |
| `insert after` |  |

### `MsoPictureType`

| Value | Description |
|-------|-------------|
| `save as default` |  |
| `save as PNG file` |  |
| `save as BMP file` |  |
| `save as GIF file` |  |
| `save as JPG file` |  |
| `save as PDF file` |  |

### `MsoPictureEffectType`

| Value | Description |
|-------|-------------|
| `no effect` |  |
| `effect background removal` |  |
| `effect blur` |  |
| `effect brightness contrast` |  |
| `effect cement` |  |
| `effect crisscross etching` |  |
| `effect chalk sketch` |  |
| `effect color temperature` |  |
| `effect cutout` |  |
| `effect film grain` |  |
| `effect glass` |  |
| `effect glow diffused` |  |
| `effect glow edges` |  |
| `effect light screen` |  |
| `effect line drawing` |  |
| `effect marker` |  |
| `effect mosiaic bubbles` |  |
| `effect paint brush` |  |
| `effect paint strokes` |  |
| `effect pastels smooth` |  |
| `effect pencil grayscale` |  |
| `effect pencil sketch` |  |
| `effect photocopy` |  |
| `effect plastic wrap` |  |
| `effect saturation` |  |
| `effect sharpen soften` |  |
| `effect texturizer` |  |
| `effect watercolor sponge` |  |

### `MsoSegmentType`

| Value | Description |
|-------|-------------|
| `line` |  |
| `curve` |  |

### `MsoEditingType`

| Value | Description |
|-------|-------------|
| `auto` |  |
| `corner` |  |
| `smooth` |  |
| `symmetric` |  |

### `MsoSmartArtNodePosition`

| Value | Description |
|-------|-------------|
| `default node position` |  |
| `after node` |  |
| `before node` |  |
| `above node` |  |
| `below node` |  |

### `MsoSmartArtNodeType`

| Value | Description |
|-------|-------------|
| `default node` |  |
| `assistant node` |  |

### `MsoOrgChartLayoutType`

| Value | Description |
|-------|-------------|
| `org chart layout unset` |  |
| `org chart layout standard` |  |
| `org chart layout both hanging` |  |
| `org chart layout left hanging` |  |
| `org chart layout right hanging` |  |
| `org chart layout default` |  |

### `MsoAlignCmd`

| Value | Description |
|-------|-------------|
| `align lefts` |  |
| `align centers` |  |
| `align rights` |  |
| `align tops` |  |
| `align middles` |  |
| `align bottoms` |  |

### `MsoDistributeCmd`

| Value | Description |
|-------|-------------|
| `distribute horizontally` |  |
| `distribute vertically` |  |

### `MsoOrientation`

| Value | Description |
|-------|-------------|
| `orientation unset` |  |
| `horizontal orientation` |  |
| `vertical orientation` |  |

### `MsoZOrderCmd`

| Value | Description |
|-------|-------------|
| `bring shape to front` |  |
| `send shape to back` |  |
| `bring shape forward` |  |
| `send shape backward` |  |
| `bring shape in front of text` |  |
| `send shape behind text` |  |

### `MsoScriptLanguage`

| Value | Description |
|-------|-------------|
| `bring shape to front` |  |
| `send shape to back` |  |
| `bring shape forward` |  |
| `send shape backward` |  |

### `MsoFlipCmd`

| Value | Description |
|-------|-------------|
| `flip horizontal` |  |
| `flip vertical` |  |

### `MsoTriState`

| Value | Description |
|-------|-------------|
| `true` |  |
| `false` |  |
| `C true` |  |
| `toggle` |  |
| `tri state unset` |  |

### `MsoBlackWhiteMode`

| Value | Description |
|-------|-------------|
| `black and white unset` |  |
| `black and white mode automatic` |  |
| `black and white mode gray scale` |  |
| `black and white mode light gray scale` |  |
| `black and white mode inverse gray scale` |  |
| `black and white mode gray outline` |  |
| `black and white mode black text and line` |  |
| `black and white mode high contrast` |  |
| `black and white mode black` |  |
| `black and white mode white` |  |
| `black and white mode dont show` |  |

### `MsoBarPosition`

| Value | Description |
|-------|-------------|
| `bar left` |  |
| `bar top` |  |
| `bar right` |  |
| `bar bottom` |  |
| `bar floating` |  |
| `bar pop up` |  |
| `bar menu` |  |

### `MsoBarProtection`

| Value | Description |
|-------|-------------|
| `no protection` |  |
| `no customize` |  |
| `no resize` |  |
| `no move` |  |
| `no change visible` |  |
| `no change dock` |  |
| `no vertical dock` |  |
| `no horizontal dock` |  |

### `MsoBarType`

| Value | Description |
|-------|-------------|
| `normal command bar` |  |
| `menubar command bar` |  |
| `popup command bar` |  |

### `MsoControlType`

| Value | Description |
|-------|-------------|
| `control custom` |  |
| `control button` |  |
| `control edit` |  |
| `control drop down` |  |
| `control combobox` |  |
| `button drop down` |  |
| `split drop down` |  |
| `OCX drop down` |  |
| `generic drop down` |  |
| `graphic drop down` |  |
| `control popup` |  |
| `graphic Popup` |  |
| `button popup` |  |
| `split button popup` |  |
| `split button MRU popup` |  |
| `control label` |  |
| `expanding grid` |  |
| `split expanding grid` |  |
| `control grid` |  |
| `control gauge` |  |
| `graphic combobox` |  |
| `control pane` |  |
| `active X` |  |
| `control group` |  |
| `control tab` |  |
| `control spinner` |  |

### `MsoButtonState`

| Value | Description |
|-------|-------------|
| `button state up` |  |
| `button state down` |  |
| `button state unset` |  |

### `MsoControlOLEUsage`

| Value | Description |
|-------|-------------|
| `neither` |  |
| `server` |  |
| `client` |  |
| `both` |  |

### `MsoButtonStyle`

| Value | Description |
|-------|-------------|
| `button automatic` |  |
| `button icon` |  |
| `button caption` |  |
| `button icon and caption` |  |

### `MsoComboStyle`

| Value | Description |
|-------|-------------|
| `combobox style normal` |  |
| `combobox style label` |  |

### `MsoMenuAnimation`

| Value | Description |
|-------|-------------|
| `None` |  |
| `Random` |  |
| `Unfold` |  |
| `Slide` |  |

### `MsoHyperlinkType`

| Value | Description |
|-------|-------------|
| `hyperlink type text range` |  |
| `hyperlink type shape` |  |
| `hyperlink type inline shape` |  |

### `MsoExtraInfoMethod`

| Value | Description |
|-------|-------------|
| `append string` |  |
| `post string` |  |

### `MsoAnimationType`

| Value | Description |
|-------|-------------|
| `idle` |  |
| `greeting` |  |
| `goodbye` |  |
| `begin speaking` |  |
| `character success major` |  |
| `get attention major` |  |
| `get attention minor` |  |
| `searching` |  |
| `printing` |  |
| `gesture right` |  |
| `writing noting something` |  |
| `working at something` |  |
| `thinking` |  |
| `sending mail` |  |
| `listens to computer` |  |
| `disappear` |  |
| `appear` |  |
| `get artsy` |  |
| `get techy` |  |
| `get wizardy` |  |
| `checking something` |  |
| `look down` |  |
| `look down left` |  |
| `look down right` |  |
| `look left` |  |
| `look right` |  |
| `look up` |  |
| `look up left` |  |
| `look up right` |  |
| `saving` |  |
| `gesture down` |  |
| `gesture left` |  |
| `gesture up` |  |
| `empty trash` |  |

### `MsoButtonSetType`

| Value | Description |
|-------|-------------|
| `button none` |  |
| `button ok` |  |
| `button cancel` |  |
| `buttons ok cancel` |  |
| `buttons yes no` |  |
| `buttons yes no cancel` |  |
| `buttons back close` |  |
| `buttons next close` |  |
| `buttons back next close` |  |
| `buttons retry cancel` |  |
| `buttons abort retry ignore` |  |
| `buttons search close` |  |
| `buttons back next snooze` |  |
| `buttons tips options close` |  |
| `buttons yes all no cancel` |  |

### `MsoIconType`

| Value | Description |
|-------|-------------|
| `icon none` |  |
| `icon application` |  |
| `icon alert` |  |
| `icon tip` |  |
| `icon alert critical` |  |
| `icon alert warning` |  |
| `icon alert info` |  |

### `MsoWizardActType`

| Value | Description |
|-------|-------------|
| `Inactive` |  |
| `Active` |  |
| `Suspend` |  |
| `Resume` |  |

### `MsoDocProperties`

| Value | Description |
|-------|-------------|
| `property type number` |  |
| `property type boolean` |  |
| `property type date` |  |
| `property type string` |  |
| `property type float` |  |

### `MsoAutomationSecurity`

| Value | Description |
|-------|-------------|
| `msoAutomationSecurityLow` |  |
| `msoAutomationSecurityByUI` |  |
| `msoAutomationSecurityForceDisable` |  |

### `MsoScreenSize`

| Value | Description |
|-------|-------------|
| `resolution544x376` |  |
| `resolution640x480` |  |
| `resolution720x512` |  |
| `resolution800x600` |  |
| `resolution1024x768` |  |
| `resolution1152x882` |  |
| `resolution1152x900` |  |
| `resolution1280x1024` |  |
| `resolution1600x1200` |  |
| `resolution1800x1440` |  |
| `resolution1920x1200` |  |

### `MsoCharacterSet`

| Value | Description |
|-------|-------------|
| `Arabic character set` |  |
| `Cyrillic character set` |  |
| `English character set` |  |
| `Greek character set` |  |
| `Hebrew character set` |  |
| `Japanese character set` |  |
| `Korean character set` |  |
| `Multilingual Unicode character set` |  |
| `Simplified Chinese character set` |  |
| `Thai character set` |  |
| `Traditional Chinese character set` |  |
| `Vietnamese character set` |  |

### `MsoEncoding`

| Value | Description |
|-------|-------------|
| `encoding Thai` |  |
| `encoding Japanese ShiftJIS` |  |
| `encoding simplified Chinese` |  |
| `encoding Korean` |  |
| `encoding Big5 traditional Chinese` |  |
| `encoding little endian` |  |
| `encoding big endian` |  |
| `encoding central European` |  |
| `encoding Cyrillic` |  |
| `encoding Western` |  |
| `encoding Greek` |  |
| `encoding Turkish` |  |
| `encoding Hebrew` |  |
| `encoding Arabic` |  |
| `encoding Baltic` |  |
| `encoding Vietnamese` |  |
| `encoding ISO88591 Latin1` |  |
| `encoding ISO88592 central Europe` |  |
| `encoding ISO88593 Latin3` |  |
| `encoding ISO88594 Baltic` |  |
| `encoding ISO88595 Cyrillic` |  |
| `encoding ISO88596 Arabic` |  |
| `encoding ISO88597 Greek` |  |
| `encoding ISO88598 Hebrew` |  |
| `encoding ISO88599 Turkish` |  |
| `encoding ISO885915 Latin9` |  |
| `encoding ISO2022 Japanese no half width Katakana` |  |
| `encoding ISO2022 Japanese JISX02021984` |  |
| `encoding ISO2022 Japanese JISX02011989` |  |
| `encoding ISO2022KR` |  |
| `encoding ISO2022CN traditional Chinese` |  |
| `encoding ISO2022CN simplified Chinese` |  |
| `encoding Mac Roman` |  |
| `encoding Mac Japanese` |  |
| `encoding Mac traditional Chinese` |  |
| `encoding Mac Korean` |  |
| `encoding Mac Arabic` |  |
| `encoding Mac Hebrew` |  |
| `encoding Mac Greek1` |  |
| `encoding Mac Cyrillic` |  |
| `encoding Mac simplified Chinese GB2312` |  |
| `encoding Mac Romania` |  |
| `encoding Mac Ukraine` |  |
| `encoding Mac Latin2` |  |
| `encoding Mac Icelandic` |  |
| `encoding Mac Turkish` |  |
| `encoding Mac Croatia` |  |
| `encoding EBCDIC US Canada` |  |
| `encoding EBCDIC International` |  |
| `encoding EBCDIC multilingual ROECE Latin2` |  |
| `encoding EBCDIC Greek modern` |  |
| `encoding EBCDIC Turkish Latin5` |  |
| `encoding EBCDIC Germany` |  |
| `encoding EBCDIC Denmark Norway` |  |
| `encoding EBCDIC Finland Sweden` |  |
| `encoding EBCDIC Italy` |  |
| `encoding EBCDIC Latin America Spain` |  |
| `encoding EBCDIC United Kingdom` |  |
| `encoding EBCDIC Japanese Katakana extended` |  |
| `encoding EBCDIC France` |  |
| `encoding EBCDIC Arabic` |  |
| `encoding EBCDIC Greek` |  |
| `encoding EBCDIC Hebrew` |  |
| `encoding EBCDIC Korean extended` |  |
| `encoding EBCDIC Thai` |  |
| `encoding EBCDIC Icelandic` |  |
| `encoding EBCDIC Turkish` |  |
| `encoding EBCDIC Russian` |  |
| `encoding EBCDIC Serbian Bulgarian` |  |
| `encoding EBCDIC Japanese Katakana extended and Japanese` |  |
| `encoding EBCDIC US Canada and Japanese` |  |
| `encoding EBCDIC extended and Korean` |  |
| `encoding EBCDIC simplified Chinese extended and simplified Chinese` |  |
| `encoding EBCDIC US Canada and traditional Chinese` |  |
| `encoding EBCDIC Japanese Latin extended and Japanese` |  |
| `encoding OEM United States` |  |
| `encoding OEM Greek` |  |
| `encoding OEM Baltic` |  |
| `encoding OEM multilingual LatinI` |  |
| `encoding OEM multilingual LatinII` |  |
| `encoding OEM Cyrillic` |  |
| `encoding OEM Turkish` |  |
| `encoding OEM Portuguese` |  |
| `encoding OEM Icelandic` |  |
| `encoding OEM Hebrew` |  |
| `encoding OEM Canadian French` |  |
| `encoding OEM Arabic` |  |
| `encoding OEM Nordic` |  |
| `encoding OEM CyrillicII` |  |
| `encoding OEM modern Greek` |  |
| `encoding EUC Japanese` |  |
| `encoding EUC Chinese simplified Chinese` |  |
| `encoding EUC Korean` |  |
| `encoding EUC Taiwanese traditional Chinese` |  |
| `encoding Devanagari` |  |
| `encoding Bengali` |  |
| `encoding Tamil` |  |
| `encoding Telugu` |  |
| `encoding Assamese` |  |
| `encoding Oriya` |  |
| `encoding Kannada` |  |
| `encoding Malayalam` |  |
| `encoding Gujarati` |  |
| `encoding Punjabi` |  |
| `encoding Arabic ASMO` |  |
| `encoding Arabic transparent ASMO` |  |
| `encoding Korean Johab` |  |
| `encoding Taiwan CNS` |  |
| `encoding Taiwan TCA` |  |
| `encoding Taiwan Eten` |  |
| `encoding Taiwan IBM5550` |  |
| `encoding Taiwan teletext` |  |
| `encoding Taiwan Wang` |  |
| `encoding IA5IRV` |  |
| `encoding IA5 German` |  |
| `encoding IA5 Swedish` |  |
| `encoding IA5 Norwegian` |  |
| `encoding US ASCII` |  |
| `encoding T61` |  |
| `encoding ISO6937 nonspacing accent` |  |
| `encoding KOI8R` |  |
| `encoding Ext alpha lowercase` |  |
| `encoding KOI8U` |  |
| `encoding Europa3` |  |
| `encoding HZGB simplified Chinese` |  |
| `encoding UTF7` |  |
| `encoding UTF8` |  |

### `reset`

| Value | Description |
|-------|-------------|
| `command bar` |  |
| `command bar control` |  |

---

## Microsoft Word Suite

Word specific classes and commands

### Commands

### `Word help`

Displays on-line Help information.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `help type` | `WdHelpType` | no | The type of help to be shown. |

### `accept`

Accepts the specified tracked change. The revision marks are removed, and the change is incorporated into the document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4024` | no |  |

### `accept all revisions`

Accepts all tracked changes in the document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `accept all shown revisions`

Accepts all shown tracked changes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `activate object`

Activate this object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4004` | no |  |

### `add addin`

Returns an add in object that represents an add-in added to the list of available add-ins.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `file name` | `text` | no | The path to the add-in to add. |
| `install` | `boolean` | yes | If true, the add-in will be installed, else it will only be added to the list of available add-ins. |

**Returns:** `add in`

### `add building block to category`

Creates a new building block.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `building block category` | no |  |
| `name` | `text` | no | The name of the building block entry. |
| `from location` | `text range` | no | The value of the buildling block entry. |
| `description` | `text` | yes | The description of the buildling block entry. |
| `insert options` | `WdDocPartInsertOptions` | no | Specifies whether the building block entry is inserted as a page, a paragraph, or inline. |

**Returns:** `building block`

### `add building block to template`

Creates a new building block entry in a template and returns a building block object that represents the new building block entry.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `template` | no |  |
| `name` | `text` | no | The name of the building block entry. |
| `type` | `WdBuildingBlockTypes` | no | The type of building block to create. Must be building block equations. |
| `category` | `text` | no | The category of the new building block entry. |
| `from location` | `text range` | no | The value of the buildling block entry. |
| `description` | `text` | yes | The description of the buildling block entry. |
| `insert options` | `WdDocPartInsertOptions` | no | Specifies whether the building block entry is inserted as a page, a paragraph, or inline. |

**Returns:** `building block`

### `add coauth lock`

Add a co-authoring lock.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `to range` | `text range` | yes | Specifies a range that will be locked. |
| `to paragraph` | `paragraph` | yes | Specifies a paragraph that will be locked. |
| `to column` | `column` | yes | Specifies a column that will be locked. |
| `to cell` | `cell` | yes | Specifies a cell that will be locked. |
| `to row` | `row` | yes | Specifies a row that will be locked. |
| `to table` | `table` | yes | Specifies a table that will be locked. |
| `to selection` | `selection object` | yes | Specifies a selection that will be locked. |
| `lock type` | `WdLockType` | yes | Specifies a lock type. |
| `in range` | `selection object` | yes | The text range into which to add the new lock. |
| `in coauthoring` | `coauthoring` | yes | The coauthoring into which to add the new lock. |
| `in coauthor` | `coauthor` | yes | The coauthor into which to add the new lock. |

**Returns:** `coauth lock`

### `add math ac entry`

Creates an equation auto correct entry.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math autocorrect` | no |  |
| `name` | `text` | no | The name of the autocorrect entry. |
| `value` | `text` | no | The value of the autocorrect entry. |

### `add math argument`

Inserts an argument into an equation with variable number of arguments, i.e. math delimiter and math equation array objects, and returns a math object object. You must specify only one function, row, or argument.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `before arg` | `integer` | yes | The ordinal position of the existing argument before which to add the new argument. |
| `to function` | `math function` | yes | The function to which to add the new argument. |
| `to matrix row` | `math matrix row` | yes | The matrix row to which to add the new argument. |
| `to matrix column` | `math matrix column` | yes | The matrix column to which to add the new argument. |

**Returns:** `math object`

### `add math function`

Inserts a new structure, such as a fraction, into an equation at the specified position.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math object` | no |  |
| `in location` | `text range` | no | The place at which to insert an equation. |
| `math function type` | `WdOMathFunctionType` | no | The type of equation to insert. |
| `number of arguments` | `integer` | yes | The number of arguments in the equation. |
| `number of columns` | `integer` | yes | The number of columns in the equation. |

### `add math recognized function`

Creates a new recognized function.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math autocorrect` | no |  |
| `name` | `text` | no | The name of the recognized function. |

### `add matrix column`

Creates a matrix column and adds it to a matrix and returns a math matrix column object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math matrix` | no |  |
| `before column` | `math matrix column` | yes | An existing column in the matrix before which to place the new column. |

**Returns:** `math matrix column`

### `add matrix row`

Creates a matrix row and adds it to a matrix and returns a math matrix row object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math matrix` | no |  |
| `before row` | `math matrix row` | yes | An existing row in the matrix before which to place the new row. |

**Returns:** `math matrix row`

### `append to spike`

Deletes the specified range and adds the contents of the range to the Spike which is a built-in autotext entry.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `template` | no |  |
| `range` | `text range` | no | The range that's deleted and appended to the Spike. |

**Returns:** `auto text entry` — The Spike autotext entry object.

### `apply bullet default`

Adds bullets and formatting to the paragraphs in the range for the list format object. If the paragraphs are already formatted with bullets, this method removes the bullets and formatting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list format` | no |  |
| `default list behavior` | `WdDefaultListBehavior` | yes | Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. |

### `apply document theme`

Applies an Office theme to the document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `file name` | `text` | no | The file path of the theme file to apply to this document. |

### `apply list format template`

Applies a set of list-formatting characteristics to the list format object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list format` | no |  |
| `list template` | `list template` | no | The list template to be applied. |
| `continue previous list` | `boolean` | yes | True to continue the numbering from the previous list. False to start a new list. |
| `apply to` | `WdListApplyTo` | yes | The portion of the list that the list template is to be applied to. |
| `default list behavior` | `WdDefaultListBehavior` | yes | Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. |

### `apply list template`

Applies a set of list-formatting characteristics to the specified Word list object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `Word list` | no |  |
| `list template` | `list template` | no | The list template to be applied. |
| `continue previous list` | `boolean` | yes | Set to true to continue the numbering from the previous list. Set to false to start a new list. |
| `default list behavior` | `WdDefaultListBehavior` | yes | Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. |

### `apply number default`

Adds the default numbering scheme to the paragraphs in the range for the list format object. If the paragraphs are already formatted as a numbered list, this method removes the numbers and formatting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list format` | no |  |
| `default list behavior` | `WdDefaultListBehavior` | yes | Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. |

### `apply outline number default`

Adds the default outline-numbering scheme to the paragraphs in the range for the list format object. If the paragraphs are already formatted as an outline-numbered list, this method removes the numbers and formatting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list format` | no |  |
| `default list behavior` | `WdDefaultListBehavior` | yes | Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. |

### `apply page borders to all sections`

Applies the specified page-border formatting to all sections in a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `border options` | no |  |

### `apply picture bullet`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list level` | no |  |
| `path` | `text` | no |  |

**Returns:** `inline shape`

### `arrange windows`

Arrange windows on the screen.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `arrange style` | `WdArrangeStyle` | yes | How the windows show be arranged |

### `auto format`

Automatically formats a document

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `autosave`

Causes an autosave to occur for the given document or all documents if no document is specified.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `document` | `document` | yes | The document to autosave. |

### `break link`

Breaks the link between the source file and the specified picture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `link format` | no |  |

### `build key code`

Returns a unique number for the specified key combination.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `key1` | `WdKey` | no | A key you specify by using one of the constants |
| `key2` | `WdKey` | yes | A key you specify by using one of the constants |
| `key3` | `WdKey` | yes | A key you specify by using one of the constants |
| `key4` | `WdKey` | yes | A key you specify by using one of the constants |

**Returns:** `integer` — The unique key code

### `build up`

Converts an equation to professional/built up format.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math object` | no |  |

### `calculate selection`

Calculates a mathematical expression within a selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

**Returns:** `real`

### `can check in`

Returns True if Word can check in a specified document to a server.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

**Returns:** `boolean`

### `can check out`

Returns True if Word can check out a specified document from a server.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `file name` | `text` | no | The server path and name of the document. |

**Returns:** `boolean`

### `can continue previous list`

Returns whether the formatting from the previous list can be continued.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4023` | no |  |
| `list template` | `list template` | no | A list template that's been applied to previous paragraphs in the document. |

**Returns:** `WdContinue`

### `centimeters to points`

Converts a measurement from centimeters to points.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `centimeters` | `real` | no | The centimeter value to be converted to points. |

**Returns:** `real`

### `change default table style`

Sets the default table style.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `style` | `WdBuiltinStyle` | no |  |
| `change in template` | `boolean` | no |  |

### `change file open directory`

Sets the folder in which Microsoft Word searches for documents.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `path` | `text` | no | The path to the folder in which Microsoft Word searches for documents. |

### `check`

Simulates the data merge operation, pausing to report each error as it occurs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |

### `check consistency`

Searches all text in a Japanese language document and displays instances where character usage is inconsistent for the same words.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `check grammar`

Checks a string for grammatical errors. Returns a boolean to indicate whether the string contains grammatical errors. True if the string contains no errors.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4002` | no |  |
| `text to check` | `text` | yes | The string you want to check for grammatical errors. |

**Returns:** `boolean`

### `check in`

Returns a document from a local computer to a server, and sets the local document to read-only so that it cannot be edited locally.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `save changes` | `boolean` | no | True saves the document to the server location. The default value is True. |
| `comments` | `text` | no | Comments for the revision of the document being checked in. Only applies if SaveChanges equals True. |
| `make public` | `boolean` | no | True allows the user to perform a publish on the document after being checked in. |

### `check in with version`

Returns a document from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `save changes` | `boolean` | no | True saves the document to the server location. The default value is True. |
| `comments` | `text` | no | Comments for the revision of the document being checked in. Only applies if SaveChanges equals True. |
| `make public` | `boolean` | no | True allows the user to perform a publish on the document after being checked in. |
| `version type` | `WdCheckInVersionType` | no | Version number of the document. |

### `check out`

Copies a specified document from a server to a local computer for editing. Returns a String that represents the local path and file name of the document checked out.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `file name` | `text` | no | The server path and name of the document. |

### `check spelling`

Checks a string for spelling errors. Returns true if the string has no spelling errors.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4003` | no |  |
| `text to check` | `text` | no | The string you want to check for spelling errors. |
| `custom dictionary` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `ignore uppercase` | `boolean` | yes | Set to true to ignore capitalization. |
| `main dictionary` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary2` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary3` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary4` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary5` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary6` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary7` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary8` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary9` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary10` | `dictionary` | yes | A dictionary object that should be used to check the string. |

**Returns:** `boolean`

### `clean string`

Removes nonprinting characters character codes 1 Ôøê 29 and special Microsoft Word characters from the specified string or changes them to spaces character code 32. Returns the result as a string.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `item to check` | `text` | no | The string you want to clean. |

**Returns:** `text` — The cleaned string.

### `clear`

Removes the object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4012` | no |  |

### `clear all fuzzy options`

Clears all nonspecific search options associated with Japanese text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `find` | no |  |

### `clear formatting`

Removes text and paragraph formatting from a selection or from the formatting specified in a find or replace operation.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4011` | no |  |

### `click object`

Clicks the specified field. If the field is a GOTOBUTTON field, this method moves the insertion point to the specified location or selects the specified bookmark. If the field is a HYPERLINK field, this method jumps to the target location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `field` | no |  |

### `close print preview`

Switches the specified document from print preview to the previous view. If the specified document isn't in print preview, an error occurs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `collapse outline`

Collapses the text under the selection or the specified range by one heading level.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `view` | no |  |
| `text range` | `text range` | yes | The range of paragraphs to be collapsed. If this argument is omitted, the entire selection is collapsed. |

### `compare`

Compares this document with another.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `path` | `text` | no | The name and path of the document with which the specified document is to be compared. |
| `author name` | `text` | yes | The name of the author to use for changes. |
| `target` | `WdCompareTarget` | yes | A value indicating where to place the comparison results. If this argument is omitted, the result will appear in a new document. |
| `detect format changes` | `boolean` | yes | A value indicating where to place the comparison results. If this argument is omitted, the result will appear in a new document. |
| `ignore all comparison warnings` | `boolean` | yes | A value indicating where to place the comparison results. If this argument is omitted, the result will appear in a new document. |
| `add to recent files` | `boolean` | yes | A value indicating where to place the comparison results. If this argument is omitted, the result will appear in a new document. |

### `compute statistics`

Computes a set of readability statistics for this document.  This must be called before accessing the readability statistics for this document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `statistic` | `WdStatistic` | no |  |
| `include footnotes and endnotes` | `boolean` | yes |  |

**Returns:** `integer`

### `convert`

Converts a multiple-level list to a single-level list, or vice versa.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list template` | no |  |
| `level` | `integer` | yes | The level to use for formatting the new list. When converting a multiple-level list to a single-level list, this argument can be a number from 1 through 9. When converting a single-level list to a multiple-level list, 1 is the only valid value. |

**Returns:** `list template` — The newly converted list template object

### `convert left scripts to sub and super scripts`

Converts an equation with a superscript or subscript to the left of the base of the equation to an equation with a base of a superscript or subscript.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math left scripts` | no |  |

**Returns:** `math function`

### `convert numbers to text`

Changes the list numbers and listnum fields in the document object to text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4009` | no |  |
| `number type` | `WdNumberType` | yes | The type of number to be converted. |

### `convert sub and super scripts to left scripts`

Converts an equation with a base superscript or subscript to an equation with a superscript or subscript to the left of the base.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math sub and super script` | no |  |

**Returns:** `math function`

### `convert to literal text`

Converts an equation to literal text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math object` | no |  |

### `convert to math text`

Converts an equation to math text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math object` | no |  |

### `convert to normal text`

Converts an equation to normal text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math object` | no |  |

### `copy bookmark`

Sets the bookmark specified by the name argument to the location marked by another bookmark, and returns a bookmark object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `bookmark` | no |  |
| `name` | `text` | no | The name of the new bookmark. |

**Returns:** `bookmark`

### `copy format`

Copies the character formatting of the first character in the selected text. If a paragraph mark is selected, Word copies paragraph formatting in addition to character formatting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `copy object`

Copies the content of the object to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4020` | no |  |

### `copy styles from template`

Copies styles from the specified template to a document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `template` | `text` | no | The template file name. |

### `count numbered items`

Returns the number of bulleted or numbered items and listnum fields in the document object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4010` | no |  |
| `number type` | `WdNumberType` | yes | The type of number to be counted. |
| `level` | `integer` | yes | A number that corresponds to the numbering level you want to count. If this argument is omitted, all levels are counted. |

**Returns:** `integer`

### `create data source`

Creates a Word document that uses a table to store data for a data merge. The new data source is attached to the specified document, which becomes a main document if it's not one already.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `name` | `text` | yes | The path and file name for the new data source. |
| `password document` | `text` | yes | The password required to open the new data source. |
| `write password` | `text` | yes | The password required to save changes to the data source. |
| `header record` | `text` | yes | Field names for the header record. |
| `MS Query` | `boolean` | yes | Set to true to launch Microsoft Query, if it's installed. |
| `SQL statement` | `text` | yes | Defines query options for retrieving data. |
| `SQL statement1` | `text` | yes | If the query string is longer than 255 characters, SQL statement specifies the first portion of the string, and SQL statement1 specifies the second portion. |
| `connection` | `text` | yes | When retrieving data through ODBC, you specify a connection string. |
| `link to source` | `boolean` | yes | Set to true to perform the query specified by connection and SQL statement each time the main document is opened. |

### `create header source`

Creates a Word document that stores a header record that's used in place of the data source header record in a mail merge. This method attaches the new header source to the specified document, which becomes a main document if it's not one already.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `name` | `text` | no | The path and file name for the new header source. |
| `password document` | `text` | yes | The password required to open the new header source. |
| `write password` | `text` | yes | The password required to save changes to the new header source. |
| `header record` | `text` | yes | A string that specifies the field names for the header record. |

### `create letter content`

Create a new letter content object for use with the letter wizard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `date format` | `text` | no | The date format to be used for a letter created by the letter wizard. |
| `include header footer` | `boolean` | no | Set to true if the header and footer from the page design template are to be included in a letter created by the letter wizard. |
| `page design` | `text` | no | The name and file path of the template attached to a document created by the letter wizard. |
| `letter style` | `WdLetterStyle` | no | Sets the layout of a letter created by the letter wizard. |
| `letterhead` | `boolean` | no | Set to true if space is to be reserved for a preprinted letterhead in a letter created by the letter wizard. |
| `letterhead location` | `WdLetterheadLocation` | no | Sets the location of a preprinted letterhead in a letter created by the letter wizard. |
| `letterhead size` | `real` | no | Sets the amount of space in points to be reserved for a preprinted letterhead in a letter created by the letter wizard. |
| `recipient name` | `text` | no | Sets the name of the person who'll be receiving the letter created by the letter wizard. |
| `recipient address` | `text` | no | Sets the mailing address of the person who'll be receiving the letter created by the letter wizard. |
| `salutation` | `text` | no | Sets the salutation text for the letter created by the letter wizard. |
| `salutation type` | `WdSalutationType` | no | Sets the salutation type for a letter created by the letter wizard. |
| `recipient reference` | `text` | no | Sets the reference line for example, In reply to: for a letter created by the letter wizard. |
| `mailing instructions` | `text` | no | Sets the mailing instruction text for a letter created by the letter wizard for example, Certified Mail. |
| `attention line` | `text` | no | Sets the attention line text for a letter created by the letter wizard. |
| `subject` | `text` | no | Sets the subject text of a letter created by the letter wizard. |
| `cc list` | `text` | no | Sets the carbon copy recipients for a letter created by the letter wizard. |
| `return address` | `text` | no | Sets the return address for a letter created with the letter wizard. |
| `sender name` | `text` | no | Sets the name of the person creating a letter with the letter wizard. |
| `closing` | `text` | no | Sets the closing text for a letter created by the letter wizard for example, Sincerely yours. |
| `sender company` | `text` | no | Sets the company name of the person creating a letter with the letter wizard. |
| `sender job title` | `text` | no | Sets the job title of the person creating a letter with the letter wizard. |
| `sender initials` | `text` | no | Sets the initials of the person creating a letter with the letter wizard. |
| `enclosure count` | `integer` | no | Sets the number of enclosures for a letter created by the letter wizard. |

**Returns:** `letter content` — The newly created letter content object.

### `create new document`

Create a new document

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `attached template` | `template` | yes | Reference or path to template to base new document on. |
| `new template` | `boolean` | yes | Set to true if the document to be created is to become a template for other new documents. |
| `new document type` | `WdNewDocumentType` | yes | Sets the type of document to be created. |

**Returns:** `document` — A reference to the newly created document.

### `create new document for hyperlink`

Creates a new document linked to the specified hyperlink.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `hyperlink object` | no |  |
| `file name` | `text` | no | The file name of the specified document. |
| `edit now` | `boolean` | no | Set to true to have the specified document open immediately in its associated editing environment. The default value is true. |
| `overwrite` | `boolean` | no | Set to true to overwrite any existing file of the same name in the same folder. Set to false if any existing file of the same name is preserved and the file name argument specifies a new file name. |

### `create new equation`

Creates an equation, from the text equation contained within the specified range, and returns a text range object that contains the new equation.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `from range` | `text range` | no | Specifies a range that contains a text equation. |
| `in document` | `document` | yes | The document into which to add the new equation. |
| `in range` | `selection object` | yes | The text range into which to add the new equation. |
| `in selection` | `selection object` | yes | The selection into which to add the new equation. |

### `create new field`

Create a new field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `text range` | `text range` | no | The location for the new field. |
| `field type` | `WdFieldType` | yes | The field type for the new field. |
| `field text` | `text` | yes | The field code for the new field. |
| `preserve formatting` | `boolean` | yes | Set to true to preserve formatting when field is updated |

### `create new mailing label document`

Creates a new label document using either the default label options or ones that you specify. Returns a document object that represents the new document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `mailing label` | no |  |
| `name` | `text` | yes | The mailing label name |
| `address` | `text` | yes | The text for the mailing label. |
| `auto text` | `text` | yes | The name of the autotext entry that includes the mailing label text. |
| `extract address` | `boolean` | yes | Set to true to use the text marked by the envelope address bookmark, a user-defined bookmark, as the recipient's address. |
| `laser tray` | `WdPaperTray` | yes | The laser printer tray. |
| `single label` | `boolean` | yes | Set to true if the text is placed within a single label on a sheet that contains multiple labels. This argument is used in conjunction with row and column. The default value is false. |
| `row` | `integer` | yes | Specifies the row in which to place the text when single label is set to true. |
| `column` | `integer` | yes | Specifies the column in which to place the text when single label is set to true. |

**Returns:** `document` — A reference to the newly created document.

### `create range`

Returns a text range object by using the specified starting and ending character positions.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `start` | `integer` | yes | The starting character position |
| `end` | `integer` | yes | The ending character position |

**Returns:** `text range`

### `create textbox`

Adds a default-size text box around the selection. If the selection is an insertion point, this method changes the pointer to a cross-hair pointer so that the user can draw a text box.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `cut object`

Removes the object from the document and places it on the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4021` | no |  |

### `data form`

Displays the data form dialog box, in which you can modify data records. You can use this method with a mail merge main document, a mail merge data source, or any document that contains data delimited by table cells or separator characters

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `delete all comments`

Deletes all comments.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `delete all shown comments`

Deletes all shown comments.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `disable`

Removes the specified key combination if it's currently assigned to a command. After you use this method, the key combination has no effect.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `key binding` | no |  |

### `display Word dialog`

Displays the specified built-in Word dialog box until either the user closes it or the specified amount of time has passed. Returns a Long that indicates which button was clicked to close the dialog box.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `dialog` | no |  |
| `time out` | `integer` | yes | The amount of time that Word will wait before closing the dialog box automatically. One unit is approximately 0.001 second. If this argument is omitted, the dialog box is closed when the user dismisses it. |

**Returns:** `integer` — -2 means the close button was pressed.  -1 means the ok button was pressed.  0 means the cancel button was pressed. If the number is greater than 0 then a command button was pressed.

### `do Word repeat`

Repeats the most recent editing action one or more times. Returns true if the commands were repeated successfully.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `times` | `integer` | yes | The number of times you want to repeat the last command. |

**Returns:** `boolean`

### `duplicate page`

Duplicates the page of the current selection and moves the selection to the new page. This command only works when in Publishing Layout View.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `edit data source`

Opens or switches to the mail merge data source.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |

### `edit header source`

Opens the header source attached to a mail merge main document, or activates the header source if it's already open.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |

### `edit main document`

Activates the mail merge main document associated with the specified header source or data source document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |

### `edit type`

Sets options for the specified text form field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text input` | no |  |
| `form field type` | `WdTextFormFieldType` | no | The text box type. |
| `default type` | `text` | yes | The default text that appears in the text box. |
| `type format` | `text` | yes | The formatting string used to format the text, number, or date for example, 0.00, Title Case. |
| `enabled` | `boolean` | yes | True to enable the form field for text entry |

### `enable`

Formats the first character in the specified paragraph as a dropped capital letter.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `drop cap` | no |  |

### `end key`

Moves or extends the selection to the end of the specified unit. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `move` | `WdUnits` | yes | The unit by which the selection is to be moved or extended |
| `extend` | `WdMovementType` | yes | Specifies the way the selection is moved. |

**Returns:** `text range` — The number of characters the selection was actually moved, or it returns zero if the move was unsuccessful.

### `endnote convert`

Converts endnotes to footnotes, or vice versa.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `endnote options` | no |  |

### `escape key`

Cancels a mode such as extend or column select.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `execute data merge`

Performs the specified data merge operation. Performs the specified data merge operation.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `pause` | `boolean` | yes | Set to true for Microsoft Word pause and display a troubleshooting dialog box if a mail merge error is found. False to report errors in a new document. |

### `execute dialog`

Applies the current settings of a Microsoft Word dialog box.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `dialog` | no |  |

### `execute find`

Runs the specified find operation. Returns true if the find operation is successful.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `find` | no |  |
| `find text` | `text` | yes | The text to be searched for. Use an empty string to search for formatting only. |
| `match case` | `boolean` | yes | Set to true to specify that the find text be case sensitive. |
| `match whole word` | `boolean` | yes | Set to true to have the find operation locate only entire words, not text that's part of a larger word. |
| `match wildcards` | `boolean` | yes | Set to true if the text to find contains wildcards. |
| `match sounds like` | `boolean` | yes | Set to true to have the find operation locate words that sound similar to the find text. |
| `match all word forms` | `boolean` | yes | Set to true to have the find operation locate all forms of the find text for example, sit locates sitting and sat. |
| `match forward` | `boolean` | yes | Set to true to search forward toward the end of the document. |
| `wrap find` | `WdFindWrap` | yes | Controls what happens if the search begins at a point other than the beginning of the document and the end of the document is reached or vice versa if forward is set to false. |
| `find format` | `boolean` | yes | Set to true to have the find operation locate formatting in addition to or instead of the find text. |
| `replace with` | `text` | yes | The replacement text. To delete the text specified by the Find argument, use an empty string. You specify special characters and advanced search criteria just as you do for the find text argument. |
| `replace` | `WdReplace` | yes | Specifies how many replacements are to be made: one, all, or none. |

**Returns:** `WdInsertionPoint`

### `execute key binding`

Runs the command associated with the specified key combination.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `key binding` | no |  |

### `expand`

Expands the selection. Returns the number of characters added to the range or selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `by` | `WdUnits` | yes | The unit by which to expand the range. |

**Returns:** `any` — The number of characters that the selection was expanded by.

### `expand outline`

Expands the text under the selection or the specified range by one heading level.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `view` | no |  |
| `text range` | `text range` | yes | The range of paragraphs to be expanded. If this argument is omitted, the entire selection is expanded. |

### `extend`

Extends the selection to the next larger unit of text. The progression of selected units of text is as follows: word, sentence, paragraph, section, entire document. The selection is extended by moving the active end of the selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `character` | `text` | yes | The character through which the selection is extended. This argument is case sensitive and must evaluate to a string or an error occurs. Also, if the value of this argument is longer than a single character, Word ignores the command entirely. |

### `find key`

Returns a key binding object that represents the specified key combination

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `key code` | `integer` | no | A key code returned from the build key code method. |
| `key_code_2` | `WdKey` | yes | A secondary key code returned from the build key code method. |

**Returns:** `key binding` — A key binding object.

### `find record`

Searches the contents of the specified data merge data source for text in a particular field. Returns true if the search text is found.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge data source` | no |  |
| `find text` | `text` | no | The text to be looked for. |
| `field name` | `text` | no | The name of the field to be searched. |

**Returns:** `boolean`

### `fit to pages`

Decreases the font size of text just enough so that the document will fit on one fewer pages. An error occurs if Word is unable to reduce the page count by one.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `follow`

Displays a cached document associated with the specified hyperlink object, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `hyperlink object` | no |  |
| `new window` | `boolean` | yes | Set to true to display the target document in a new window. The default value is false. |
| `extra info` | `text` | yes | A string or byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use extra info to specify the coordinates of an image map. |

### `follow hyperlink`

This method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application. If the hyperlink uses the file protocol, this method opens the document instead of downloading it.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `address` | `text` | yes | The address of the target document |
| `sub address` | `text` | yes | The location within the target document. The default value is an empty string. |
| `new window` | `boolean` | yes | Set to true to display the target location in a new window. The default value is false. |
| `add history` | `boolean` | yes | Set to true to add the link to the current day's history folder. |
| `extra info` | `text` | yes | A string or a byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use extra info to specify the coordinates of an image map, the contents of a form, or a file name. |

### `footnote convert`

Converts endnotes to footnotes, or vice versa.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `footnote options` | no |  |

### `get active writing style`

Returns the writing style for a specified language.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `language ID` | `WdLanguageID` | no | The language to get the writing style for. |

**Returns:** `text` — The writing style of the specified language

### `get border`

Returns the specified border object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4019` | no |  |
| `which border` | `WdBorderType` | no |  |

**Returns:** `border`

### `get cross reference items`

Returns an list of items that can be cross-referenced based on the specified cross-reference type.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `reference type` | `WdReferenceType` | no | The type of item you want to insert a cross-reference to. |

**Returns:** `text`

### `get default file path`

Returns the default folders for items such as documents, templates, and graphics.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `file path type` | `WdDefaultFilePath` | no | Which path should be returned. |

**Returns:** `text` — The specified default path.

### `get dialog`

Returns a dialog object for the specified dialog.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `WdWordDialog` | no |  |

**Returns:** `dialog`

### `get document compatibility`

Returns the current state of the specified compatibility item for this document. Compatibility options affect how a document is displayed in Microsoft Word.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `compatibility item` | `WdCompatibility` | no | The compatibility item to be checked. |

**Returns:** `boolean`

### `get international information`

Get the specified international information

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `WdInternationalIndex` | no |  |

**Returns:** `text`

### `get keys bound to`

Returns a list key binding objects that represents all the key combinations assigned to the specified item.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `key category` | `WdKeyCategory` | no | The category of the key combination. |
| `command` | `text` | no | The name of the command. |

**Returns:** `specifier` — A list of key binding objects.

### `get list gallery`

Returns the specified list gallery object object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `WdListGalleryType` | no |  |

**Returns:** `list gallery`

### `get next field`

Returns the next field object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

**Returns:** `field`

### `get previous field`

Returns the previous field object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

**Returns:** `field`

### `get private profile string`

Returns a string in a settings file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `system object` | no |  |
| `file name` | `text` | no | The file name for the settings file. If there's no path specified,  the Microsoft folder in the users preference folder is used. |
| `section` | `text` | no | The name of the section in the settings file that contains key. |
| `key` | `text` | no | The key whose setting you want to retrieve. |

**Returns:** `text`

### `get profile string`

Returns a string from the application's settings file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `system object` | no |  |
| `section` | `text` | no | The name of the section in the settings file that contains key. |
| `key` | `text` | no | The key whose setting you want to retrieve. |

**Returns:** `text`

### `get selection information`

Returns information about the selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `information type` | `WdInformation` | no | The information to be returned. |

**Returns:** `text` — The requested information.

### `get spelling suggestions`

Returns a record that specifies the error kind and a list of words.  The AEKeyword for the error kind is type and AEKeyword for the list is list.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `item to check` | `text` | no | The string you want to check for spelling errors. |
| `custom dictionary` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `ignore uppercase` | `boolean` | yes | Set to true to ignore capitalization. |
| `main dictionary` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `suggestion mode` | `WdSpellingWordType` | yes | Specifies the way Microsoft Word makes spelling suggestions.  The default is spelling word type spell word. |
| `custom dictionary2` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary3` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary4` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary5` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary6` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary7` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary8` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary9` | `dictionary` | yes | A dictionary object that should be used to check the string. |
| `custom dictionary10` | `dictionary` | yes | A dictionary object that should be used to check the string. |

**Returns:** `record`

### `get story range`

Returns a range object that represents the story specified by the story type argument.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `story type` | `WdStoryType` | no | The story type to be retrieved. |

**Returns:** `story range`

### `get synonym info object`

Returns a synonym info object that contains the information from the thesaurus on the synonyms, antonyms, or related words and expressions for the specified word or phrase.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `item to check` | `text` | no | The specified word or phrase. |
| `language ID` | `WdLanguageID` | yes | The language used for the theasurus. |

**Returns:** `synonym info` — The synonym info object.

### `get webpage font`

Return the specified web page font object for a given character set.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `MsoCharacterSet` | no |  |

**Returns:** `web page font` — The requested web page font object.

### `get zoom`

Returns a zoom object of the specified type for this pane.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pane` | no |  |
| `zoom type` | `WdViewType` | no | The type of zoom object that should be returned. |

**Returns:** `zoom` — The requested zoom object.

### `grow font`

Increases the font size to the next available size. If the selection or range contains more than one font size, each size is increased to the next available setting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `font` | no |  |

### `home key`

Moves or extends the selection to the beginning of the specified unit. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `move` | `WdUnits` | yes | The unit by which the selection is to be moved or extended |
| `extend` | `WdMovementType` | yes | Specifies the way the selection is moved. |

**Returns:** `integer` — The number of characters the selection was actually moved, or it returns zero if the move was unsuccessful.

### `inches to points`

Converts a measurement from inches to points.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `inches` | `real` | no | The inch value to be converted to points. |

**Returns:** `real`

### `insert`

Insert the string at the specified location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `text` | `text` | no | The string to be inserted |
| `at` | `location specifier` | no | The place the new text should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |

### `insert auto text`

Attempts to match the text in the specified range or the text surrounding the range with an existing auto text entry name.  If any such match is found, the auto text entry is inserted.  If no match an error occurs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The range to check for the auto text entry name. |

### `insert auto text entry`

Inserts the auto text entry in place of the specified range. Returns a text range object that represents the auto text entry.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `auto text entry` | no |  |
| `where` | `text range` | no | The location for the auto text entry. |
| `rich text` | `boolean` | yes | Set to true to insert the AutoText entry with its original formatting. |

**Returns:** `text range` — A text range object that represents the inserted auto text entry.

### `insert break`

Inserts a break in the specified place of the specified kind.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The place the new break should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |
| `break type` | `WdBreakType` | yes | The type of break to be inserted. |

### `insert building block`

Inserts the value of a building block into a document and returns a text range object that represents the contents of the building block within the document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `building block` | no |  |
| `in location` | `text range` | no | The location of where to place the contents of the building block. |
| `as rich text` | `boolean` | yes | True inserts the building block as rich, formatted text. False inserts the building block as plain text. |

**Returns:** `text range`

### `insert caption`

Inserts a caption immediately preceding or following the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The place the new caption should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |
| `caption label` | `WdCaptionLabelID` | no | The caption label to be inserted. |
| `title` | `text` | yes | The string to be inserted immediately following the label in the caption. |
| `caption position` | `WdCaptionPosition` | yes | Specifies whether the caption will be inserted above or below the specified range. |

### `insert cells`

Adds cells to an existing table. The number of cells inserted is equal to the number of cells in the selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `shift cells` | `WdInsertCells` | yes | Specifies how the new cells will be inserted. |

### `insert columns`

Inserts columns to the left of the column that contains the selection. If the selection isn't in a table, an error occurs. The number of columns inserted is equal to the number of columns selected.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `position` | `WdInsertRightLeft` | yes | Specifies the position the new columns will be inserted. |

### `insert cross reference`

Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The place the new cross reference should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |
| `reference type` | `WdReferenceType` | no | The type of item for which a cross-reference is to be inserted. |
| `reference kind` | `WdReferenceKind` | no | The information to be included in the cross-reference. |
| `reference item` | `text` | no | If the reference kind is reference type bookmark, this argument specifies a bookmark name.  For all other reference types values, this argument specifies the item number or name in the Reference type box in the Cross-reference dialog box. |
| `insert as hyperlink` | `boolean` | yes | Set this to true to insert the cross-reference as a hyperlink to the reference item. |
| `include position` | `boolean` | yes | Set this to true to insert above or below, depending on the location of the reference item in relation to the cross-reference. |

### `insert database`

Retrieves data from a data source and inserts the data as a table in place of the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The place where the new data table should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |
| `format` | `WdTableFormat` | yes | Specify the format of the new table. |
| `style` | `integer` | yes | The attributes of the auto-format specified by the format argument. Use the sum of the following items None = 0, Borders = 1, Shading = 2, Font = 4, Color = 8, Auto Fit = 16, Heading Rows = 32, Last Row = 64, First Column = 128, Last Column = 256. |
| `link to source` | `boolean` | yes | Set to true to establish a link between the new table and the data source. |
| `connection` | `text` | yes | When retrieving data through Open Database Connectivity, you specify a connection string. |
| `SQL statement` | `text` | yes | An optional query string that retrieves a subset of the data in a primary data source to be inserted into the document. |
| `SQL statement1` | `text` | yes | If the query string is longer than 255 characters, SQL statement denotes the first portion of the string and SQL statement1 denotes the second portion. |
| `password document` | `text` | yes | The password if any required to open the data source. |
| `password template` | `text` | yes | If the data source is a word document, this argument is the password if any required to open the attached template. |
| `write password` | `text` | yes | The password required to save changed to the document. |
| `write password template` | `text` | yes | The password required to save changed to the template. |
| `data source` | `text` | yes | The path and file name of the data source. |
| `from` | `integer` | yes | The number of the first data record in the range of records to be inserted. |
| `to` | `integer` | yes | The number of the last data record in the range of records to be inserted. |
| `include fields` | `boolean` | yes | Set this to true to include field names from the data source in the first row of the new table. |

### `insert date time`

Insert the correct date or time, or both, either as text or as a time field at the specified location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The place where the date-time data should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |
| `date time format` | `text` | yes | The format to be used for displaying the date or time, or both.  The format is like MMMM dd, yyyy |
| `insert as field` | `boolean` | yes | Set to true to insert the specified information as a time field.  The default is true. |

### `insert envelope data`

Inserts an envelope as a separate section at the beginning of the specified document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `envelope` | no |  |
| `extract address` | `boolean` | yes | Set to true to use the text marked by the envelope address bookmark, a user-defined bookmark, as the recipient's address. |
| `address` | `text` | yes | A string that specifies the recipient's address. This argument is ignored if extract address is true. |
| `auto text` | `text` | yes | A string that specifies an autotext entry to use for the address. If specified, address is ignored. |
| `omit return address` | `boolean` | yes | Set to true to not insert a return address. |
| `return address` | `text` | yes | A string that specifies the return address. |
| `return autotext` | `text` | yes | A string that specifies an autotext entry to use for the return address. If specified, return address is ignored. |
| `print bar code` | `boolean` | yes | Set to true to add a POSTNET bar code. For U.S. mail only. |
| `print FIMA` | `boolean` | yes | Set to true to add a Facing Identification Mark FIMA for use in presorting courtesy reply mail. For U.S. mail only. |
| `envelope size` | `text` | yes | A string that specifies the envelope size. The string must match one of the sizes listed in the envelope size box in the envelope options dialog box, for example, Size 10 or C4. |
| `envelope height` | `integer` | yes | The height of the envelope, measured in points, when the size argument is set to Custom size. |
| `envlope width` | `integer` | yes | The width of the envelope, measured in points, when the size argument is set to Custom size. |
| `feed source` | `boolean` | yes | Set to true to use the feed source property of the envelope object to specify which paper tray to use when printing the envelope. |
| `address from left` | `integer` | yes | The distance, measured in points, between the left edge of the envelope and the recipient's address. |
| `address from top` | `integer` | yes | The distance, measured in points, between the top edge of the envelope and the recipient's address. |
| `return address from left` | `integer` | yes | The distance, measured in points, between the left edge of the envelope and the return address. |
| `return address from top` | `integer` | yes | The distance, measured in points, between the top edge of the envelope and the return address. |
| `default face up` | `boolean` | yes | Set to true to print the envelope face up, false to print it face down. |
| `default orientation` | `WdEnvelopeOrientation` | yes | The orientation for the envelope. |
| `size from page setup` | `boolean` | yes | Set to true if the envelope's address areas are sized according to settings in the envelopes dialog box in the page setup dialog box. False if they are sized according to custom settings. The default value is true. |
| `show page setup dialog` | `boolean` | yes | Set to true if the page setup dialog box is displayed to allow adjustment of settings. The size from pagesetup must be set to true for the box to be displayed. The default value is false. |
| `create new document` | `boolean` | yes | Set to true if the envelope is not inserted into the active document but created separately. The default value is true. |

### `insert file`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The place where the file data should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |
| `file name` | `text` | no | The path and file name of the file to be inserted. |
| `file range` | `text` | yes | If the specified file is a Microsoft Word document, this parameter refers to a bookmark.  If the file is an Excel worksheet, this parameter refers to a named range or a cell range i.e. R1C1:R3C4. |
| `confirm conversions` | `boolean` | yes | Set to true to have Microsoft Word prompt you to confirm conversion when inserting files in formats other than Microsoft Word document format. |
| `link` | `boolean` | yes | Set to true to insert the file using an include text field. |

### `insert formula`

Inserts an = Formula field that contains a formula at the selection. The formula replaces the selection, if the selection isn't collapsed.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `formula` | `text` | yes | The mathematical formula you want the = Formula field to evaluate. Spreadsheet-type references to table cells are valid. |
| `number format` | `text` | yes | A format for the result of the = Formula field. |

### `insert math break`

Inserts a break into an equation and returns a math break object that represents the break.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math object` | no |  |
| `at range` | `text range` | no | The position at which to insert the break in the equation. |

**Returns:** `math break`

### `insert page`

Insert a page following the page of the current selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `insert paragraph`

Replaces the specified range with a new paragraph.  If you do not want to replace the range, use the collapse range method before using this method.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The place where the new paragraph should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |

### `insert rows`

Inserts the specified number of new rows above the row that contains the selection. If the selection isn't in a table, an error occurs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `position` | `WdInsertAboveBelow` | yes | Specifies the position the new rows will be inserted. |
| `number of rows` | `integer` | yes | The number of rows to insert.  The default is one. |

### `insert symbol`

Inserts a symbol in place of the specified range.  If you do not want to replace the range, use the collapse range method before using this method.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `text range` | no | The place where the symbol should be inserted.  It must specify a range within a document i.e. the fifth word of paragraph 1 of document 1. |
| `character number` | `integer` | no | The character number for the specified symbol. This value will always be the sum of 31 and the number that corresponds to the position of the symbol in the table of symbols. |
| `font` | `text` | yes | The name of the font that contains the symbol. |
| `unicode` | `boolean` | yes | Set to true to insert the unicode character specified by character number, false to insert the ANSI character specified by the character number.  The default is false. |
| `bias` | `WdFontBias` | yes | Sets the font bias for symbols. This argument is useful for setting the correct font bias for east asian characters. |

### `key string`

Returns the key combination string for the specified keys

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `key code` | `integer` | no | A key you specify by using one of the constants |
| `key_code_2` | `WdKey` | yes | A key you specify by using one of the constants |

**Returns:** `text` — String that corresponds to the specified key code.

### `large scroll`

Scrolls a window by the specified number of screens. This method is equivalent to clicking just before or just after the scroll boxes on the horizontal and vertical scroll bars.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4005` | no |  |
| `down` | `integer` | yes | The number of screens to scroll the window down. |
| `up` | `integer` | yes | The number of screens to scroll the window up. |
| `to right` | `integer` | yes | The number of screens to scroll the window to the right. |
| `to left` | `integer` | yes | The number of screens to scroll the window to the left. |

### `linearize`

Converts an equation to linear/built down format.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math object` | no |  |

### `lines to points`

Converts a measurement from lines to points. 1 line = 12 points.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `lines` | `real` | no | The line value to be converted to points. |

**Returns:** `real`

### `list commands`

Creates a new document and then inserts a table of Microsoft Word commands along with their associated shortcut keys and menu assignments.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `list all commands` | `boolean` | no | True to include all Microsoft Word commands and their assignments. False to include only commands with customized assignments. |

### `list indent`

Increases the list level of the paragraphs in the range for the list format object, in increments of one level.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list format` | no |  |

### `list outdent`

Decreases the list level of the paragraphs in the range for the list format object, in increments of one level.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list format` | no |  |

### `lower limit to upper limit`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math lower limit` | no |  |

**Returns:** `math function`

### `make compatibility default`

Uses the correct settings of the document compatibility options set by the set compatibility options method as the default for new documents.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `make new data merge ask field`

Create a new data merge ask field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |
| `name` | `text` | no | The bookmark name that the response or default text is assigned to. Use a REF field with the bookmark name to display the result in a document. |
| `prompt` | `text` | yes | The text that's displayed in the dialog box. |
| `default ask text` | `text` | yes | The default response, which appears in the text box when the dialog box is displayed. |
| `ask once` | `boolean` | yes | Set to true to display the dialog box only once instead of each time a new data record is merged. |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `make new data merge fill in field`

Create a new data merge fill in field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |
| `prompt` | `text` | yes | The text that's displayed in the dialog box. |
| `default fill in text` | `text` | yes | The default response, which appears in the text box when the dialog box is displayed. |
| `ask once` | `boolean` | yes | Set to true to display the dialog box only once instead of each time a new data record is merged. |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `make new data merge if field`

Create a new data merge if field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |
| `merge field` | `text` | no | The merge field name. |
| `comparison` | `WdMailMergeComparison` | no | The operator used in the comparison. |
| `compare to` | `text` | yes | The text to compare with the contents of data merge field |
| `true text` | `text` | yes | The text that's inserted if the comparison is true. |
| `false text` | `text` | yes | The text that's inserted if the comparison is false. |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `make new data merge next field`

Create a new data merge next field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `make new data merge next if field`

Create a new data merge next if field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |
| `merge field` | `text` | no | The merge field name. |
| `comparison` | `WdMailMergeComparison` | no | The operator used in the comparison. |
| `compare to` | `text` | yes | The text to compare with the contents of data merge field |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `make new data merge rec field`

Create a new data merge rec field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `make new data merge sequence field`

Create a new data merge sequence field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `make new data merge set field`

Create a new data merge set field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |
| `name` | `text` | no | The bookmark name that value text is assigned to. |
| `value text` | `text` | yes | The text associated with the bookmark specified by the main argument. |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `make new data merge skip if field`

Create a new data merge skip if field

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `text range` | `text range` | no | The location for the new field. |
| `merge field` | `text` | no | The merge field name. |
| `comparison` | `WdMailMergeComparison` | no | The operator used in the comparison. |
| `compare to` | `text` | yes | The text to compare with the contents of data merge field |

**Returns:** `data merge field` — A reference to the newly created data merge field.

### `manual hyphenation`

Initiates manual hyphenation of a document, one line at a time. The user is prompted to accept or decline suggested hyphenations.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `mark entry for table of contents`

Inserts a table of contents entry field after the specified range. The method returns a field object representing the new field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `range` | `text range` | no | The location of the entry. The new field is inserted after range. |
| `entry` | `text` | no | The text that appears in the table of contents. To indicate a subentry, include the main entry text and the subentry text, separated by a colon for example, Introduction:The Product. |
| `table ID` | `text` | yes | A one-letter identifier for the table of figures item for example, i for an illustration. |
| `level` | `integer` | yes | A level for the entry in the table of figures. |

**Returns:** `field` — The newly created field

### `mark entry for table of figures`

Inserts a table of figures entry field after the specified range. The method returns a field object representing the new field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `range` | `text range` | no | The location of the entry. The new field is inserted after range. |
| `entry` | `text` | no | The text that appears in the table of figures. To indicate a subentry, include the main entry text and the subentry text, separated by a colon for example, Introduction:The Product. |
| `table ID` | `text` | yes | A one-letter identifier for the table of figures item for example, i for an illustration. |
| `level` | `integer` | yes | A level for the entry in the table of figures. |

**Returns:** `field` — The newly created field

### `mark for index`

Inserts an index entry field after the specified range. The method returns a field object representing the new field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `range` | `text range` | no | The location of the entry. The new field is inserted after range. |
| `entry` | `text` | no | The text that appears in the index. To indicate a subentry, include the main entry text and the subentry text, separated by a colon for example, Introduction:The Product. |
| `cross reference` | `text` | yes | A cross-reference that will appear in the index . |
| `bookmark name` | `text` | yes | The name of the bookmark that marks the range of pages you want to appear in the index. If this argument is omitted, the number of the page containing the new field appears in the index. |

**Returns:** `field` — The newly created field

### `merge`

Merges the changes marked with revision marks from one document to another.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `file name` | `text` | no | The file path and name of the document to merge with this document. |

### `merge subdocuments`

Merges the specified subdocuments of a master document into a single subdocument.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `first subdocument` | `subdocument` | yes | The first subdocument in a range of subdocuments to be merged. |
| `last subdocument` | `subdocument` | yes | The last subdocument in a range of subdocuments to be merged. |

### `millimeters to points`

Converts a measurement from millimeters to points.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `millimeters` | `real` | no | The millimeter value to be converted to points. |

**Returns:** `real`

### `modified`

if the specified list template is not the built-in list template for that position in the list gallery. Index goes from 1 to 7

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list gallery` | no |  |
| `index` | `integer` | no | A number from 1 to 7 that corresponds to the position of the template in the bullets and numbering dialog box. |

**Returns:** `boolean`

### `next for browser`

Moves the selection to the next item indicated by the browser target. Use the browser target property to change the browser target.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `browser` | no |  |

### `next header footer`

If the selection is in a header, this method moves to the next header within the current section or to the first header in the following section. If the selection is in a footer, this method moves to the next footer.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `view` | no |  |

### `next revision`

Locates and returns the next tracked change as a revision object. The changed text becomes the current selection. Use the properties of the resulting revision object to see what type of change it is, who made it, and so forth.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `wrap` | `boolean` | yes | Set to true to continue searching for a revision at the beginning of the document when the end of the document is reached. The default value is false. |

**Returns:** `revision` — The newly created revision object.

### `open as document`

Opens the specified template as a document and returns a document object. Opening a template as a document allows the user to edit the contents of the template.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `template` | no |  |

**Returns:** `document` — The document object that corresponds to the template that was opened as a document.

### `open data source`

Attaches a data source to the specified document, which becomes a main document if it's not one already.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `name` | `text` | no | The data source file name. You can specify a Microsoft Query file instead of specifying a data source, a connection string, and a query string. |
| `format` | `WdOpenFormat` | yes | The file converter used to open the document. |
| `confirm conversions` | `boolean` | yes | Set to true to display the convert file dialog box if the file isn't in Microsoft Word format. |
| `read only` | `boolean` | yes | Set to true to open the data source on a read-only basis. |
| `link to source` | `boolean` | yes | Set to true to perform the query specified by connection and SQL statement each time the main document is opened. |
| `add to recent files` | `boolean` | yes | Set to true to add the file name to the list of recently used files. |
| `password document` | `text` | yes | The password used to open the data source. |
| `password template` | `text` | yes | The password used to open the template. |
| `revert` | `boolean` | yes | Controls what happens if name is the file name of an open document. True to discard any unsaved changes to the open document and reopen the file. False to activate the open document. |
| `write password` | `text` | yes | The password used to save changes to the document. |
| `write password template` | `text` | yes | The password used to save changes to the template. |
| `connection` | `text` | yes | When retrieving data through ODBC, you specify a connection string. |
| `SQL statement` | `text` | yes | Defines query options for retrieving data. |
| `SQL statement1` | `text` | yes | If the query string is longer than 255 characters, SQL statement specifies the first portion of the string, and SQL statement1 specifies the second portion. |

### `open header source`

Attaches a mail merge header source to the specified document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `name` | `text` | no | The file name of the header source. |
| `format` | `WdOpenFormat` | yes | The file converter used to open the document. |
| `confirm conversions` | `boolean` | yes | Set to true to display the convert file dialog box if the file isn't in Microsoft Word format. |
| `read only` | `boolean` | yes | Set to true to open the header source on a read-only basis. |
| `add to recent files` | `boolean` | yes | Set to true to add the file name to the list of recently used files. |
| `password document` | `text` | yes | The password required to open the header source document. |
| `password template` | `text` | yes | The password required to open the header source template. |
| `revert` | `boolean` | yes | Controls what happens if name is the file name of an open document. True to discard any unsaved changes to the open document and reopen the file. False to activate the open document. |
| `write password` | `text` | yes | The password required to save changes to the document data source. |
| `write password template` | `text` | yes | The password required to save changes to the template data source. |

### `open recent file`

Opens the recent file and returns a document object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `recent file` | no |  |

**Returns:** `document`

### `open subdocument`

Opens the specified subdocument. Returns a document object representing the opened object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `subdocument` | no |  |

**Returns:** `document` — The document object that contains the newly opened subdocument.

### `open version`

Opens the specified version of the document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document version` | no |  |

### `organizer copy`

Copies the specified autotext entry, toolbar, style, or macro project item from the source document or template to the destination document or template.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `source` | `text` | no | The document or template file name that contains the item you want to copy. |
| `destination` | `text` | no | The document or template file name to which you want to copy an item. |
| `name` | `text` | no | The name of the autotext entry, toolbar, style, or macro you want to copy. |
| `organizer object type` | `WdOrganizerObject` | no | The kind of item you want to copy. |

### `organizer delete`

Deletes the specified style, autotext entry, toolbar, or macro project item from a document or template.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `source` | `text` | no | The document or template file name that contains the item you want to copy. |
| `name` | `text` | no | The name of the autotext entry, toolbar, style, or macro you want to delete. |
| `organizer object type` | `WdOrganizerObject` | no | The kind of item you want to delete. |

### `organizer rename`

Renames the specified style, autotext entry, toolbar, or macro project item in a document or template.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `source` | `text` | no | The document or template file name that contains the item you want to rename. |
| `name` | `text` | no | The name of the autotext entry, toolbar, style, or macro you want to rename. |
| `new name` | `text` | no | The new name for the autotext entry, toolbar, style, or macro. |
| `organizer object type` | `WdOrganizerObject` | no | The kind of item you want to rename. |

### `page scroll`

Scrolls through the window page by page.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4007` | no |  |
| `down` | `integer` | yes | The number of pages to be scrolled down. If this argument is omitted, this value is assumed to be 1. |
| `up` | `integer` | yes | The number of pages to be scrolled up. |

### `paste format`

Applies formatting copied with the copy format method to the selection. If a paragraph mark was selected when the copy format method was used, Word applies paragraph formatting in addition to character formatting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `paste object`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `picas to points`

Converts a measurement from picas to points.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `picas` | `real` | no | The pica value to be converted to points. |

**Returns:** `real`

### `points to centimeters`

Converts a measurement from points to centimeters.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `points` | `real` | no | The point value to be converted to centimeters. |

**Returns:** `real`

### `points to inches`

Converts a measurement from points to inches.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `points` | `real` | no | The point value to be converted to inches. |

**Returns:** `real`

### `points to lines`

Converts a measurement from points to lines. 1 line = 12 points.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `points` | `real` | no | The point value to be converted to lines. |

**Returns:** `real`

### `points to millimeters`

Converts a measurement from points to millimeters.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `points` | `real` | no | The point value to be converted to millimeters. |

**Returns:** `real`

### `points to picas`

Converts a measurement from points to picas.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `points` | `real` | no | The point value to be converted to picas. |

**Returns:** `real`

### `previous for browser`

Moves the selection to the previous item indicated by the browser target. Use the browser target property to change the browser target.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `browser` | no |  |

### `previous header footer`

If the selection is in a header, this method moves to the previous header within the current section or to the last header in the previous section. If the selection is in a footer, this method moves to the previous footer.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `view` | no |  |

### `previous revision`

Locates and returns the previous tracked change as a revision object. The changed text becomes the current selection. Use the properties of the resulting revision object to see what type of change it is, who made it, and so forth.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `wrap` | `boolean` | yes | Set to true to continue searching for a revision at the beginning of the document when the end of the document is reached. The default value is false. |

**Returns:** `revision` — The newly created revision object.

### `print out`

Prints out all or part of the specified or active document. This command has been deprecated; use the Print command in the Standard Suite.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4001` | no |  |
| `append` | `boolean` | yes | Set to true to append the specified document to the file name specified by the output file name argument. Setting to false will overwrite the contents of the output file name argument. |
| `print out range` | `WdPrintOutRange` | yes | The page range. |
| `output file name` | `text` | yes | If print to file is true, this argument specifies the path and file name of the output file. |
| `page from` | `integer` | yes | The starting page number when print out range is set to print from to. |
| `page to` | `integer` | yes | The ending page number when print out range is set to print from to. |
| `print out item` | `WdPrintOutItem` | yes | The item to be printed. |
| `print copies` | `integer` | yes | The number of copies to be printed. |
| `print out page type` | `WdPrintOutPages` | yes | The type of pages to be printed. |
| `print to file` | `boolean` | yes | Set to true to send printer instructions to a file. Make sure to specify a file name with output file name argument. |
| `collate` | `boolean` | yes | When printing multiple copies of a document, true to print all pages of the document before printing the next copy. |
| `file name` | `text` | yes | The path and file name of the document to be printed. If this argument is omitted, Microsoft Word prints the active document. |
| `manual duplex print` | `boolean` | yes | True to print a two-sided document on a printer without a duplex printing kit. |

### `print out envelope`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `envelope` | no |  |
| `extract address` | `boolean` | yes | Set to true to use the text marked by the envelope address bookmark, a user-defined bookmark, as the recipient's address. |
| `address` | `text` | yes | A string that specifies the recipient's address. This argument is ignored if extract address is true. |
| `auto text` | `text` | yes | A string that specifies an autotext entry to use for the address. If specified, address is ignored. |
| `omit return address` | `boolean` | yes | Set to true to not insert a return address. |
| `return address` | `text` | yes | A string that specifies the return address. |
| `return autotext` | `text` | yes | A string that specifies an autotext entry to use for the return address. If specified, return address is ignored. |
| `print bar code` | `boolean` | yes | Set to true to add a POSTNET bar code. For U.S. mail only. |
| `print FIMA` | `boolean` | yes | Set to true to add a Facing Identification Mark FIMA for use in presorting courtesy reply mail. For U.S. mail only. |
| `envelope size` | `text` | yes | A string that specifies the envelope size. The string must match one of the sizes listed in the envelope size box in the envelope options dialog box, for example, Size 10 or C4. |
| `envelope height` | `integer` | yes | The height of the envelope, measured in points, when the size argument is set to Custom size. |
| `envlope width` | `integer` | yes | The width of the envelope, measured in points, when the size argument is set to Custom size. |
| `feed source` | `boolean` | yes | Set to true to use the feed source property of the envelope object to specify which paper tray to use when printing the envelope. |
| `address from left` | `integer` | yes | The distance, measured in points, between the left edge of the envelope and the recipient's address. |
| `address from top` | `integer` | yes | The distance, measured in points, between the top edge of the envelope and the recipient's address. |
| `return address from left` | `integer` | yes | The distance, measured in points, between the left edge of the envelope and the return address. |
| `return address from top` | `integer` | yes | The distance, measured in points, between the top edge of the envelope and the return address. |
| `default face up` | `boolean` | yes | Set to true to print the envelope face up, false to print it face down. |
| `default orientation` | `WdEnvelopeOrientation` | yes | The orientation for the envelope. |
| `size from page setup` | `boolean` | yes | Set to true if the envelope's address areas are sized according to settings in the envelopes dialog box in the page setup dialog box. False if they are sized according to custom settings. The default value is true. |
| `show page setup dialog` | `boolean` | yes | Set to true if the page setup dialog box is displayed to allow adjustment of settings. The size from pagesetup must be set to true for the box to be displayed. The default value is false. |

### `print out mailing label`

Prints a label or a page of labels with the same address.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `mailing label` | no |  |
| `name` | `text` | yes | The mailing label name. |
| `address` | `text` | yes | The text for the label address. |
| `extract address` | `boolean` | yes | Set to true to use the text marked by the envelope address bookmark, a user-defined bookmark, as the recipient's address. |
| `laser tray` | `WdPaperTray` | yes | The laser printer tray. |
| `single label` | `boolean` | yes | Set to true if the text is placed within a single label on a sheet that contains multiple labels. This argument is used in conjunction with row and column. The default value is false. |
| `row` | `integer` | yes | Specifies the row in which to place the text when single label is set to true. |
| `column` | `integer` | yes | Specifies the column in which to place the text when single label is set to true. |

### `print preview`

Switches the view to print preview.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `protect`

Helps to protect the specified document from changes. When a document is protected, users can make only limited changes, such as adding annotations, making revisions, or completing a form.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `protection type` | `WdProtectionType` | no | The protection type for the specified document. |
| `no reset` | `boolean` | yes | Set to false to reset form fields to their default values. Set to true to retain the current form field values if the specified document is protected. If protection type isn't allow only form fields, the no reset argument is ignored. |
| `password` | `text` | yes | The password required to remove protection from the specified document. |
| `use irm` | `boolean` | yes | Set to true if protection should use IRM. |
| `enforce style locks` | `boolean` | yes | Set to true if style locks should be enforced. |

### `rebind`

Changes the command assigned to the specified key binding.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `key binding` | no |  |
| `key category` | `WdKeyCategory` | no | The key category of the specified key binding. |
| `command` | `text` | no | The name of the specified command. |
| `command parameter` | `text` | yes | Additional text, if any, required for the command specified by command |

### `redo`

Redoes the last action that was undone. It reverses the undo method. Returns true if the actions were redone successfully.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `times` | `integer` | yes | The number of actions to be undone. |

**Returns:** `boolean`

### `reject`

Rejects the specified tracked change. The revision marks are removed, leaving the original text intact.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4025` | no |  |

### `reject all revisions`

Rejects all tracked changes in the document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `reject all shown revisions`

Rejects all shown tracked changes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `reload`

Reloads a cached document by resolving the hyperlink to the document and downloading it.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `remove numbers`

Removes numbers or bullets from the document

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4008` | no |  |
| `number type` | `WdNumberType` | yes | The type of number to be removed. |

### `remove page`

Removes the page of the selection and moves the selection to the following page. When removing the last page, the selection is moved to the previous page. This command only works when in Publishing Layout View.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `remove subscript`

Removes the subscript for an equation and returns a math function object that represents the updated equation without the subscript.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math sub and super script` | no |  |

**Returns:** `math function`

### `remove superscript`

Removes the superscript for an equation and returns a math function object that represents the updated equation without the superscript.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math sub and super script` | no |  |

**Returns:** `math function`

### `remve emphemeral locks`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `in range` | `selection object` | yes | The text range into which to add the new lock. |
| `in coauthoring` | `coauthoring` | yes | The coauthoring into which to add the new lock. |
| `in coauthor` | `coauthor` | yes | The coauthor into which to add the new lock. |

### `repaginate`

Repaginates the entire document

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `reset`

Removes changes that were made to a font.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `font` | no |  |

### `reset continuation notice`

Resets the footnote or endnote continuation notice to the default notice. The default notice is blank.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4015` | no |  |

### `reset continuation separator`

Resets the footnote or endnote continuation separator to the default separator. The default separator is a long horizontal line that separates document text from notes continued from the previous page.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4014` | no |  |

### `reset ignore all`

Clears the list of words that were previously ignored during a spelling check. After you run this method, previously ignored words are checked along with all the other words.

### `reset list gallery`

Resets the list template specified by index for the specified list gallery to the built-in list template format.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list gallery` | no |  |
| `index` | `integer` | no |  |

### `reset separator`

Resets the footnote or endnote separator to the default separator. The default separator is a short horizontal line that separates document text from notes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4013` | no |  |

### `retrieve language`

Returns the language object for the specified language.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `WdLanguageID` | no |  |

**Returns:** `language`

### `run VB macro`

Runs a Visual Basic macro.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `macro name` | `text` | no | The name of the VB macro to run. |

### `run letter wizard`

Runs the Letter Wizard on the document

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `letter content` | `letter content` | yes | A letter content object. Any filled properties in the letter content object show up as prefilled elements in the letter wizard dialog boxes. If this argument is omitted, the letter content property is automatically used to get an object from the document. |
| `wizard mode` | `boolean` | yes | Set to true to display the letter wizard dialog box as a series of steps with a Next, Back, and Finish button. Set to false to display the Letter Wizard dialog box as if it were opened from the menu. The default value is true. |

### `save as`

Saves the document with a new name or format.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `file name` | `text` | yes | The name for the document. The default is the current folder and file name. If a document with the specified file name already exists, the document is overwritten without the user being prompted first. |
| `file format` | `WdSaveFormat` | yes | The format in which the document is saved. |
| `lock comments` | `boolean` | yes | Set to true to lock the document for comments. The default is false. |
| `password` | `text` | yes | A password string for opening the document. |
| `add to recent files` | `boolean` | yes | True to add the document to the list of recently used files on the File menu. The default is true. |
| `write password` | `text` | yes | A password string for saving changes to the document. |
| `read only recommended` | `boolean` | yes | True to have Microsoft Word suggest read-only status whenever the document is opened. The default is false. |
| `embed truetype fonts` | `boolean` | yes | Set to true to save TrueType fonts with the document. |
| `save native picture format` | `boolean` | yes | If graphics were imported from another platform for example, Macintosh, true to save only the Windows version of the imported graphics. |
| `save forms data` | `boolean` | yes | True to save the data entered by a user in a form as a data record. |
| `text encoding` | `unsigned integer` | yes | Text encoding to use when saving out as text file |
| `insert line breaks` | `boolean` | yes |  |
| `allow substitutions` | `boolean` | yes |  |
| `line ending type` | `WdLineEndingType` | yes |  |
| `Add BiDi Marks` | `boolean` | yes | True to save with BiDi Marks. |

### `save version`

Saves a version of the document with a comment.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `comment` | `text` | yes | The comment that will be saved for this version. |

### `screen refresh`

Updates the display on the monitor with the current information in the video memory buffer. You can use this method after using the screen updating property to disable screen updates.

### `select cell`

Selects the entire cell containing the current selection. To use this method, the current selection must be contained within a single cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `select column`

Selects the column that contains the insertion point, or selects all columns that contain the selection. If the selection isn't in a table, an error occurs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `select current alignment`

Extends the selection forward until text with a different paragraph alignment is encountered.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `select current color`

Extends the selection forward until text with a different color is encountered.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `select current font`

Extends the selection forward until text in a different font or font size is encountered.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `select current indent`

Extends the selection forward until text with different left or right paragraph indents is encountered.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `select current spacing`

Extends the selection forward until a paragraph with different line spacing is encountered.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `select current tabs`

Extends the selection forward until a paragraph with different tab stops is encountered.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `select row`

Selects the row that contains the insertion point, or selects all rows that contain the selection. If the selection isn't in a table, an error occurs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `send html mail`

Opens a message window for sending the specified document, formatted as html, through Microsoft Outlook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `send mail`

Opens a message window for sending the specified document through your registered mail program.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `set active writing style`

Sets the writing style for the specified language.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `language ID` | `WdLanguageID` | no | The language to set the writing style for in the specified document. |
| `writing style` | `text` | no | The writing style i.e. Technical |

### `set all fuzzy options`

Activates all nonspecific search options associated with Japanese text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `find` | no |  |

### `set as font template default`

Sets the font formatting as the default for the active document and all new documents based on the active template. The default font formatting is stored in the Normal style.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `font` | no |  |

### `set as page setup template default`

Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `page setup` | no |  |

### `set default file path`

Sets the default folders for items such as documents, templates, and graphics.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `file path type` | `WdDefaultFilePath` | no | Which path should be set. |
| `path` | `text` | no | The new file path to be set |

### `set document compatibility`

Sets the current state of the specified compatibility item for this document. Compatibility options affect how a document is displayed in Microsoft Word.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `compatibility item` | `WdCompatibility` | no | The compatibility item to be set. |
| `is compatible` | `boolean` | no | The value to be set |

### `set number of text columns`

Arranges text into the specified number of text columns.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `page setup` | no |  |
| `number of columns` | `integer` | no | The number of columns the text is to be arranged into. |

### `set private profile string`

Sets a string in a settings file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `system object` | no |  |
| `file name` | `text` | no | The file name for the settings file. If there's no path specified,  the Microsoft folder in the users preference folder is used. |
| `section` | `text` | no | The name of the section in the settings file that contains key. |
| `key` | `text` | no | The key whose setting you want to set. |
| `private profile string` | `text` | no | The string to be set |

### `set profile string`

Sets a string from the application's settings file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `system object` | no |  |
| `section` | `text` | no | The name of the section in the settings file that contains key. |
| `key` | `text` | no | The key whose setting you want to set. |
| `profile string` | `text` | no | The string to be set |

### `show`

Displays and carries out actions initiated in the specified built-in Word dialog box. Returns a number which indicates the button used to dismiss the dialog box.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `dialog` | no |  |
| `time out` | `integer` | yes | The amount of time that Word will wait before closing the dialog box automatically. One unit is approximately 0.001 second. If this argument is omitted, the dialog box is closed when the user dismisses it. |

**Returns:** `integer` — -2 means the close button was pressed.  -1 means the ok button was pressed.  0 means the cancel button was pressed. If the number is greater than 0 then a command button was pressed.

### `show all headings`

Toggles between showing all text, headings and body text, and showing only headings.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `view` | no |  |

### `show heading`

Shows all headings up to the specified heading level and hides subordinate headings and body text. This method generates an error if the view isn't outline view or master document view.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `view` | no |  |
| `level` | `integer` | no | The outline heading level, a number from 1 to 9. |

### `shrink discontiguous selection`

Deselects all but the most recently selected text when a selection contains multiple, unconnected selections.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `shrink font`

Decreases the font size to the next available size. If the selection or range contains more than one font size, each size is decreased to the next available setting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `font` | no |  |

### `shrink selection`

Shrinks the selection to the next smaller unit of text. The progression is as follows: entire document, section, paragraph, sentence, word, insertion point.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `small scroll`

Scrolls a window by the specified number of lines. This method is equivalent to clicking the scroll arrows on the horizontal and vertical scroll bars.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4006` | no |  |
| `down` | `integer` | yes | The number of lines to scroll the window down. |
| `up` | `integer` | yes | The number of lines to scroll the window up. |
| `to right` | `integer` | yes | The number of lines to scroll the window to the right. A line corresponds to the distance scrolled by clicking the right scroll arrow on the horizontal scroll bar once. |
| `to left` | `integer` | yes | The number of lines to scroll the window to the left. A line corresponds to the distance scrolled by clicking the left scroll arrow on the horizontal scroll bar once. |

### `speak text`

Have the selected text be spoken.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `split subdocument`

Divides an existing subdocument into two subdocuments at the same level in master document view or outline view. The division is at the beginning of the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `subdocument` | no |  |
| `text range` | `text range` | no | The range that, when the subdocument is split, becomes a separate subdocument. |

### `split table in selection`

Inserts an empty paragraph above the first row in the selection. If the selection isn't in the first row of the table, the table is split into two tables.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `substitute font`

Sets font-mapping options, which are reflected in the font substitution dialog box

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `unavailable font` | `text` | no | The name of a font not available on your computer that you want to map to a different font for display and printing. |
| `substitute font` | `text` | no | The name of a font available on your computer that you want to substitute for the unavailable font. |

### `swap with endnotes`

Converts all footnotes in a document to endnotes and vice versa.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `footnote options` | no |  |

### `swap with footnotes`

Converts all footnotes in a document to endnotes and vice versa.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `endnote options` | no |  |

### `three way merge`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `local document` | `document` | no | The local working copy. |
| `server document` | `document` | no | The server copy downloaded. |
| `base document` | `document` | no | The base copy. |
| `favor source` | `boolean` | no |  |

### `toggle portrait`

Switches between portrait and landscape page orientations for a document or section.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `page setup` | no |  |

### `toggle ribbon`

Toggles whether or not the ribbon is shown in this window.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `toggle show all reviewers`

Toggles whether or not all reviewers are shown in this window.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `window` | no |  |

### `toggle show codes`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `type backspace`

Deletes the character preceding a collapsed selection, an insertion point. If the selection isn't collapsed to an insertion point, the selection is deleted.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `type paragraph`

Inserts a new, blank paragraph. If the selection isn't collapsed to an insertion point, it's replaced by the new paragraph.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |

### `type text`

Inserts the specified text. If the Word options property replace selection is true, the selection is replaced by the specified text. If replace selection is false, the specified text is inserted before the selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection object` | no |  |
| `text` | `text` | no |  |

### `undo`

Undoes the last action or a sequence of actions, which are displayed in the undo list. Returns true if the actions were successfully undone.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `times` | `integer` | yes | The number of actions to be undone. |

**Returns:** `boolean`

### `undo clear`

Clear the list of actions that can be undone.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `unlink`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4017` | no |  |

### `unload addins`

Unloads all loaded add-ins and, depending on the value of the remove from list argument, removes them from the list of add-ins.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `remove from list` | `boolean` | no |  |

### `unlock`

Remove the co-authoring lock.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `coauth lock` | no |  |

### `unprotect`

Removes protection from the specified document. If the document isn't protected, this method generates an error.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |
| `password` | `text` | yes | The password string used to protect the document. Passwords are case-sensitive. If the document is protected with a password and the correct password isn't supplied, a dialog box prompts the user for the password. |

### `update`

Updates the values shown in a built-in Microsoft Word dialog box, updates the specified link, or updates the entries shown in specified index, table of authorities, table of figures or table of contents.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4016` | no |  |

### `update document`

Updates the envelope in the document with the current envelope settings.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `envelope` | no |  |

### `update field`

Updates the result of the field object. When applied to a field object, returns true if the field is updated successfully.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `field` | no |  |

**Returns:** `boolean`

### `update page numbers`

Updates the page numbers for items in the specified table of contents or table of figures.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4022` | no |  |

### `update source`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4018` | no |  |

### `update styles`

Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `upgrade`

upgrade document

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `upper limit to lower limit`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `math upper limit` | no |  |

**Returns:** `math function`

### `use address book`

Selects the address book that's used as the data source for a mail merge operation.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `data merge` | no |  |
| `book type` | `text` | no |  |

### `use default folder suffix`

Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `web options` | no |  |

### `view property browser`

Displays the property window for the selected control in the specified document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### `web page preview`

Displays a preview of the document as it would look if saved as a Web page.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document` | no |  |

### Classes

### `Word comment`

Represents a single comment.

*Plural:* `Word comments`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `author` | `text` | r/w |  |
| `comment index` | `integer` | r/o |  |
| `comment text` | `text range` | r/o |  |
| `date value` | `date` | r/o |  |
| `initials` | `text` | r/w |  |
| `note reference` | `text range` | r/o |  |
| `scope` | `text range` | r/o |  |
| `show tip` | `boolean` | r/o |  |

### `Word list`

Represents a single list format that's been applied to specified paragraphs in a document.

*Plural:* `Word lists`

**Contains:** `paragraph`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `single list template` | `boolean` | r/o | Returns if the entire Word list object uses the same list template. |
| `style name` | `text` | r/o |  |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the Word list. |

### `Word options`

Represents application and document options in Word.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `EnableMisusedWordsDictionary` | `boolean` | r/w | Returns or sets if Microsoft Word checks for misused words when checking the spelling and grammar in a document. |
| `IME automatic control` | `boolean` | r/w | Returns or sets if Microsoft Word is set to automatically open and close the Japanese Input Method Editor. |
| `RTF in clipboard` | `boolean` | r/w | Returns or sets if all text copied from Word to the Clipboard retains its character and paragraph formatting. |
| `allow accented uppercase` | `boolean` | r/w | Returns or sets if accents are retained when a French language character is changed to uppercase. |
| `allow click and type mouse` | `boolean` | r/w | Returns or sets if click and type functionality is enabled. |
| `allow drag and drop` | `boolean` | r/w | Returns or sets if dragging and dropping can be used to move or copy a selection. |
| `allow fast save` | `boolean` | r/w | Returns or sets if Word saves only changes to a document. When reopening the document, Word uses the saved changes to reconstruct the document. |
| `animate screen movements` | `boolean` | r/w | Returns or sets if Word animates mouse movements, uses animated cursors, and animates actions such as background saving and find and replace operations. |
| `apply east asian fonts to ascii` | `boolean` | r/w | Returns or sets if Microsoft Word applies East Asian fonts to Latin text. |
| `auto format apply bulleted lists` | `boolean` | r/w | Returns or sets if characters, such as asterisks, hyphens, and greater-than signs, at the beginning of list paragraphs are replaced with bullets from the Bullets and Numbering dialog box when Word formats a document or range automatically. |
| `auto format apply headings` | `boolean` | r/w | Returns or sets if styles are automatically applied to headings when Word formats a document or range automatically. |
| `auto format apply lists` | `boolean` | r/w | Returns or sets if  if styles are automatically applied to lists when Word formats a document or range automatically. |
| `auto format apply other paragraphs` | `boolean` | r/w | Returns or sets if styles are automatically applied to paragraphs that aren't headings or list items when Word formats a document or range automatically. |
| `auto format as you type apply borders` | `boolean` | r/w | Returns or sets if a series of three or more hyphens -, equal signs =, or underscore characters _ are automatically replaced by a specific border line when the ENTER key is pressed. |
| `auto format as you type apply bulleted lists` | `boolean` | r/w | Returns or sets if bullet characters, such as asterisks, hyphens, and greater-than signs, are replaced with bullets from the bullets and numbering dialog box as you type. |
| `auto format as you type apply closings` | `boolean` | r/w | Returns or sets if Microsoft Word to automatically applies the closing style to letter closings as you type. |
| `auto format as you type apply dates` | `boolean` | r/w | Returns or sets if Microsoft Word  automatically applies the date style to dates as you type. |
| `auto format as you type apply first indents` | `boolean` | r/w | Returns or sets if Microsoft Word automatically replaces a space entered at the beginning of a paragraph with a first-line indent. |
| `auto format as you type apply headings` | `boolean` | r/w | Returns or sets if styles are automatically applied to headings as you type. |
| `auto format as you type apply numbered lists` | `boolean` | r/w | Returns or sets if paragraphs are automatically formatted as numbered lists with a numbering scheme from the Bullets and Numbering dialog box according to what's typed. |
| `auto format as you type apply tables` | `boolean` | r/w | Returns or set if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. The plus signs become the column borders, and the hyphens become the column widths. |
| `auto format as you type auto letter wizard` | `boolean` | r/w | Returns or sets if Microsoft Word to automatically starts the Letter Wizard when the user enters a letter salutation or closing. |
| `auto format as you type define styles` | `boolean` | r/w | Returns or sets if Word automatically creates new styles based on manual formatting. |
| `auto format as you type delete auto spaces` | `boolean` | r/w | Returns or sets if Microsoft Word to automatically deletes spaces inserted between Japanese and Latin text as you type. |
| `auto format as you type format list item beginning` | `boolean` | r/w | Returns or sets if Word repeats character formatting applied to the beginning of a list item to the next list item. |
| `auto format as you type insert closings` | `boolean` | r/w | Returns or sets if Microsoft Word to automatically inserts the corresponding memo closing when the user enters a memo heading. |
| `auto format as you type insert overs` | `boolean` | r/w | Returns or sets if Word will automatically inset certain Japanese characters for other Japanese characters. |
| `auto format as you type match parentheses` | `boolean` | r/w | Returns or sets if Microsoft Word to automatically corrects improperly paired parentheses. |
| `auto format as you type replace east asian dashes` | `boolean` | r/w | Returns or sets if Microsoft Word to automatically corrects long vowel sounds and dashes. |
| `auto format as you type replace fractions` | `boolean` | r/w | Returns or sets if typed fractions are replaced with fractions from the current character set as you type. |
| `auto format as you type replace hyperlinks` | `boolean` | r/w | Returns or sets if e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLs, are automatically changed to hyperlinks as you type. |
| `auto format as you type replace ordinals` | `boolean` | r/w | Returns or sets if the ordinal number suffixes st, nd, rd, and th are replaced with the same letters in superscript as you type. For example, 1st is replaced with 1 followed by st formatted as superscript. |
| `auto format as you type replace plain text emphasis` | `boolean` | r/w | Returns or sets if manual emphasis characters are automatically replaced with character formatting as you type. |
| `auto format as you type replace quotes` | `boolean` | r/w | Returns or sets if straight quotation marks are automatically changed to smart, curly, quotation marks as you type. |
| `auto format as you type replace symbols` | `boolean` | r/w | Returns or sets if two consecutive hyphens -- are replaced with an en dash Ôøê or an em dash Ôøë as you type. |
| `auto format delete auto spaces` | `boolean` | r/w | Returns or sets if Microsoft Word deletes spaces inserted between Japanese and Latin text, when Word formats a document or range automatically. |
| `auto format match parentheses` | `boolean` | r/w | Returns or sets if Microsoft Word  automatically corrects improperly paired parentheses. |
| `auto format preserve styles` | `boolean` | r/w | Returns or sets if previously applied styles are preserved when Word formats a document or range automatically. |
| `auto format replace east asian dashes` | `boolean` | r/w | Returns or sets if Microsoft Word automatically corrects long vowel sounds and dashes. |
| `auto format replace first indents` | `boolean` | r/w | Returns or sets for Microsoft Word to automatically replace a space entered at the beginning of a paragraph with a first-line indent. |
| `auto format replace fractions` | `boolean` | r/w | Returns or sets if typed fractions are replaced with fractions from the current character set when Word formats a document or range automatically. |
| `auto format replace hyperlinks` | `boolean` | r/w | Returns or sets if e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLS, are changed to hyperlinks, when Word formats a document or range automatically. |
| `auto format replace ordinals` | `boolean` | r/w | Returns or sets if the ordinal number suffixes st, nd, rd, and th are replaced with the same letters in superscript when Word formats a document or range automatically. For example, 1st is replaced with 1 followed by st formatted as superscript. |
| `auto format replace plain text emphasis` | `boolean` | r/w | Returns or sets if manual emphasis characters are replaced with character formatting when Word formats a document or range automatically. |
| `auto format replace quotes` | `boolean` | r/w | Returns or sets if straight quotation marks are automatically changed to smart, curly, quotation marks when Word formats a document or range automatically. |
| `auto format replace symbols` | `boolean` | r/w | Returns or set if two consecutive hyphens -- are replaced by an en dash Ôøê or an em dash Ôøë when Word formats a document or range automatically. |
| `auto word selection` | `boolean` | r/w | Returns or sets if dragging selects one word at a time instead of one character at a time. |
| `automatically build up equations` | `boolean` | r/w | Returns or sets whether Microsoft Word automatically converts equations to professional format. |
| `ay match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between some Japanese characters. |
| `blue screen` | `boolean` | r/w | Returns or sets if Word displays text as white characters on a blue background. |
| `button field clicks` | `integer` | r/w | Returns or sets the number of clicks, either one or two, required to run a GOTOBUTTON or MACROBUTTON field. |
| `bv match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between some Japanese characters. |
| `byte match fuzzy` | `boolean` | r/w | Returns or sets Microsoft Word ignores the distinction between full-width and half-width characters, Latin or Japanese, during a search. |
| `case match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between uppercase and lowercase letters during a search. |
| `check grammar as you type` | `boolean` | r/w | Returns or sets if Word checks grammar and marks errors automatically as you type. |
| `check grammar with spelling` | `boolean` | r/w | Returns or sets if Word checks grammar while checking spelling. |
| `check spelling as you type` | `boolean` | r/w | Returns or sets if Word checks spelling and marks errors automatically as you type. |
| `comments color` | `WdColorIndex` | r/w | Returns or sets the color of comments. |
| `confirm conversions` | `boolean` | r/w | Returns or sets if Word displays the convert file dialog box before it opens or inserts a file that isn't a Word document or template. In the convert file dialog box, the user chooses the format to convert the file from. |
| `convert high ansi to east asian` | `boolean` | r/w | Returns or sets if Microsoft Word converts text that is associated with an East Asian font to the appropriate font when it opens a document. |
| `create backup` | `boolean` | r/w | Returns or sets if Word creates a backup copy each time a document is saved. |
| `dash match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between minus signs, long vowel sounds, and dashes during a search. |
| `default border color` | `integer` | r/w | Returns or sets the default RGB color to use for new border objects. |
| `default border color index` | `WdColorIndex` | r/w | Returns or sets the default line color index for borders. |
| `default border color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the default line color for borders. |
| `default border line style` | `WdLineStyle` | r/w | Returns or sets the default border line style |
| `default border line width` | `WdLineWidth` | r/w | Returns or sets the default line width of borders. |
| `default highlight color index` | `WdColorIndex` | r/w | Returns or sets the color index  used to highlight text formatted with the highlight button. |
| `default open format` | `WdOpenFormat` | r/w | Returns or sets the default file converter used to open documents. |
| `deleted cell color` | `WdCellColor` | r/w | Returns or sets the color of cells that are deleted while change tracking is enabled. |
| `deleted text color` | `WdColorIndex` | r/w | Returns or sets the color of text that is deleted while change tracking is enabled. |
| `deleted text mark` | `WdDeletedTextMark` | r/w | Returns or sets the format of text that is deleted while change tracking is enabled. |
| `display grid lines` | `boolean` | r/w | Returns or sets if Microsoft Word displays the document grid. |
| `display paste options` | `boolean` | r/w | Returns or sets if Microsoft Word  displays the Paste Options button, which displays directly under newly pasted text. |
| `dz match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between some Japanese characters. |
| `enable sound` | `boolean` | r/w | Returns or sets if Word makes the computer respond with a sound whenever an error occurs. |
| `envelope feeder installed` | `boolean` | r/o | Returns true if the current printer has a special feeder for envelopes. |
| `grid distance horizontal` | `real` | r/w | Returns or sets the amount of horizontal space between the invisible gridlines that Word uses when you draw, move, and resize autoshapes or East Asian characters in new documents. |
| `grid distance vertical` | `real` | r/w | Returns or sets the amount of vertical space between the invisible gridlines that Word uses when you draw, move, and resize autoshapes or East Asian characters in new documents. |
| `grid origin horizontal` | `real` | r/w | Returns or sets the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing autoshapes or East Asian characters to begin in new documents. |
| `grid origin vertical` | `real` | r/w | Returns or sets the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing autoshapes or East Asian characters to begin in new documents. |
| `hf match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between some Japanese characters. |
| `hiragana match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between hiragana and katakana during a search. |
| `ignore internet and file addresses` | `boolean` | r/w | Returns or sets if file name extensions,  paths, e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLs, are ignored while checking spelling. |
| `ignore mixed digits` | `boolean` | r/w | Returns or sets if words that contain numbers are ignored while checking spelling. |
| `ignore uppercase` | `boolean` | r/w | Returns or sets if words in all uppercase letters are ignored while checking spelling. |
| `inline conversion` | `boolean` | r/w | Returns or sets if Microsoft Word displays an unconfirmed character string in the Japanese Input Method Editor as an insertion between existing character strings. |
| `insert key for paste` | `boolean` | r/w | Returns or sets if the insert key can be used for pasting the clipboard contents. |
| `inserted cell color` | `WdCellColor` | r/w | Returns or sets the color of cells that are inserted while change tracking is enabled. |
| `inserted text color` | `WdColorIndex` | r/w | Returns or sets the color of text that is inserted while change tracking is enabled. |
| `inserted text mark` | `WdInsertedTextMark` | r/w | Returns or sets how Microsoft Word formats inserted text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the inserted text color property to control the look of inserted text. |
| `iteration mark match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between types of repetition marks during a search. |
| `kanji match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between standard and nonstandard kanji ideography during a search. |
| `keep track of formatting` | `boolean` | r/w | Returns or sets if Microsoft Word keeps track of formatting. |
| `ki ku match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between some Japanese characters. |
| `live word count` | `boolean` | r/w | Returns or sets if the word count is displayed in the status bar. |
| `map paper size` | `boolean` | r/w | Returns or sets if documents formatted for another country's or region's standard paper size, for example, A4 are automatically adjusted so that they're printed correctly on your country's/region's standard paper size, for example, Letter. |
| `measurement unit` | `WdMeasurementUnits` | r/w | Returns or sets the standard measurement unit for Microsoft Word. |
| `merged cell color` | `WdCellColor` | r/w | Returns or sets the color of cells that are merged while change tracking is enabled. |
| `move from text color` | `WdColorIndex` | r/w | Returns or sets the color of text that is moved from while change tracking is enabled. |
| `move from text mark` | `WdMoveFromTextMark` | r/w | Returns or sets how Microsoft Word formats moved text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the moved text color property to control the look of moved text. |
| `move to text color` | `WdColorIndex` | r/w | Returns or sets the color of text that is moved to while change tracking is enabled. |
| `move to text mark` | `WdMoveToTextMark` | r/w | Returns or sets how Microsoft Word formats moved text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the moved text color property to control the look of moved text. |
| `old kana match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between new kana and old kana characters during a search. |
| `overtype` | `boolean` | r/w | Returns or sets if overtype mode is active. In overtype mode, the characters you type replace existing characters one by one. When overtype isn't active, the characters you type move existing text to the right. |
| `pagination` | `boolean` | r/w | Returns or sets if Microsoft Word repaginates documents in the background. |
| `paste adjust paragraph spacing` | `boolean` | r/w | Returns or sets if Microsoft Word automatically adjusts the spacing of paragraphs when cutting and pasting selections. |
| `paste adjust table formatting` | `boolean` | r/w | Returns or sets if Microsoft Word automatically adjusts the formatting of tables when cutting and pasting selections. |
| `paste adjust word spacing` | `boolean` | r/w | Returns or sets if Microsoft Word automatically adjusts the spacing of words when cutting and pasting selections. |
| `paste merge from Excel` | `boolean` | r/w | Returns or sets if text formatting will be merged when pasting from Microsoft Excel. |
| `paste merge from PowerPoint` | `boolean` | r/w | Returns or sets if text formatting will be merged when pasting from Microsoft PowerPoint. |
| `paste merge lists` | `boolean` | r/w | Returns or sets if the formatting of pasted lists will be merged with surrounding lists. |
| `paste smart cut paste` | `boolean` | r/w | Returns or sets if Microsoft Word intelligently pastes selections into a document. |
| `paste smart style behavior` | `boolean` | r/w | Returns or sets if Microsoft Word intelligently merges styles when pasting a selection from a different document. |
| `picture editor` | `text` | r/w | Returns or sets the name of the application to use to edit pictures. |
| `picture wrap types` | `WdWrapTypeMerged` | r/w | Returns or sets the wrapping that is used to insert or paste pictures. |
| `plain equations use plain text` | `boolean` | r/w | Returns or sets if equations are represented in plain text; false uses MathML |
| `print comments` | `boolean` | r/w | Returns or sets if Microsoft Word prints comments, starting on a new page at the end of the document. |
| `print drawing objects` | `boolean` | r/w | Returns or sets if Microsoft Word prints drawing objects. |
| `print field codes` | `boolean` | r/w | Returns or sets if Microsoft Word prints field codes instead of field results. |
| `print hidden text` | `boolean` | r/w | Returns or sets if hidden text is printed. |
| `print properties` | `boolean` | r/w | Returns or sets if Microsoft Word prints document summary information on a separate page at the end of the document. |
| `print reverse` | `boolean` | r/w | Returns or sets if Microsoft Word prints pages in reverse order. |
| `prolonged sound mark match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between short and long vowel sounds during a search. |
| `punctuation match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between types of punctuation marks during a search. |
| `replace selection` | `boolean` | r/w | Returns or sets if the result of typing or pasting replaces the selection. If false the result of typing or pasting is added before the selection, leaving the selection intact. |
| `revised lines color` | `WdColorIndex` | r/w | Returns or sets the color of changed lines in a document with tracked changes. |
| `revised lines mark` | `WdRevisedLinesMark` | r/w | Returns or sets the placement of changed lines in a document with tracked changes. |
| `revised properties color` | `WdColorIndex` | r/w | Returns or sets the color index used to mark formatting changes while change tracking is enabled. |
| `revised properties mark` | `WdRevisedPropertiesMark` | r/w | Returns or sets the mark used to show formatting changes while change tracking is enabled. |
| `save interval` | `integer` | r/w | Returns or sets the time interval in minutes for saving autorecover information. |
| `save normal prompt` | `boolean` | r/w | Returns or sets if Microsoft Word prompts the user for confirmation to save changes to the Normal template before it quits. if this is set to false Word automatically saves changes to the Normal template before it quits. |
| `save properties prompt` | `boolean` | r/w | Returns or sets if Microsoft Word prompts for document property information when saving a new document. |
| `send mail attach` | `boolean` | r/w | True if the send to command on the file menu inserts the active document as an attachment to a mail message. False if the send to command inserts the contents of the active document as text in a mail message. |
| `show readability statistics` | `boolean` | r/w | Returns or sets if Microsoft Word displays a list of summary statistics, including measures of readability, when it has finished checking grammar. |
| `small kana match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between diphthongs and double consonants during a search. |
| `smart cut paste` | `boolean` | r/w | Returns or sets if Microsoft Word automatically adjusts the spacing between words and punctuation when cutting and pasting occurs. |
| `smart paragraph selection` | `boolean` | r/w | Returns or sets if Microsoft Word includes the paragraph mark in a selection when selecting most or all of a paragraph. |
| `snap to grid` | `boolean` | r/w | Returns or sets if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in new documents. |
| `snap to shapes` | `boolean` | r/w | Returns or sets if Word automatically aligns autoshapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other autoshapes or East Asian characters in new documents. |
| `space match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between space markers used during a search. |
| `split cell color` | `WdCellColor` | r/w | Returns or sets the color of cells that are split while change tracking is enabled. |
| `suggest from main dictionary only` | `boolean` | r/w | Returns or sets if Microsoft Word draws spelling suggestions from the main dictionary only. If false it draws spelling suggestions from the main dictionary and any custom dictionaries that have been added. |
| `suggest spelling corrections` | `boolean` | r/w | Returns or sets if Microsoft Word always suggests alternative spellings for each misspelled word when checking spelling. |
| `tab indent key` | `boolean` | r/w | Returns or sets if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered and centered paragraphs to left-aligned. |
| `tc match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between some Japanese characters. |
| `update fields at print` | `boolean` | r/w | Returns or sets if Microsoft Word updates fields automatically before printing a document. |
| `update links at open` | `boolean` | r/w | Returns or sets if Microsoft Word automatically updates all embedded OLE links in a document when it's opened. |
| `update links at print` | `boolean` | r/w | Returns or sets if Microsoft Word updates fields automatically before printing a document. |
| `use character unit` | `boolean` | r/w | Returns or sets if Microsoft Word uses characters as the default measurement unit for the current document. |
| `use german spelling reform` | `boolean` | r/w | Returns or sets if Microsoft Word uses the German post-reform spelling rules when checking spelling. |
| `warn before saving printing sending markup` | `boolean` | r/w | Returns or sets if Microsoft Word displays a warning when saving, printing, or sending as e-mail a document containing comments or tracked changes. |
| `zj match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the distinction between some Japanese characters. |

### `add in`

Represents a single add-in, either installed or not installed.

*Plural:* `add ins`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `autoload` | `boolean` | r/o | Returns true if the specified add-in is automatically loaded when Word is started. |
| `compiled` | `boolean` | r/o | Returns true if the specified add-in is a Word add-in library. False if the add-in is a template. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `installed` | `boolean` | r/w | Returns or sets if the specified add-in is installed. |
| `name` | `text` | r/o | Returns the name of the add in. |
| `path` | `text` | r/o | Returns the disk or Web path to the specified add-in in HFS path style. |

### `application`

An AppleScript object representing the Microsoft Word application.

*Plural:* `applications`

**Contains:** `document`, `window`, `recent file`, `file converter`, `caption label`, `add in`, `command bar`, `template`, `key binding`, `dictionary`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active document` | `document` | r/o | Returns the currently active document object. |
| `active printer` | `text` | r/w | Returns or sets the name of the active printer |
| `active window` | `window` | r/o | Returns the currently active window object. |
| `animations` | `boolean` | r/w | Returns or sets if animations are enabled while a script is running. |
| `application version` | `text` | r/o | Returns the Microsoft Word version number as a string. |
| `autocorrect object` | `autocorrect` | r/o | Returns the autocorrect object |
| `automation security` | `mAuS` | r/w |  |
| `background printing status` | `integer` | r/o | Returns the number of print jobs in the background printing queue. |
| `browse extra file types` | `text` | r/w | Returns or sets to open hyperlinked HTML files in the Internet browser or Microsoft Word.  Set this property to text/html to allow hyperlinked HTML files to be opened in Microsoft Word. |
| `browser object` | `browser` | r/o | Returns the browser object.  The browser object is a tool used to move the insertion point to objects in a document. |
| `build` | `text` | r/o | Returns the version and build number of the Microsoft Word application. |
| `caps lock` | `boolean` | r/o | Returns if caps lock is turned on. |
| `caption` | `text` | r/o | Returns the name of the application. |
| `customization context` | `specifier` | r/w | Returns or set a template or document object that represents the template or document in which changes to menus, toolbars, and key bindings are stored. |
| `default save format` | `text` | r/w | Returns or sets the default format. Common settings include: document = WordDocument, document template = Template, Word 97-2004 document = Doc97, Word XML document = XML, web page = Html, Text only = Text, RTF = Rtf, unicode text = Unicode. |
| `default startup path` | `text` | r/w | Returns or sets the complete path of the startup folder, excluding the final separator. |
| `default table separator` | `text` | r/w | Returns or sets the single character used to separate text into cells when text is converted to a table. |
| `default web options object` | `default web options` | r/o | Returns the default web options object. |
| `display alerts` | `WdAlertLevel` | r/w | Returns or sets the way certain alerts or messages are handled while an AppleScript is running. |
| `display auto complete tips` | `boolean` | r/w | Returns or set if Microsoft Word displays tips that suggest text for completing words, dates, or phrases as you type. |
| `display recent files` | `boolean` | r/w | Returns or sets if the names of recently used files are displayed on the file menu. |
| `display screen tips` | `boolean` | r/w | Returns or set if comments, footnotes, endnotes, and hyperlinks are displayed as tips.  Text marked as having comments is highlighted. |
| `display scroll bars` | `boolean` | r/w | Returns or sets if Word displays a scroll bar in at least one document window.  Setting this property to true will display horizontal and vertical scrollbars in all windows.  Setting this property to false turns off all scrollbars in all windows. |
| `display status bar` | `boolean` | r/w | Returns or sets if the status bar is displayed. |
| `do print preview` | `boolean` | r/w | Returns or set if print preview is the current view. |
| `font names` | `text` | r/o | Returns the list of names of all of the available fonts |
| `keyboard script` | `integer` | r/w | Returns or sets the current keyboard script |
| `landscape font names` | `text` | r/o | Returns the list of names of all of the available landscape fonts |
| `localized language` | `integer` | r/o | Returns the Windows language ID for the locale that Microsoft Word is using. |
| `mailing label object` | `mailing label` | r/o | Returns the mailing label object. |
| `math ac` | `math autocorrect` | r/o | Returns a math autocorrect object that represents the auto correct entries for equations. |
| `name` | `text` | r/o | Returns the name of this application. |
| `normal template` | `template` | r/o | Returns the normal template object |
| `num lock` | `boolean` | r/o | Returns the state of the num lock key.  True if the keys on the numeric keypad insert numbers, false if the keys move the insertion point. |
| `path` | `text` | r/o | Returns the path to the application in HFS path style |
| `path separator` | `text` | r/o | Returns the character used to separate folder names. |
| `portrait font names` | `text` | r/o | Returns the list of names of all of the available portrait fonts |
| `selection` | `selection object` | r/o | Returns the selection object. |
| `settings` | `Word options` | r/o | Returns the word options object. |
| `show visual basic editor` | `boolean` | r/w | Return or set if the visual basic editor is visible. |
| `special mode` | `boolean` | r/o | Returns if Microsoft Word is in a special mode for example, copy text mode or move text mode. |
| `startup dialog` | `boolean` | r/w | Returns of sets if the project gallery dialog will be shown when starting Microsoft Word. |
| `status bar` | `text` | r/w | Displays the specified text in the status bar. This is a write only property. |
| `system_object` | `system object` | r/o | Returns the system object |
| `usable height` | `integer` | r/o | Returns the maximum height in points to which you can set the width of a Microsoft Window document window. |
| `usable width` | `integer` | r/o | Returns the maximum width in pixels to which you can set the width of a Microsoft Window document window. |
| `user address` | `text` | r/w | Returns of set the users mailing address. |
| `user control` | `boolean` | r/o | Returns true if the application was launched by a user.  False if the program was opened programmatically. |
| `user initials` | `text` | r/w | Returns of sets the initials, which Microsoft Word uses to construct comment marks. |
| `user name` | `text` | r/w | Returns or sets the users name, which is used on envelopes and for the author document property. |

### `auto text entry`

Represents a single AutoText entry.

*Plural:* `auto text entries`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto text value` | `text` | r/w | Returns or sets the value of this auto text entry. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/w | Returns or sets the name of this auto text entry. |
| `style name` | `text` | r/o | Returns the name of the style applied to the specified auto text entry. |

### `bookmark`

Represents a single bookmark.

*Plural:* `bookmarks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `column` | `boolean` | r/o | True if the specified bookmark is a table column. |
| `empty` | `boolean` | r/o | True if the specified bookmark is empty. An empty bookmark marks a location of a collapsed selection, it doesn't mark any text. |
| `end of bookmark` | `integer` | r/w | Returns or sets the ending character position of the bookmark. |
| `name` | `text` | r/o | The name of the bookmark object. |
| `start of bookmark` | `integer` | r/w | Returns or sets the starting character position of the bookmark. |
| `story type` | `WdStoryType` | r/o | Returns the story type for the bookmark. |
| `text object` | `text range` | r/o | Returns a text range object that represents the portion of a document that's referred to by the bookmark. |

### `border options`

Represents options associated with border object.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `always in front` | `boolean` | r/w | Returns or sets if page borders are displayed in front of the document text. |
| `distance from` | `WdBorderDistanceFrom` | r/w | Returns or sets a value that indicates whether the specified page border is measured from the edge of the page or from the text it surrounds. |
| `distance from bottom` | `integer` | r/w | Returns or sets the space in points between the text and the bottom border. |
| `distance from left` | `integer` | r/w | Returns or sets the space in points between the text and the left border. |
| `distance from right` | `integer` | r/w | Returns or sets the space in points between the right edge of the text and the right border. |
| `distance from top` | `integer` | r/w | Returns or sets the space in points between the text and the top border. |
| `enable borders` | `boolean` | r/w | Returns or sets border formatting for the specified object. |
| `enable first page in section` | `boolean` | r/w | Returns or sets if page borders are enabled for the first page in the section. |
| `enable other pages in section` | `boolean` | r/w | Returns or sets if page borders are enabled for all pages in the section except for the first page. |
| `has horizontal` | `boolean` | r/o | Returns true if a horizontal border can be applied to the object. |
| `has vertical` | `boolean` | r/o | Returns true if a vertical border can be applied to the object. |
| `inside color` | `integer` | r/w | Returns or sets the RGB color of the inside borders |
| `inside color index` | `WdColorIndex` | r/w | Returns or sets the color index of the inside borders. |
| `inside color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color of the inside borders. |
| `inside line style` | `WdLineStyle` | r/w | Returns or sets the inside border for the specified object. |
| `inside line width` | `WdLineWidth` | r/w | Returns or sets the line width of the inside border of an object. |
| `join borders` | `boolean` | r/w | Returns or sets if vertical borders at the edges of paragraphs and tables are removed so that the horizontal borders can connect to the page border. |
| `outside color` | `integer` | r/w | Returns or sets the RGB color of the outside borders |
| `outside color index` | `WdColorIndex` | r/w | Returns or sets the color index of the outside borders. |
| `outside color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color of the outside borders. |
| `outside line style` | `WdLineStyle` | r/w | Returns or sets the outside border for the specified object. |
| `outside line width` | `WdLineWidth` | r/w | Returns or sets the line width of the outside border of an object. |
| `shadow` | `boolean` | r/w | Returns or sets if the specified border is formatted as shadowed. |
| `surround footer` | `boolean` | r/w | Returns or sets if a page border encompasses the document footer. |
| `surround header` | `boolean` | r/w | Returns or sets if a page border encompasses the document header. |

### `border`

Represents a border of an object.

*Plural:* `borders`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `art style` | `WdPageBorderArt` | r/w | Returns or sets the graphical page-border design for a document. |
| `art width` | `integer` | r/w | Returns or sets the width in points of the specified graphical page border. |
| `color` | `integer` | r/w | Returns or sets the RGB color for the specified border object. |
| `color index` | `WdColorIndex` | r/w | Returns or sets the color index for the specified border. |
| `color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified border or font object. |
| `inside` | `boolean` | r/o | Returns true if an inside border can be applied to the specified object. |
| `line style` | `WdLineStyle` | r/w | Returns or sets the border line style for the specified object. |
| `line width` | `WdLineWidth` | r/w | Returns or sets the line width of an object's border. |
| `visible` | `boolean` | r/w | Returns or sets if the border object is visible |

### `browser`

Represents the browser tool used to move the insertion point to objects in a document. This tool is comprised of the three buttons at the bottom of the vertical scroll bar.

*Plural:* `browsers`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `browser target` | `WdBrowseTarget` | r/w | Returns or sets the document item that the previous and next methods locate. |

### `building block category`

*Plural:* `building block categories`

**Contains:** `building block`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `building block type` | `building block type` | r/o | Returns a building block type object that represents the type of building block for a building block category. |
| `entry_index` | `integer` | r/o | Returns an integer that represents the position of an item in a collection. |
| `name` | `text` | r/o | Returns the name of the specified object. |

### `building block type`

Represents a type of building block.

*Plural:* `building block types`

**Contains:** `building block category`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns an integer that represents the position of an item in a collection. |
| `name` | `text` | r/o | Returns a String that represents the localized name of a building block type. |

### `building block`

*Plural:* `building blocks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `building block type` | `building block type` | r/o | Returns a building block type object that represents the type for a building block. |
| `category` | `building block category` | r/o | Returns a Category object that represents the category for a building block. |
| `description` | `text` | r/w | Gets or sets a string that represents the description for a building block. |
| `entry_index` | `integer` | r/o | Returns an integer that represents the position of an item in a collection. |
| `identifier` | `text` | r/o | Returns a string that represents the internal identification number for a building block. |
| `insert options` | `integer` | r/w | Gets or sets an integer that represents how to insert the contents of a building block into a document. |
| `name` | `text` | r/w | Gets or sets a string that represents the name of a building block. |
| `value` | `text` | r/w | Gets or sets a string that represents the contents of a building block. |

### `caption label`

Represents a single caption label.

*Plural:* `caption labels`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `built in` | `boolean` | r/o | Returns true if this is a built-in caption label. |
| `caption label id` | `WdCaptionLabelID` | r/o | Returns a constant that represents the type for the specified caption label if the built in property of the caption label object is true. |
| `caption label position` | `WdCaptionPosition` | r/w | Returns or sets the position of caption label text. |
| `chapter style level` | `integer` | r/w | Returns or sets the heading style that marks a new chapter when chapter numbers are included with the specified caption label. The number 1 corresponds to Heading 1, 2 corresponds to Heading 2, and so on. |
| `include chapter number` | `boolean` | r/w | Returns or sets if a chapter number is included with page numbers or a caption label. |
| `name` | `text` | r/o | Returns the name for the caption label |
| `number style` | `WdCaptionNumberStyle` | r/w | Returns or sets the number style for the caption label object. |
| `separator` | `WdSeparatorType` | r/w | Returns or sets the character between the chapter number and the sequence number. |

### `check box`

Represents a single check box form field.

*Plural:* `check boxes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto size` | `boolean` | r/w | True sizes the check box according to the font size of the surrounding text. False sizes the check box according to the size property. |
| `check box default` | `boolean` | r/w | Returns or sets the default check box value. True if the default value is checked. |
| `check box value` | `boolean` | r/w | Returns or sets if the check box is selected.  True if the check box is selected. |
| `checkbox size` | `real` | r/w | Returns or sets the size of the specified check box in points. |
| `valid` | `boolean` | r/o | Returns if the check box object is valid. |

### `coauth lock`

Represents a co-authoring lock

*Plural:* `coauth locks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `is header_footer_lock` | `boolean` | r/w | returns if this lock is a header/footer lock. |
| `lock owner` | `coauthor` | r/o | Get the owner of the co-authoring lock. |
| `lock type` | `WdLockType` | r/w | Get the type of the co-authoring lock. |
| `text object` | `text range` | r/o | Gets a text range object that represents the portion of a document that contains the co-authoring lock. |

### `coauth update`

Represents a co-authoring update

*Plural:* `coauth updates`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `text object` | `text range` | r/o | Gets a text range object that represents the portion of a document that contains the co-authoring update. |

### `coauthor`

Represents a coauthor

*Plural:* `coauthors`

**Contains:** `coauth lock`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `coauthor email address` | `text` | r/o | The email address of coauthor. |
| `coauthor id` | `text` | r/o | The ID of coauthor. |
| `coauthor name` | `text` | r/o | The name of coauthor. |
| `is me` | `boolean` | r/w | returns if this coauthor is me. |

### `coauthoring`

Represents a coauthoring

**Contains:** `coauthor`, `coauth lock`, `coauth update`, `conflict`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `can merge` | `boolean` | r/w | returns if coauthoring can merge. |
| `can share` | `boolean` | r/w | returns if coauthoring can share. |
| `myself` | `coauthor` | r/w | returns me as author. |
| `pending updates` | `boolean` | r/w | returns if any updates are pending. |

### `conflict`

Represents a conflict

*Plural:* `conflicts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `conflict type` | `WdRevisionType` | r/o | Returns the revision type. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the conflict |

### `custom label`

Represents a custom mailing label.

*Plural:* `custom labels`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `dot matrix` | `boolean` | r/o | True if the printer type for the specified custom label is dot matrix. False if the printer type is either laser or ink jet. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `horizontal pitch` | `real` | r/w | Returns or sets the horizontal distance in points between the left edge of one custom mailing label and the left edge of the next mailing label. |
| `name` | `text` | r/w | Returns or set the name of the custom mailing label. |
| `number across` | `integer` | r/w | Returns or sets the number of custom mailing labels across a page. |
| `number down` | `integer` | r/w | Returns or sets the number of custom mailing labels down the length of a page. |
| `page size` | `WdCustomLabelPageSize` | r/w | Returns or sets the page size for the specified custom mailing label. |
| `side margin` | `real` | r/w | Returns or sets the side margin widths in points for the specified custom mailing label. |
| `top margin` | `real` | r/w | Returns or sets the distance in points between the top edge of the page and the top boundary of the body text. |
| `valid` | `boolean` | r/o | True if the various properties for example, height, width, and number down for the specified custom label work together to produce a valid mailing label. |
| `vertical pitch` | `real` | r/w | Returns or sets the vertical distance between the top of one mailing label and the top of the next mailing label. |
| `width` | `real` | r/w | Returns or sets the width of the object. |

### `data merge data field`

Represents a single data merge field in a data source.

*Plural:* `data merge data fields`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `data merge data field value` | `text` | r/o | Returns the contents of the mail merge data field or mapped data field for the current record. Use the active record property to set the active record in a data merge data source. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | Returns the name of the data merge data field object. |

### `data merge data source`

Represents the data merge data source in a data merge operation.

*Plural:* `data merge data sources`

**Contains:** `data merge field name`, `data merge data field`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active record` | `WdMailMergeActiveRecord` | r/w | Returns back the index of the current active record or an enumeration specifying the record |
| `connect string` | `text` | r/o | Returns the connection string for the specified data merge data source. |
| `first record` | `integer` | r/w | Returns or sets the number of the first data record to be merged in a data merge operation. |
| `header source name` | `text` | r/o | Returns the path and file name of the header source attached to the specified mail merge main document. |
| `header source type` | `WdMailMergeDataSource` | r/o | Returns a value that indicates the way the header source is being supplied for the mail merge operation. |
| `last record` | `integer` | r/w | Returns or sets the number of the last data record to be merged in a data merge operation. |
| `mail merge data source type` | `WdMailMergeDataSource` | r/o | Returns the type of data merge data source. |
| `name` | `text` | r/o | Returns the name of the data merge data source. |
| `query string` | `text` | r/w | Returns or sets the query string, SQL statement, used to retrieve a subset of the data in a data merge data source. |

### `data merge field name`

Represents a data merge field name in a data source.

*Plural:* `data merge field names`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | The name of the data merge field name object |

### `data merge field`

Represents a single data merge field in a document.

*Plural:* `data merge fields`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `data merge field range` | `text range` | r/w | Returns or sets a text range object that represents a field's code. A field's code is everything that's enclosed by the field characters including the leading space and trailing space characters. |
| `form field type` | `WdFieldType` | r/o | The type of this data merge field |
| `locked` | `boolean` | r/w | Returns or sets if the specified field is locked. When a field is locked, you cannot update the field results. |
| `next data merge field` | `data merge field` | r/o | Returns the next data merge field |
| `previous make merge field` | `data merge field` | r/o | Returns the previous data merge field |

### `data merge`

Represents the data merge functionality in Word.

*Plural:* `data merges`

**Contains:** `data merge field`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `data source` | `data merge data source` | r/o | Returns or sets the destination of the data merge results. |
| `destination` | `WdMailMergeDestination` | r/w | Returns or sets the destination of the data merge results. |
| `mail address field name` | `text` | r/w | Returns or sets the name of the field that contains e-mail addresses that are used when the data merge destination is electronic mail. |
| `mail as attachment` | `boolean` | r/w | Returns or sets if the merge documents are sent as attachments when the data merge destination is an e-mail message or a fax. |
| `mail subject` | `text` | r/w | Returns or sets the subject line used when the data merge destination is electronic mail. |
| `main document type` | `WdMailMergeMainDocType` | r/w | Returns or sets the data merge main document type. |
| `state` | `WdMailMergeState` | r/o | Returns the current state of a data merge operation. |
| `suppress blank lines` | `boolean` | r/w | Returns or sets if blank lines are suppressed when data merge fields in a data merge main document are empty. |
| `view data merge field codes` | `boolean` | r/w | Returns or sets if merge field names are displayed in a data merge main document. False if information from the current data record is displayed. |

### `default web options`

Contains global application-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page. You can return or set attributes either at the application global level or at the document level.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow png` | `boolean` | r/w | Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save a document as a Web page. |
| `always save in default encoding` | `boolean` | r/w | Returns or saves if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened.  The default value is False. |
| `check if office is htmleditor` | `boolean` | r/w | Returns or sets if Microsoft Word checks to see whether an Office application is the default HTML editor when you start Word. |
| `check if word is default htmleditor` | `boolean` | r/w | Returns or sets if Microsoft Word checks to see whether it is the default HTML editor when you start Word. The default value is true. |
| `encoding` | `MsoEncoding` | r/w | Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document |
| `pixels per inch` | `integer` | r/w | Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. |
| `screen size` | `MsoScreenSize` | r/w | Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser. |
| `update links on save` | `boolean` | r/w | Returns or sets if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved. |
| `use long file names` | `boolean` | r/w | Returns or sets if long file names are used when you save the document as a Web page. |

### `dialog`

Represents a built-in dialog box.

*Plural:* `dialogs`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `default dialog tab` | `WdWordDialogTab` | r/w | Returns or sets the active tab when the specified dialog box is displayed. |
| `dialog type` | `WdWordDialog` | r/o | The built-in dialog this object represents. |

### `document version`

Represents a single version of a document.

*Plural:* `document versions`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `comment` | `text` | r/o | Returns the comment associated with the specified version of a document. |
| `date value` | `date` | r/o | The date and time that the document version was saved. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `saved by` | `text` | r/o | Returns the name of the user who saved the specified version of the document. |

### `document`

Represents a Microsoft Word document.

*Plural:* `documents`

**Contains:** `document property`, `custom document property`, `bookmark`, `table`, `footnote`, `endnote`, `Word comment`, `section`, `paragraph`, `word`, `sentence`, `character`, `field`, `form field`, `Word style`, `frame`, `table of figures`, `variable`, `revision`, `table of contents`, `table of authorities`, `window`, `index`, `subdocument`, `story range`, `hyperlink object`, `shape`, `list template`, `Word list`, `inline shape`, `document version`, `readability statistic`, `grammatical error`, `spelling error`, `math object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active window` | `window` | r/o | Returns a window object that represents the active window, the window with the focus. If there are no windows open, an error occurs. |
| `attached template` | `template` | r/w | Returns of sets a template object that represents the template attached to the specified document. To set this property, specify either the name of the template or a template object. |
| `auto hyphenation` | `boolean` | r/w | Returns or sets if automatic hyphenation is turned on for the specified document. |
| `background shape` | `shape` | r/w | Returns or sets a shape object that represents the background image for the specified document. |
| `click and type paragraph style` | `WdBuiltinStyle` | r/w | Returns or sets the default paragraph style applied to text by the Click and Type feature in the specified document. |
| `coauthoring` | `coauthoring` | r/o |  |
| `compatible version` | `WdCompatibilityVersion` | r/w | Returns or sets the compatibility options to a given application version. |
| `consecutive hyphens limit` | `integer` | r/w | Returns or sets the maximum number of consecutive lines that can end with hyphens. If this property is set to zero, any number of consecutive lines can end with hyphens. |
| `data merge` | `data merge` | r/o | Returns the data merge object. |
| `default tab stop` | `real` | r/w | Returns or sets the interval in points between the default tab stops in the specified document. |
| `default table style` | `WdBuiltinStyle` | r/o | Returns the default table style for this document. |
| `document theme` | `office theme` | r/o | Returns an office theme object that represents the Microsoft Office theme applied to a document. |
| `document_type` | `WdDocumentType` | r/o | Returns the document type. |
| `east asian line break` | `WdFarEastLineBreakLevel` | r/w | Returns or sets the line break control level for the specified document. This property is ignored if the Ôøíeast asian line break controlÔøì property is set to false. Note that Ôøíeast asian line break controlÔøì is a paragraph format property. |
| `embed true type fonts` | `boolean` | r/w | Returns or set if Microsoft Word embeds TrueType fonts in a document when it's saved. This allow others to view the document with the same fonts that were used to create it. |
| `endnote options` | `endnote options` | r/o | Returns the endnote options object. |
| `envelope object` | `envelope` | r/o | Returns the envelop object. |
| `field options` | `field options` | r/o |  |
| `footnote options` | `footnote options` | r/o | Returns the footnote options object. |
| `full name` | `text` | r/o | Returns the full name of the document in HFS path style. |
| `grammar checked` | `boolean` | r/w | True if a grammar check has been run on the document. False if some of the document hasn't been checked for grammar. To recheck the grammar in the document, set the grammar checked property to false. |
| `grid distance horizontal` | `real` | r/w | Returns or sets the amount of horizontal space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize shape object or east asian characters in the document. |
| `grid distance vertical` | `real` | r/w | Returns or sets the amount of vertical space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize shape object or east asian characters in the document. |
| `grid origin from margin` | `boolean` | r/w | Returns or sets if Microsoft Word starts the character grid from the upper-left corner of the page. |
| `grid origin horizontal` | `real` | r/w | Returns or sets the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing shape object or east asian characters to begin in the document. |
| `grid origin vertical` | `real` | r/w | Returns or sets the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing shape object or east asian characters to begin in the document. |
| `grid space between horizontal lines` | `integer` | r/w | Returns or sets the interval at which Microsoft Word displays horizontal character gridlines in print layout view. |
| `grid space between vertical lines` | `integer` | r/w | Returns or sets the interval at which Microsoft Word displays vertical character gridlines in print layout view. |
| `has password` | `boolean` | r/o | True if a password is required to open the specified document. |
| `hyphenate caps` | `boolean` | r/w | Returns or sets if words in all capital letters can be hyphenated. |
| `hyphenation zone` | `integer` | r/w | Returns or sets the width of the hyphenation zone, in points. The hyphenation zone is the maximum amount of space that Microsoft Word leaves between the end of the last word in a line and the right margin. |
| `integrals uses subscripts` | `boolean` | r/o | Gets or sets a value that specifies the default location of limits for integrals. |
| `is master document` | `boolean` | r/o | True if the specified document is a master document. A master document includes one or more subdocuments. |
| `is subdocument` | `boolean` | r/o | True if the specified document is opened in a separate document window as a subdocument of a master document. |
| `justification mode` | `WdJustificationMode` | r/w | Returns or sets the character spacing adjustment for the specified document. |
| `letter content` | `letter content` | r/w | Return or sets the letter content object associated with the document. |
| `math binary operator break` | `WdOMathBreakBin` | r/o | Gets or sets a value that specifies where Microsoft Word places binary operators when equations span two or more lines. |
| `math default justification` | `WdOMathJc` | r/o | Gets or sets a value that indicates the default justification. |
| `math font name` | `text` | r/o | Gets or sets the name of the font that is used in a document to display equations. |
| `math left margin` | `real` | r/o | Gets or sets a value that specifies the left margin for equations. |
| `math right margin` | `real` | r/o | Gets or sets a value that specifies the right margin for equations. |
| `math subtraction operator` | `WdOMathBreakSub` | r/o | Gets or sets a value that specifies how Microsoft Word handles a subtraction operator that falls before a line break. |
| `math wrap` | `real` | r/o | Gets or sets a value that specifies the placement of the second line of an equation that wraps to a new line. |
| `name` | `text` | r/o | Returns the name of the document. |
| `nary uses subscripts` | `boolean` | r/o | Gets or sets a value that specifies the default location of limits for n-ary objects other than integrals |
| `no line break after` | `text` | r/w | Returns or sets the kinsoku characters after which Microsoft Word will not break a line. |
| `no line break before` | `text` | r/w | Returns or sets the kinsoku characters before which Microsoft Word will not break a line. |
| `page setup` | `page setup` | r/w | Returns or sets the page setup object. |
| `password` | `text` | r/w | Sets a password that must be supplied to open the specified document. This is write-only property |
| `path` | `text` | r/o | Returns the path to the document in HFS path style. |
| `posix full name` | `text` | r/o | Returns the full name of the document in POSIX path style. |
| `print forms data` | `boolean` | r/w | Returns or sets if Microsoft Word prints onto a preprinted form only the data entered in the corresponding online form. |
| `print post script over text` | `boolean` | r/w | Returns or sets if PRINT field instructions such as PostScript commands in a document are to be printed on top of text and graphics when a PostScript printer is used. |
| `print revisions` | `boolean` | r/w | Returns or sets if revision marks are printed with the document. False if revision marks aren't printed that is, tracked changes are printed as if they'd been accepted. |
| `protection type` | `WdProtectionType` | r/o | Returns the protection type for the specified document. |
| `read only` | `boolean` | r/o | True if changes to the document cannot be saved to the original document. |
| `read only recommended` | `boolean` | r/w | Returns or set if Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only. |
| `remove date and time` | `boolean` | r/w | Returns or sets if Microsoft Word removes date and time from revisions upon saving a document. |
| `remove personal information` | `boolean` | r/w | Returns or sets if Microsoft Word removes all user information from comments, revisions, and the properties dialog box upon saving a document. |
| `save format` | `WdSaveFormat` | r/o | Returns the file format of the specified document or file converter. Will be a unique number that specifies an external file converter or a constant. |
| `save forms data` | `boolean` | r/w | Returns or sets if Microsoft Word saves the data entered in a form as a tab-delimited record for use in a database. |
| `save subset fonts` | `boolean` | r/w | Returns or sets if Microsoft Word saves a subset of the embedded TrueType fonts with the document. |
| `save versions on close` | `boolean` | r/w | Sets or returns whether or not versions are automatically saved when a document is closed. |
| `saved` | `boolean` | r/w | Returns or set the saved state. True if the specified document or template hasn't changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed. |
| `show Word comments by` | `text` | r/w | Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'. |
| `show grammatical errors` | `boolean` | r/w | Returns or sets if grammatical errors are marked by a wavy green line in the document. To view grammatical errors in your document, you must set the check grammar as you type property of the Word options class to true. |
| `show hidden bookmarks` | `boolean` | r/w | Returns or sets if hidden bookmarks are shown. |
| `show revisions` | `boolean` | r/w | Returns or sets if tracked changes in the specified document are shown on the screen. |
| `show spelling errors` | `boolean` | r/w | Returns or sets if Microsoft Word underlines spelling errors in the document.  To view spelling errors in your document, you must set the check grammar as you type property of the Word options class to true. |
| `snap to grid` | `boolean` | r/w | Returns or sets if shape object or east asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in the specified document. |
| `snap to shapes` | `boolean` | r/w | Returns or sets if Microsoft Word automatically aligns shape object or east asian characters with invisible gridlines that go through the vertical and horizontal edges of other shape object or east asian characters in the document. |
| `spelling checked` | `boolean` | r/w | True if a spelling check has been run on the document. False if some of the document hasn't been checked for spelling. To see if the document contains spelling errors, get the count of spelling errors for the document. |
| `subdocuments expanded` | `boolean` | r/w | Returns or set if the subdocuments in the document are expanded. |
| `text object` | `text range` | r/o | Returns a range object that represents the main document story. |
| `track revisions` | `boolean` | r/w | Returns or sets if changes are tracked in the document. |
| `unavailable fonts` | `text` | r/o | Returns a list of fonts used in the document that are not available on the current system. |
| `update styles on open` | `boolean` | r/w | Returns or sets if the styles in the specified document are updated to match the styles in the attached template each time the document is opened. |
| `use default math settings` | `boolean` | r/o | Gets or sets a value that indicates whether to use the default math settings when creating new equations. |
| `use small fractions` | `boolean` | r/o | Gets or sets a value that indicates whether to use small fractions in equations contained within the document. |
| `web options` | `web options` | r/o | Returns the web options object. |
| `write password` | `text` | r/w | Sets a password for saving changes to the specified document. This is write-only property |
| `write reserved` | `boolean` | r/o | True if the specified document is protected with a write password. |

### `drop cap`

Represents a dropped capital letter at the beginning of a paragraph.

*Plural:* `drop caps`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `distance from text` | `real` | r/w | Returns or sets the distance in points between the dropped capital letter and the paragraph text. |
| `drop position` | `WdDropPosition` | r/w | Returns or sets the position of a dropped capital letter. |
| `font name` | `text` | r/w | Returns or sets the name of the font for the dropped capital letter. |
| `lines to drop` | `integer` | r/w | Returns or sets the height in lines of the specified dropped capital letter. |

### `drop down`

Represents a drop-down form field that contains a list of items in a form.

*Plural:* `drop downs`

**Contains:** `list entry`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `drop down default` | `integer` | r/w | Returns or sets the default drop-down item. The first item in a drop-down form field is 1, the second item is 2, and so on. |
| `drop down value` | `integer` | r/w | Returns or sets the number of the selected item in a drop-down form field. |
| `valid` | `boolean` | r/o | Returns if the drop down object is valid. |

### `endnote options`

A representation of the options associated with endnotes.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `endnote continuation notice` | `text range` | r/o | Returns a text range object that represents the endnote continuation notice. |
| `endnote continuation separator` | `text range` | r/o | Returns a text range object that represents the endnote continuation separator. |
| `endnote location` | `WdEndnoteLocation` | r/w | Returns or sets the position of all endnotes. |
| `endnote number style` | `WdNoteNumberStyle` | r/w | Returns or sets the number style for endnotes. |
| `endnote numbering rule` | `WdNumberingRule` | r/w | Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. |
| `endnote separator` | `text range` | r/o | Returns a text range object that represents the endnote separator. |
| `endnote starting number` | `integer` | r/w | Returns or sets the starting endnote number. |

### `endnote`

Represents an endnote.

*Plural:* `endnotes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `note reference` | `text range` | r/o | Returns a text range object that represents a endnote mark. |
| `text object` | `text range` | r/o | Returns a text range object that represents the portion of a document that's contained in the endnote object. |

### `envelope`

Represents an envelope.

*Plural:* `envelopes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `address` | `text range` | r/o | Returns the envelope delivery address as a text range object. |
| `address from left` | `real` | r/w | Returns or sets the distance in points between the left edge of the envelope and the delivery address. |
| `address from top` | `real` | r/w | Returns or sets the distance in points between the top edge of the envelope and the delivery address. |
| `address style` | `Word style` | r/o | Returns a Word style object that represents the delivery address style for the envelope. |
| `default face up` | `boolean` | r/w | Returns or sets if envelopes are fed face up by default. |
| `default height` | `real` | r/w | Returns or sets the default envelope height, in points. |
| `default omit return address` | `boolean` | r/w | Returns or sets if the return address is omitted from envelopes by default. |
| `default orientation` | `WdEnvelopeOrientation` | r/w | Returns or sets the default orientation for feeding envelopes. |
| `default print FIMA` | `boolean` | r/w | Returns or sets if a Facing Identification Mark FIM-A to envelopes by default. A FIM-A code is used to presort courtesy reply mail. The default print barcode property must be set to true before this property is set. |
| `default print bar code` | `boolean` | r/w | Returns or sets if a POSTNET bar code is added to envelopes or mailing labels by default. For U.S. mail only. This property must be set to true before the default print FIMA property is set |
| `default size` | `text` | r/w | Returns or sets the default envelope size. If you set either the default height or default width property, the property  is automatically changed to return custom size. |
| `default width` | `real` | r/w | Returns or sets the default envelope width, in points. |
| `feed source` | `WdPaperTray` | r/w | Returns or sets the paper tray for the envelope. |
| `return address` | `text range` | r/o | Returns the envelope return address as a text range object. |
| `return address from left` | `real` | r/w | Returns or sets the distance in points between the left edge of the envelope and the return address. |
| `return address from top` | `real` | r/w | Returns or sets the distance in points between the top edge of the envelope and the return address. |
| `return address style` | `Word style` | r/o | Returns a Word style object that represents the return address style for the envelope. |

### `field options`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `locked` | `boolean` | r/w |  |

### `field`

Represents a field.

*Plural:* `fields`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `field code` | `text range` | r/w | Returns a text range object that represents a field's code. A field's code is everything that's enclosed by the field characters including the leading space and trailing space characters. |
| `field kind` | `WdFieldKind` | r/o | Returns the type of link for a field object. |
| `field text` | `text` | r/w | Returns or sets data in an ADDIN field. The data is not visible in the field code or result. It is only accessible by returning the value of the data property. If the field isn't an ADDIN field, this property will cause an error. |
| `field type` | `WdFieldType` | r/o | Returns the field type. |
| `inline shape` | `inline shape` | r/o | Returns the inline shape object associated with this field |
| `link format` | `link format` | r/o | Returns the link format object associated with this field object. |
| `locked` | `boolean` | r/w | Returns or sets if the specified field is locked. When a field is locked, you cannot update the field results. |
| `next field` | `field` | r/o | Returns the next field object. |
| `previous field` | `field` | r/o | Returns the previous field object. |
| `result range` | `text range` | r/w | Returns or sets a text range object that represents a field's result. You can access a field result without changing the view from field codes. |
| `show codes` | `boolean` | r/w | Returns or sets if field codes are displayed for the specified field instead of field results. |

### `file converter`

Represents a file converter that's used to open or save files.

*Plural:* `file converters`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `can open` | `boolean` | r/o | Returns true if the specified file converter is designed to open files. |
| `can save` | `boolean` | r/o | Returns true if the specified file converter is designed to save files. |
| `class name` | `text` | r/o | Returns a unique name that identifies the file converter. |
| `extensions` | `text` | r/o | Returns the file name extensions associated with the specified file converter object. |
| `format name` | `text` | r/o | Returns the name of the specified file converter. |
| `name` | `text` | r/o | Returns the name of the file converter. |
| `open format` | `integer` | r/o | Returns the file format of the specified file converter. It will be a unique number that represents an external file converter. |
| `path` | `text` | r/o | Returns the disk or Web path to the specified file converter in HFS path style. |
| `save format` | `integer` | r/o | Returns the file format of the specified document or file converter. It will be a unique number that specifies an external file converter. |

### `find`

Represents the criteria for a find operation.

*Plural:* `finds`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `content` | `text` | r/w | Returns or sets the text in the find object. |
| `font object` | `font` | r/o | Returns the font object associated with this find object. |
| `format` | `boolean` | r/w | Returns or set if formatting is included in the find operation. |
| `forward` | `boolean` | r/w | Returns or sets if the find operation searches forward through the document. False if it searches backward through the document. |
| `found` | `boolean` | r/o | True if the search produces a match. |
| `frame` | `frame` | r/o | Returns the frame object associated with the find object. |
| `highlight` | `integer` | r/w | Returns or sets if highlight formatting is included in the find criteria |
| `language ID` | `WdLanguageID` | r/w | Returns or sets the language for the find object |
| `language ID east asian` | `WdLanguageID` | r/w | Returns or sets an east asian language for the template. |
| `match all word forms` | `boolean` | r/w | Returns or sets if all forms of the text to find are found by the find operation for instance, if the text to find is sit, sat and sitting are found as well. |
| `match byte` | `boolean` | r/w | Returns or sets if Microsoft Word distinguishes between full-width and half-width letters or characters during a search. |
| `match case` | `boolean` | r/w | Returns or sets if the find operation is case sensitive. |
| `match fuzzy` | `boolean` | r/w | Returns or sets if Microsoft Word uses the nonspecific search options for Japanese text during a search. |
| `match sounds like` | `boolean` | r/w | Returns or sets if words that sound similar to the text to find are returned by the find operation. |
| `match whole word` | `boolean` | r/w | Returns or sets if the find operation locates only entire words and not text that's part of a larger word. |
| `match wildcards` | `boolean` | r/w | Returns or sets if the text to find contains wildcards. |
| `no proofing` | `boolean` | r/w | Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores. |
| `paragraph format` | `paragraph format` | r/w | Returns or sets the paragraph format object associated with the find object. |
| `replacement` | `replacement` | r/o | Returns the replacement object associated with the find object. |
| `style` | `WdBuiltinStyle` | r/w | Returns or sets the Word style associated with the find object. |
| `supplemental language ID` | `WdLanguageID` | r/w | Returns or sets the language for the text range object |
| `wrap` | `WdFindWrap` | r/w | Returns or sets what happens if the search begins at a point other than the beginning of the document and the end of the document is reached or vice versa if forward is set to false or if the search text isn't found in the specified selection or range. |

### `font`

Contains font attributes, such as font name, size, and color, for an object.

*Plural:* `fonts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `all caps` | `boolean` | r/w | Returns or sets if the font is formatted as all capital letters. |
| `animation` | `WdAnimation` | r/w | Returns or sets the type of animation applied to the font. |
| `ascii name` | `text` | r/w | Returns or sets the font used for Latin text characters with character codes from 0 through 127. |
| `bold` | `boolean` | r/w | Returns or sets if the font is formatted as bold. |
| `bold bi` | `boolean` | r/w | Returns or sets if the font is formatted as bold. |
| `border options` | `border options` | r/o | Returns back border options object associated with this text range object. |
| `color` | `integer` | r/w | Returns or sets the RGB color of the font object. |
| `color index` | `WdColorIndex` | r/w | Returns or sets the color for the font object using an index. |
| `color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified border or font object. |
| `complex script name` | `text` | r/w | Returns or sets a bidi name - complex script name. |
| `disable character space grid` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the number of characters per line for the corresponding font object. |
| `double strike through` | `boolean` | r/w | Returns or sets if the specified font is formatted as double strikethrough text. |
| `east asian name` | `text` | r/w | Returns or sets an East Asian font name. |
| `emboss` | `boolean` | r/w | Returns or sets if the specified font is formatted as embossed. |
| `emphasis mark` | `WdEmphasisMark` | r/w | Returns or sets the emphasis mark for a character or designated character string. |
| `engrave` | `boolean` | r/w | Returns or sets if the specified font is formatted as engraved. |
| `font position` | `integer` | r/w | Returns or sets the position of text in points relative to the base line. A positive number raises the text, and a negative number lowers it. |
| `font size` | `real` | r/w | Returns or sets the font size. |
| `hidden` | `boolean` | r/w | Returns or sets if the font is formatted as hidden text. |
| `italic` | `boolean` | r/w | Returns or sets if the font is formatted as italic. |
| `italic bi` | `boolean` | r/w | Returns or sets if the font is formatted as italic. |
| `kerning` | `real` | r/w | Returns or sets the minimum font size for which Microsoft Word will adjust kerning automatically. |
| `name` | `text` | r/w | Returns or sets the font name associated with this font object. |
| `other name` | `text` | r/w | Returns or sets the font used for characters with character codes from 128 through 255. |
| `outline` | `boolean` | r/w | Returns or sets if the specified font is formatted as outline. |
| `scaling` | `integer` | r/w | Returns or sets the scaling percentage applied to the font. This property stretches or compresses text horizontally as a percentage of the current size. The scaling range is from 1 through 600. |
| `shading` | `shading` | r/o | Returns the shading object associated with the font object. |
| `shadow` | `boolean` | r/w | Returns or sets if the specified font is formatted as shadowed. |
| `small caps` | `boolean` | r/w | Returns or sets if the font is formatted as small capital letters. |
| `spacing` | `real` | r/w | Returns or sets the spacing in points between characters. |
| `strike through` | `boolean` | r/w | Returns or sets if the font is formatted as strikethrough text. |
| `subscript` | `boolean` | r/w | Returns or sets if the font is formatted as subscript. |
| `superscript` | `boolean` | r/w | Returns or sets if the font is formatted as superscript. |
| `underline` | `WdUnderline` | r/w | Returns or sets the type of underline applied to the font. |
| `underline color` | `integer` | r/w | Returns or sets the RGB color of the underline for the font object. |
| `underline color theme index` | `MsoThemeColorIndex` | r/w | Returns a value specifying the color of the underline for the selected text. |

### `footnote options`

A representation of the options associated with footnotes.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `footnote continuation notice` | `text range` | r/o | Returns a text range object that represents the footnote continuation notice. |
| `footnote continuation separator` | `text range` | r/o | Returns a text range object that represents the footnote continuation separator. |
| `footnote location` | `WdFootnoteLocation` | r/w | Returns or sets the position of all footnotes. |
| `footnote number style` | `WdNoteNumberStyle` | r/w | Returns or sets the number style for footnotes. |
| `footnote numbering rule` | `WdNumberingRule` | r/w | Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. |
| `footnote separator` | `text range` | r/o | Returns a text range object that represents the footnote separator. |
| `footnote starting number` | `integer` | r/w | Returns or sets the starting footnote number. |

### `footnote`

Represents a footnote positioned at the bottom of the page or beneath text.

*Plural:* `footnotes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `note reference` | `text range` | r/o | Returns a text range object that represents a footnote mark. |
| `text object` | `text range` | r/o | Returns a text range object that represents the portion of a document that's contained in the footnote object. |

### `form field`

Represents a single form field.

*Plural:* `form fields`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `calculate on exit` | `boolean` | r/w | Returns or sets if references to the specified form field are automatically updated whenever the field is exited. |
| `check box` | `check box` | r/o | Returns the check box object associated with the form filed object. |
| `drop down` | `drop down` | r/o | Returns the drop down object associated with the form filed object. |
| `enabled` | `boolean` | r/w | Returns or sets if this form field object is enabled. |
| `form field result` | `text` | r/w | Returns or sets a string that represents the result of the specified form field. |
| `form field type` | `WdFieldType` | r/o | The type of this form field. |
| `help text` | `text` | r/w | Returns or sets the text that's displayed in a message box when the form field has the focus and the user presses F1. |
| `name` | `text` | r/w | Returns or sets the name of the form field. |
| `next form field` | `form field` | r/o | Returns the next form field object. |
| `own help` | `boolean` | r/w | Returns or set the source of the text that's displayed in a message box when a form field has the focus and the user presses F1. If true, the text specified by the helptext property is displayed. If false, the text in the autotext entry is displayed |
| `own status` | `boolean` | r/w | Returns or sets the source of the text that's displayed in the status bar when a form field has the focus. If true, the text specified by the status text property is displayed. If false, the text of the autotext entry is displayed. |
| `previous form field` | `form field` | r/o | Returns the previous form field object. |
| `status text` | `text` | r/w | Returns or sets the text that's displayed in the status bar when a form field has the focus. |
| `text input` | `text input` | r/o | Returns the text input object associated with the form filed object. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the form field object. |

### `frame`

Represents a frame.

*Plural:* `frames`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border options` | `border options` | r/o | Returns back border options object associated with this frame object. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `height rule` | `WdFrameSizeRule` | r/w | Returns or sets the rule for determining the height of the specified frame. |
| `horizontal distance from text` | `real` | r/w | Returns or sets the horizontal distance between a frame and the surrounding text, in points. |
| `horizontal position` | `real` | r/w | Returns or sets the horizontal distance between the edge of the frame and the item specified by the relative horizontal position property. |
| `lock anchor` | `boolean` | r/w | Returns or sets if the specified frame is locked. The frame anchor indicates where the frame will appear in Draft view. You cannot reposition a locked frame anchor. |
| `relative horizontal position` | `WdRelativeHorizontalPosition` | r/w | Returns or sets what the horizontal position of a frame is relative. |
| `relative vertical position` | `WdRelativeVerticalPosition` | r/w | Returns or sets what the vertical position of a frame is relative. |
| `shading` | `shading` | r/o | Returns the shading object associated with the frame object. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the frame object. |
| `text wrap` | `boolean` | r/w | Returns or sets if document text wraps around the specified frame. |
| `vertical distance from text` | `real` | r/w | Returns or sets the vertical distance in points between a frame and the surrounding text. |
| `vertical position` | `real` | r/w | Returns or sets the vertical distance between the edge of the frame and the item specified by the relative vertical position property. |
| `width` | `real` | r/w | Returns or sets the width of the object. |
| `width rule` | `WdFrameSizeRule` | r/w | Returns or sets the rule used to determine the width of a frame. |

### `header footer`

Represents a single header or footer.

*Plural:* `header footers`

**Contains:** `page number`, `shape`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `header footer index` | `WdHeaderFooterIndex` | r/o | Returns a constant that represents the specified header or footer in a document or section. |
| `is header` | `boolean` | r/o | Returns true if this object is a header. |
| `link to previous` | `boolean` | r/w | Returns or sets if the specified header or footer is linked to the corresponding header or footer in the previous section. When a header or footer is linked, its contents are the same as in the previous header or footer. |
| `page number options` | `page number options` | r/o | Return the page number options object associated with this header footer object. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the header or footer. |

### `heading style`

Represents a style used to build a table of contents or figures.

*Plural:* `heading styles`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `level` | `integer` | r/w | Returns or sets the level for the heading style in a table of contents or table of figures. |
| `style` | `WdBuiltinStyle` | r/w | Returns or sets the style associated with the heading style object. |

### `hyperlink object`

Represents a hyperlink.

*Plural:* `hyperlink objects`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `email subject` | `text` | r/w | Returns or sets the text string for the specified hyperlinkÔøïs subject line. The subject line is appended to the hyperlinkÔøïs Internet address, or URL. |
| `extra info required` | `boolean` | r/o | Returns true if extra information is required to resolve the specified hyperlink. |
| `hyperlink address` | `text` | r/w | Returns or sets the address, for example, a file name or URL of the specified hyperlink. |
| `hyperlink type` | `MsoHyperlinkType` | r/o | The type of this hyperlink |
| `name` | `text` | r/o | The name of this hyperlink object. |
| `screen tip` | `text` | r/w | Returns or sets the text that appears as a screen tip when the mouse pointer is positioned over the specified hyperlink. |
| `shape` | `shape` | r/o |  |
| `sub address` | `text` | r/w | Returns or sets a named location in the destination of the specified hyperlink. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the hyperlink. |
| `text to display` | `text` | r/w | Returns or sets the specified hyperlink's visible text in a document. |

### `index`

Represents a single index.

*Plural:* `indexes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accented letters` | `boolean` | r/w | Returns or sets if the specified index contains separate headings for accented letters, for example, words that begin with Ôøã are under one heading and words that begin with A are under another. |
| `heading separator` | `WdHeadingSeparator` | r/w | Returns or sets the text between alphabetic groups, entries that start with the same letter in the index. |
| `index filter` | `WdIndexFilter` | r/w | Returns or sets a value that specifies how Microsoft Word classifies the first character of entries in the specified index. |
| `index type` | `WdIndexType` | r/w | Returns or sets the index type. |
| `number of columns` | `integer` | r/w | Sets or returns the number of columns for each page of an index. |
| `right align page numbers` | `boolean` | r/w | Returns or sets if page numbers are aligned with the right margin in an index. |
| `sort by` | `WdIndexSortBy` | r/w | Returns or sets the sorting criteria for the specified index. |
| `tab leader` | `WdTabLeader` | r/w | Returns or sets the character between entries and their page numbers in an index. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the index |

### `key binding`

Represents a custom key assignment in the current context.

*Plural:* `key bindings`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `binding context` | `specifier` | r/o | Returns an object that represents the storage location of the specified key binding. This property can return a document or template object. |
| `binding key string` | `text` | r/o | Returns the key combination string for the specified keys. |
| `command` | `text` | r/o | Returns the command assigned to the specified key combination. |
| `command parameter` | `text` | r/o | Returns the command parameter assigned to the specified shortcut key. |
| `key category` | `WdKeyCategory` | r/o | Returns the type of item assigned to the specified key binding. |
| `key code` | `integer` | r/o | Returns a unique number for the first key in the specified key binding.  You create this number by using the build key code method. |
| `key_code_2` | `integer` | r/o | Returns a unique number for the second key in the specified key binding. You create this number by using the build key code method. |
| `protected` | `boolean` | r/o | Returns true if you cannot change the specified key binding in the customize keyboard dialog box. |

### `letter content`

Represents the elements of a letter created by the letter wizard.

*Plural:* `letter contents`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `attention line` | `text` | r/w | Returns or sets the attention line text for a letter created by the letter wizard. |
| `cc list` | `text` | r/w | Returns or sets the carbon copy recipients for a letter created by the letter wizard. |
| `closing` | `text` | r/w | Returns or sets the closing text for a letter created by the letter wizard, for example, Sincerely yours. |
| `date format` | `text` | r/w | Returns or sets the date for a letter created by the letter wizard. |
| `enclosure count` | `integer` | r/w | Returns or sets the number of enclosures for a letter created by the letter wizard. |
| `include header footer` | `boolean` | r/w | Returns or sets if the header and footer from the page design template are included in a letter created by the letter wizard. |
| `letter style` | `WdLetterStyle` | r/w | Returns or sets the layout of a letter created by the letter wizard. |
| `letterhead` | `boolean` | r/w | Returns or sets if space is reserved for a preprinted letterhead in a letter created by the letter wizard. |
| `letterhead location` | `WdLetterheadLocation` | r/w | Returns or sets the location of the preprinted letterhead in a letter created by the letter wizard. |
| `letterhead size` | `real` | r/w | Returns or sets the amount of space in points to be reserved for a preprinted letterhead in a letter created by the letter wizard. |
| `mailing instructions` | `text` | r/w | Returns or sets the mailing instruction text for a letter created by the letter wizard, for example, Certified Mail. |
| `page design` | `text` | r/w | Returns or sets the name of the template attached to the document created by the letter wizard. |
| `recipient address` | `text` | r/w | Returns or sets the return address for a letter created with the letter wizard. |
| `recipient name` | `text` | r/w | Returns or sets the name of the person who'll be receiving the letter created by the letter wizard. |
| `recipient reference` | `text` | r/w | Returns or sets the reference line, for example, In reply to: for a letter created by the letter wizard. |
| `return address` | `text` | r/w | Returns or sets the return address for a letter created with the letter wizard. |
| `salutation` | `text` | r/w | Returns or sets the salutation text for a letter created by the letter wizard. |
| `salutation type` | `WdSalutationType` | r/w | Returns or sets the type of salutation for a letter created by the letter wizard. |
| `sender city` | `text` | r/o |  |
| `sender company` | `text` | r/w | Returns or sets the company name of the person creating a letter with the letter wizard. |
| `sender initials` | `text` | r/w | Returns or sets the initials of the person creating a letter with the letter wizard. |
| `sender job title` | `text` | r/w | Returns or sets the job title of the person creating a letter with the letter wizard. |
| `sender name` | `text` | r/w | Returns or sets the name of the person creating a letter with the letter wizard. |
| `subject` | `text` | r/w | Returns or sets the subject text of a letter created by the letter wizard. |

### `line numbering`

Represents line numbers in the left margin or to the left of each newspaper-style column.

*Plural:* `line numberings`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active line` | `boolean` | r/w | Returns or sets if line numbering is active for the specified document, section, or sections. |
| `count by` | `integer` | r/w | Returns or sets the numeric increment for line numbers. For example, if the count by property is set to 5, every fifth line will display the line number. Line numbers are only displayed in print layout view and print preview. |
| `distance from text` | `real` | r/w | Returns or sets the distance in points between the right edge of line numbers and the left edge of the document text. |
| `restart mode` | `WdNumberingRule` | r/w | Returns or sets the way line numbering runsÔøë that is, whether it starts over at the beginning of a new page or section or runs continuously. |
| `starting number` | `integer` | r/w | Returns or sets the starting line number. |

### `link format`

Represents the linking characteristics for a picture.

*Plural:* `link formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto update` | `boolean` | r/w | Returns or sets if the specified link is updated automatically when the container file is opened or when the source file is changed. |
| `link type` | `WdLinkType` | r/o | Returns the link type. |
| `locked` | `boolean` | r/w | Returns or sets if inline shape object is locked to prevent automatic updating. |
| `save picture with document` | `boolean` | r/w | Returns or sets if the specified picture is saved with the document. |
| `source full name` | `text` | r/w | Returns or sets the path and name of the source file for the specified picture. |
| `source name` | `text` | r/o | Returns the name of the source file for the specified picture. |
| `source path` | `text` | r/o | Returns the path of the source file for the specified picture. |

### `list entry`

Represents an item in a drop-down form field.

*Plural:* `list entries`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/w | Returns or sets the name of this list entry object. |

### `list format`

Represents the list formatting attributes that can be applied to the paragraphs in a range.

*Plural:* `list formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `Word list` | `Word list` | r/o | Returns a Word list object that represents the first formatted list contained in the list format object. |
| `list level number` | `integer` | r/w | Returns or sets the list level for the first paragraph in the list format object. |
| `list picture bullet` | `inline shape` | r/w |  |
| `list string` | `text` | r/o | Returns a string that represents the appearance of the list value of the first paragraph in the range for the list format object. For example, the second paragraph in an alphabetic list would return B. |
| `list template` | `list template` | r/o | Returns a list template object that represents the list formatting for the list format object. |
| `list type` | `WdListType` | r/o | Returns the type of lists that are contained in the range for the list format object. |
| `list value` | `integer` | r/o | Returns the numeric value of the first paragraph in the range for the specified list format object. For example, the list value property applied to the second paragraph in an alphabetic list would return 2. |
| `single list` | `boolean` | r/o | Returns if the specified list format object contains only one list. |
| `single list template` | `boolean` | r/o | True if the entire list format object uses the same list template |

### `list gallery`

Represents a single gallery of list formats.

*Plural:* `list galleries`

**Contains:** `list template`

### `list level`

Represents a single list level, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list.

*Plural:* `list levels`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `font object` | `font` | r/o | Returns the font object associated with this list level object. |
| `linked style` | `text` | r/w | Returns or sets the name of the style that's linked to the specified list level object. |
| `list level alignment` | `WdListLevelAlignment` | r/w | Returns or sets a constant that represents the alignment for the list level of the list template. |
| `number format` | `text` | r/w | Returns or sets the number format for the specified list level. |
| `number position` | `real` | r/w | Returns or sets the position in points of the number or bullet for the specified list level object. |
| `number style` | `WdListNumberStyle` | r/w | Returns or sets the number style for the list level object. |
| `picture bullet` | `inline shape` | r/o |  |
| `reset on higher` | `integer` | r/w | Returns or sets the list level that must appear before the specified list level restarts numbering at 1. |
| `start at` | `integer` | r/w | Returns or sets the starting number for the specified list level object. |
| `tab position` | `real` | r/w | Returns or sets the tab position for the specified list level object. |
| `text position` | `real` | r/w | Returns or sets the position in points for the second line of wrapping text for the specified list level object. |
| `trailing character` | `WdTrailingCharacter` | r/w | Returns or sets the character inserted after the number for the specified list level. |

### `list template`

Represents a single list template that includes all the formatting that defines a list.

*Plural:* `list templates`

**Contains:** `list level`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | Returns or sets the name of this list template object. |
| `outline numbered` | `boolean` | r/w | Returns or sets if the specified list template object is outline numbered. |

### `mailing label`

Represents a mailing label.

**Contains:** `custom label`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `default label name` | `text` | r/w | Returns or sets the name for the default mailing label. |
| `default laser tray` | `WdPaperTray` | r/w | Returns or sets the default paper tray that contains sheets of mailing labels |
| `default print bar code` | `boolean` | r/w | Returns or sets if a POSTNET bar code is added to envelopes or mailing labels by default. For U.S. mail only. This property must be set to true before the default print FIMA property is set |

### `math accent`

Represents an equation that has an accent mark above the base.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `char` | `integer` | r/w | Gets or sets an integer that represents the accent character for the accent object. |
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |

### `math autocorrect entry`

Represents an individual entry in the Math AutoCorrect engine.

*Plural:* `math autocorrect entries`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `autocorrect value` | `text` | r/w | Gets or sets a string that represents the contents of an equation auto correct entry. |
| `entry_index` | `integer` | r/o | Gets an integer that represents the position of an item in the collection. |
| `name` | `text` | r/w | Gets or sets a string that represents the name of an equation auto correct entry. |

### `math autocorrect`

Represents the Math AutoCorrect feature in Microsoft Word.

**Contains:** `math autocorrect entry`, `math recognized function`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `replace text` | `boolean` | r/w | Gets or sets whether Microsoft Word automatically replaces strings in equations with the corresponding math AutoCorrect definitions. |
| `use outside equations` | `boolean` | r/w | Gets or sets whether Microsoft Word uses math autocorrect rules outside equations in a document. |

### `math bar`

Represents the mathematical overbar for an object in an equation.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `is top bar` | `boolean` | r/w | Gets or sets the position of a bar in a bar object. True specifies a mathematical overbar. False specifies a mathematical underbar |

### `math border box`

Represents an invisible box around an equation or part of an equation to which you can assign properties that affect the layout or mathematical formatting of the entire box.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bottom left to top right strikethrough` | `boolean` | r/w | Gets or sets if a diagonal strikethrough from lower left to upper right is drawn. |
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `hide bottom` | `boolean` | r/w | Gets or sets whether to hide the bottom border of an equation's bounding box. |
| `hide left` | `boolean` | r/w | Gets or sets whether to hide the left border of an equation's bounding box. |
| `hide right` | `boolean` | r/w | Gets or sets whether to hide the right border of an equation's bounding box. |
| `hide top` | `boolean` | r/w | Gets or sets whether to hide the top border of an equation's bounding box. |
| `horizontal strikethrough` | `boolean` | r/w | Gets or sets if a horizontal strikethrough is drawn. |
| `top left to bottom right strikethrough` | `boolean` | r/w | Gets or sets if a diagonal strikethrough from upper left to lower right is drawn. |
| `vertical strikethrough` | `boolean` | r/w | Gets or sets if vertical strikethrough is drawn. |

### `math box`

Represents an invisible box around an equation or part of an equation to which you can apply properties that affect the mathematical or formatting properties, such as line breaks.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `is differential` | `boolean` | r/w | Gets or sets whether the box acts as the mathematical differential, in which case the box receives the appropriate horizontal spacing for a differential. |
| `is operator emulator` | `boolean` | r/w | Gets or sets if the box and its contents behave as a single operator and inherits the properties of an operator. |
| `non breaking` | `boolean` | r/w | Gets or sets whether breaks are allowed inside the box object. |

### `math break`

Represents individual line breaks in an equation.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `align at` | `integer` | r/w | Gets or sets an integer that represents the operator in one line, to which to align consecutive lines in an equation. |
| `text range` | `text range` | r/o | Returns a text range object that represents the portion of a document that is contained in the specified object. |

### `math delimiter`

Represents a delimiter object, consisting of opening and closing delimiters, e.g. parentheses, braces, brackets, or vertical bars, and one or more elements contained inside the delimiters.

**Contains:** `math object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `beginning character` | `integer` | r/w | Gets or sets an Integer that represents the beginning delimiter character in a delimiter object. |
| `delimiter shape` | `WdOMathShapeType` | r/w | Gets or sets the appearance of delimiters, e.g. parentheses, braces, and brackets, in relationship to the content that they surround. |
| `ending character` | `integer` | r/w | Gets or sets an Integer that represents the ending delimiter character in a delimiter object. |
| `hide left delimiter` | `boolean` | r/w | Gets or sets whether to hide the opening delimiter in a delimiter object. |
| `hide right delimiter` | `boolean` | r/w | Gets or sets whether to hide the closing delimiter in a delimiter object. |
| `separator character` | `integer` | r/w | Gets or sets an Integer that represents the separator character in a delimiter object when the delimiter object contains two or more arguments. |
| `should grow` | `boolean` | r/w | Gets or sets whether delimiter characters grow to the full height of the arguments that they contain. |

### `math equation array`

Represents a mathematical equation array object, consisting of one or more equations that can be vertically justified as a unit respect to surrounding text on the line.

**Contains:** `math object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `distribute equally` | `boolean` | r/w | Gets or sets if the equations in an equation array are distributed equally within the margins of its container, such as a column, cell, or page width. |
| `row spacing` | `integer` | r/w | Gets or sets an integer that represents the spacing between the rows in an equation array. |
| `row spacing rule` | `WdOMathSpacingRule` | r/w | Gets or sets the spacing rule that defines spacing in an equation array. |
| `use max width` | `boolean` | r/w | Gets or sets whether the equations in an equation array are spaced to the maximum width of the equation array. |
| `vertical alignment` | `WdOMathVertAlignType` | r/w | Gets or sets the type of vertical alignment for an equation array with respect to the text that surrounds the array. |

### `math fraction`

Represents a fraction, consisting of a numerator and denominator separated by a fraction bar. The fraction bar can be horizontal or diagonal, depending on the fraction properties.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `denominator` | `math object` | r/o | Returns a math object object that represents the denominator for an equation that contains a fraction. |
| `fraction type` | `WdOMathFracType` | r/w | Gets or sets the layout of a fraction, whether it is stacked, skewed, linear, or without a fraction bar. |
| `numerator` | `math object` | r/o | Returns a math object object that represents the numerator for the fraction. |

### `math func`

Represents a math func object that represents a type of mathematical function that consists of a function name, such as sin or cos, and an argument.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `func name` | `math object` | r/o | Returns an OMath object that represents the name of a mathematical function, such as sin or cos. |

### `math function`

Represents a mathematical function or structure such as fractions, integrals, sums, and radicals.

*Plural:* `math functions`

**Contains:** `math object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accent` | `math accent` | r/o | Returns a math accent object that represents a base character with a combining accent mark. |
| `border box` | `math border box` | r/o | Returns a math border box object that represents a border drawn around an equation or part of an equation. |
| `box` | `math box` | r/o | Returns a math box object to which you can apply properties. |
| `delimiter` | `math delimiter` | r/o | Returns an math delimiter object that represents the delimiter function |
| `equation array` | `math equation array` | r/o | Returns a math equation array object that represents an equation array function. |
| `fraction` | `math fraction` | r/o | Returns a math fraction object that represents a fraction. |
| `func` | `math func` | r/o | Returns a math func object that represents a type of mathematical function. |
| `function type` | `WdOMathFunctionType` | r/o | Gets the type of the function. |
| `group char` | `math group char` | r/o | Returns a math group char object that represents a horizontal character placed above or below text in an equation. |
| `left scripts` | `math left scripts` | r/o |  |
| `lower limit` | `math lower limit` | r/o | Returns a math lower limit object that represents the lower limit for a function. |
| `math obj` | `math object` | r/o | Returns an OMath object that represents the equation. |
| `matrix` | `math matrix` | r/o | Returns a math matrix object that represents a mathematical matrix. |
| `nary` | `math nary` | r/o | Returns a math nary object that represents the n-ary operation. |
| `overbar` | `math bar` | r/o | Returns a math bar object that represents the mathematical overbar for an object. |
| `phantom` | `math phantom` | r/o | Returns a math phantom object that represents an object used for advanced layout of an equation. |
| `radical` | `math radical` | r/o | Returns a math radical object that represents the mathematical radical function. |
| `sub and super script` | `math sub and super script` | r/o | Returns a math sub and super script object that represents a mathematical subscript-superscript object that consists of a base, a subscript, and a superscript. |
| `subscript` | `math subscript` | r/o | Returns a math subscript object that represents the mathematical subscript function. |
| `superscript` | `math superscript` | r/o | Returns a math superscript object that represents the mathematical superscript function. |
| `text range` | `text range` | r/o | Returns a text range object that represents the portion of a document that is contained in the math function. |
| `upper limit` | `math upper limit` | r/o | Returns a math upper limit object that represents upper limit function. |

### `math group char`

Represents a group character object, consisting of a character drawn above or below text, often with the purpose of visually grouping items.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `align top` | `boolean` | r/w | Returns or sets whether the grouping character is aligned vertically with the surrounding text or whether the base text that is either above or below the grouping character is aligned vertically with the surrounding text. |
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `is on top` | `boolean` | r/w | Gets or sets whether the grouping character is placed above the base text of the group character object. False displays the group character under the base text. |
| `the character` | `integer` | r/w | Gets or sets an integer that represents the character placed above or below text in a group character object. |

### `math left scripts`

Represents an equation that contains a superscript or subscript to the left of the base.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `subscript` | `math object` | r/o | Returns a math object object that represents the subscript for a pre-sub-superscript object. |
| `superscript` | `math object` | r/o | Returns a math object object that represents the superscript for a pre-sub-superscript object. |

### `math lower limit`

Represents the lower limit mathematical construct, consisting of text on the baseline and reduced-size text immediately below it.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `limit` | `math object` | r/o | Returns a math object object that represents the limit of the lower limit object. |

### `math matrix column`

Represents a matrix column.

**Contains:** `math object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `column index` | `integer` | r/w | Gets an integer that represents the ordinal position of a matrix column within the collection of matrix columns. |
| `horizontal alignment` | `WdOMathHorizAlignType` | r/w | Gets or sets the horizontal alignment for arguments in a matrix column. |

### `math matrix row`

Represents a matrix row.

*Plural:* `math matrix rows`

**Contains:** `math object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `row index` | `integer` | r/w | Gets an integer that represents the ordinal position of a matrix row within the collection of matrix rows. |

### `math matrix`

Represents an equation matrix.

**Contains:** `math matrix row`, `math matrix column`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `column gap` | `integer` | r/w | Gets or sets an integer that represents the spacing between columns in a matrix. |
| `column gap rule` | `WdOMathSpacingRule` | r/w | Gets or sets the spacing rule for the space that appears between columns in a matrix. |
| `column spacing` | `integer` | r/w | Gets or sets an integer that represents the spacing for columns in a matrix. |
| `hide placeholders` | `boolean` | r/w | Gets or sets whether placeholders in a matrix are hidden from display. |
| `row spacing` | `integer` | r/w | Gets or sets an integer that represents the spacing for rows in a matrix. |
| `row spacing rule` | `WdOMathSpacingRule` | r/w | Gets or sets the spacing rule for rows in a matrix. |
| `vertical alignment` | `WdOMathVertAlignType` | r/w | Gets or sets the vertical alignment for a matrix. |

### `math nary`

Represents the mathematical n-ary object, consisting of an n-ary object, a base/operand, and optional upper limits and lower limits.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `hides lower limit` | `boolean` | r/w | Gets or sets whether to hide the lower limit of an n-ary operator. |
| `hides upper limit` | `boolean` | r/w | Gets or sets whether to hide the upper limit of an n-ary operator. |
| `lower limit` | `math object` | r/o | Returns a math object object that represents the lower limit of an n-ary operator. |
| `operator character` | `integer` | r/w | Gets or sets an integer that represents a character used as the n-ary operator. |
| `should grow` | `boolean` | r/w | Gets or sets whether n-ary operators grow to the full height of the arguments that they contain. |
| `upper limit` | `math object` | r/o | Returns a math object object that represents the upper limit of an n-ary operator. |
| `use sub and super script positioning` | `boolean` | r/w | Gets or sets the positioning of n-ary limits in the subscript-superscript or upper limit-lower limit position. |

### `math object`

Represents an equation

*Plural:* `math objects`

**Contains:** `math function`, `math break`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `align point` | `integer` | r/w | Gets or sets an integer that represents the character position of the alignment point in the equation. |
| `argument index` | `integer` | r/o | Gets an integer that specifies the argument index of this component relative to the containing math object. |
| `containing function` | `math function` | r/o | Gets the math function that represents the parent, or containing, function. |
| `display type` | `WdOMathType` | r/w | Sets or returns the display type of the equation. |
| `justification` | `WdOMathJc` | r/w | Gets or sets the justification for the equation. |
| `nesting level` | `integer` | r/o | Returns an integer that represents the nesting level for the math object. |
| `parent argument` | `math object` | r/o | Gets a math object that represents the parent, or containing, argument. |
| `parent column` | `math matrix column` | r/o | Gets a math matrix column object that represents the parent column in a matrix. |
| `parent object` | `math object` | r/o | Gets the math object that represents the parent object. |
| `parent row` | `math matrix row` | r/o | Gets a math matrix row object that represents the parent row in a matrix. |
| `script size` | `integer` | r/w | Gets or sets an integer that represents the script size of an argument, for example, text, script, or script-script. |
| `text range` | `text range` | r/o | Gets a text range object that represents the portion of a document that contains the equation. |

### `math phantom`

Represents a phantom object, which has two primary uses: 1. adding the spacing of the phantom base without displaying that base or 2. suppressing part of the glyph from spacing considerations.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `shows invisibles` | `boolean` | r/w | Gets or sets whether the contents of a phantom object are visible. |
| `smash` | `boolean` | r/w | Gets or sets if the contents of the phantom are visible but that the height is not taken into account in the spacing of the layout. |
| `transparent` | `boolean` | r/w | Gets or sets whether a phantom object is transparent. |
| `zero ascent` | `boolean` | r/w | Gets or sets whether the ascent of the phantom contents is ignored in the spacing of the layout. |
| `zero descent` | `boolean` | r/w | Gets or sets whether the descent of the phantom contents is ignored in the spacing of the layout. |
| `zero width` | `boolean` | r/w | Gets or sets whether the width of a phantom object is ignored in the spacing of the layout. |

### `math radical`

Represents the mathematical radical object, consisting of a radical, a base, and an optional degree.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `degree` | `math object` | r/o | Returns a math object object that represents the degree for a radical. |
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `hide degree` | `boolean` | r/w | Gets or sets whether to hide the degree for the radical. |

### `math recognized function`

*Plural:* `math recognized functions`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Gets an integer that represents the position of an item in the collection. |
| `name` | `text` | r/o | Gets a string that represents the name of an equation recognized function. |

### `math sub and super script`

Represents an equation with a base that contains a superscript or subscript.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `align scripts` | `boolean` | r/w | Gets or sets whether to horizontally align subscripts and superscripts in the sub-superscript object. |
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `subscript` | `math object` | r/o | Returns a math object object that represents the subscript for a subscript-superscript object. |
| `superscript` | `math object` | r/o | Returns a math object object that represents the superscript for a subscript-superscript object. |

### `math subscript`

Represents an equation with a base that contains a subscript.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `subscript` | `math object` | r/o | Returns a math object that represents the subscript for a subscript object. |

### `math superscript`

Represents an equation with a base that contains a superscript.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `superscript` | `math object` | r/o | Returns a math object object that represents the superscript for a superscript object. |

### `math upper limit`

Represents the upper limit mathematical construct, consisting of text on the baseline and reduced-size text immediately above it.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `equation base` | `math object` | r/o | Returns a math object object that represents the base of the equation object. |
| `limit` | `math object` | r/o | Returns a math object object that represents the limit of the upper limit object. |

### `page number options`

Represents options associated with page number objects

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `chapter page separator` | `WdSeparatorType` | r/w | Returns or sets the separator character used between the chapter number and the page number. |
| `heading level for chapter` | `integer` | r/w | Returns or sets the heading level style that's applied to the chapter titles in the document. Can be a number from zero through 8, corresponding to heading levels 1 through 9. |
| `include chapter number` | `boolean` | r/w | Returns or sets if a chapter number is included with page numbers. |
| `number style` | `WdPageNumberStyle` | r/w | Returns or sets the number style for the page number objects. |
| `restart numbering at section` | `boolean` | r/w | Returns or sets if page numbering starts at 1 again at the beginning of the specified section. |
| `show first page number` | `boolean` | r/w | Returns or sets if the page number appears on the first page in the section. |
| `starting number` | `integer` | r/w | Returns or sets the starting page number. |

### `page number`

Represents a page number in a header or footer.

*Plural:* `page numbers`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `WdPageNumberAlignment` | r/w | Returns or sets a constant that represents the alignment for the page number. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |

### `page setup`

Represents the page setup description.

*Plural:* `page setups`

**Contains:** `text column`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bottom margin` | `real` | r/w | Returns or sets the distance in points between the bottom edge of the page and the bottom boundary of the body text. |
| `chars line` | `real` | r/w | Returns or sets the number of characters per line in the document grid. |
| `different first page header footer` | `boolean` | r/w | Returns or set if a different header or footer is used on the first page. |
| `first page tray` | `WdPaperTray` | r/w | Returns or sets the paper tray to use for the first page of a document or section. |
| `footer distance` | `real` | r/w | Returns or sets the distance in points between the footer and the bottom of the page. |
| `gutter` | `real` | r/w | Returns or sets the amount in points of extra margin space added to each page in a document or section for binding. |
| `gutter position` | `WdGutterStyle` | r/w | Returns or sets on which side the gutter appears in a document. |
| `header distance` | `real` | r/w | Returns or sets the distance in points between the header and the top of the page. |
| `layout mode` | `WdLayoutMode` | r/w | Returns or sets the layout mode for the current document. |
| `left margin` | `real` | r/w | Returns or sets the distance in points between the left edge of the page and the left boundary of the body text. |
| `line between text columns` | `boolean` | r/w | Returns or sets if vertical lines appear between all the columns. |
| `line numbering` | `line numbering` | r/w | Returns or sets the line numbering object that represents the line numbers for the specified page setup object. |
| `lines page` | `real` | r/w | Returns or sets the number of lines per page in the document grid. |
| `mirror margins` | `boolean` | r/w | Returns or sets if the inside and outside margins of facing pages are the same width. |
| `odd and even pages header footer` | `boolean` | r/w | Returns or sets if the specified page setup object has different headers and footers for odd-numbered and even-numbered pages. |
| `orientation` | `WdOrientation` | r/w | Returns or sets the orientation of the page. |
| `other pages tray` | `WdPaperTray` | r/w | Returns or sets the paper tray to be used for all but the first page of a document or section. |
| `page height` | `real` | r/w | Returns or sets the height of the page in points. |
| `page width` | `real` | r/w | Returns or sets the width of the page in points. |
| `paper size` | `WdPaperSize` | r/w | Returns or sets the paper size. |
| `right margin` | `real` | r/w | Returns or sets the distance in points between the right edge of the page and the right boundary of the body text. |
| `section start` | `WdSectionStart` | r/w | Returns or sets the type of section break for the specified object. |
| `show grid` | `boolean` | r/w | Determines whether to show the grid. |
| `spacing between text columns` | `real` | r/w | Returns or sets the spacing in points between columns. |
| `suppress endnotes` | `boolean` | r/w | Returns or sets if endnotes are printed at the end of the next section that doesn't suppress endnotes. Suppressed endnotes are printed before the endnotes in that section. |
| `text columns evenly spaced` | `boolean` | r/w | Returns or sets if text columns are evenly spaced. |
| `top margin` | `real` | r/w | Returns or sets the distance in points between the top edge of the page and the top boundary of the body text. |
| `vertical alignment` | `WdVerticalAlignment` | r/w | Returns or sets the vertical alignment of text on each page in a document or section. |
| `width of text columns` | `real` | r/w | Returns or sets the width of all text columns |

### `pane`

Represents a window pane.

*Plural:* `panes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `browse to window` | `boolean` | r/w | Returns or sets if lines wrap at the right edge of the pane rather than at the right margin of the page. |
| `browse width` | `integer` | r/o | Returns the width in points of the area in which text wraps in the specified pane. This property works only when you're in web layout view. |
| `display rulers` | `boolean` | r/w | Returns or sets if rulers are displayed for the window |
| `display vertical ruler` | `boolean` | r/w | Returns or sets if vertical rulers are displayed for the window |
| `document` | `document` | r/o | Returns a document object associated with this pane. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `horizontal percent scrolled` | `integer` | r/w | Returns or sets the horizontal scroll position as a percentage of the pane width. |
| `minimum font size` | `integer` | r/w | Returns or sets the minimum font size in points displayed for the specified pane. |
| `next pane` | `pane` | r/o | Returns the next pane object. |
| `previous pane` | `pane` | r/o | Returns the previous pane object. |
| `selection` | `selection object` | r/o | Returns the selection object that represents a selected text range or the insertion point. |
| `vertical percent scrolled` | `integer` | r/w | Returns or sets the vertical scroll position as a percentage of the pane width. |
| `view` | `view` | r/o | Returns a view object that represents the view for the pane. |

### `range endnote options`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `endnote location` | `WdEndnoteLocation` | r/w | Returns or sets the position of endnotes in a range or selection. |
| `endnote number style` | `WdNoteNumberStyle` | r/w | Returns or sets the number style for endnotes in a range or selection. |
| `endnote numbering rule` | `WdNumberingRule` | r/w | Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. |
| `endnote starting number` | `integer` | r/w | Returns or sets the starting endnote number in a range or selection. |

### `range footnote options`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `footnote location` | `WdFootnoteLocation` | r/w | Returns or sets the position of footnotes in a range or selection. |
| `footnote number style` | `WdNoteNumberStyle` | r/w | Returns or sets the number style for footnotes in a range or selection. |
| `footnote numbering rule` | `WdNumberingRule` | r/w | Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. |
| `footnote starting number` | `integer` | r/w | Returns or sets the starting number for footnotes in a range or selection. |

### `recent file`

Represents a recently used file.

*Plural:* `recent files`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | The name of the recently used file. |
| `path` | `text` | r/o | Returns the disk or Web path to the recent file in HFS path style. |
| `read only` | `boolean` | r/w | Returns or sets if changes to the document cannot be saved to the original document. |

### `replacement`

Represents the replace criteria for a find-and-replace operation.

*Plural:* `replacements`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `content` | `text` | r/w | Returns or sets the text to replace in the specified range or selection. |
| `font object` | `font` | r/o | The font object associated with this replacement object. |
| `frame` | `frame` | r/o | The frame object associated with this replacement object. |
| `highlight` | `integer` | r/w | Returns or sets if highlight formatting is applied to the replacement text. |
| `language ID` | `WdLanguageID` | r/w | Returns or sets the language for the replacement object |
| `language ID east asian` | `WdLanguageID` | r/w | Returns or sets an east asian language for the template. |
| `no proofing` | `boolean` | r/w | Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores. |
| `paragraph format` | `paragraph format` | r/w | Returns or set the paragraph format object associated with this replacement object. |
| `style` | `WdBuiltinStyle` | r/w | Returns or sets the Word style associated with the replacement object. |

### `reviewer`

Property of View: a person who has made tracked changes in the viewed document.

*Plural:* `reviewers`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `visible` | `boolean` | r/w |  |

### `revision`

Represents a change marked with a revision mark.

*Plural:* `revisions`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `author` | `text` | r/o | Returns the name of the user who made the specified tracked change. |
| `cells` | `null` | r/o |  |
| `date value` | `date` | r/o | The date and time that the tracked change was made. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `format description` | `text` | r/o |  |
| `moved range` | `text range` | r/o |  |
| `revision type` | `WdRevisionType` | r/o | Returns the revision type. |
| `style` | `null` | r/o |  |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the revision |

### `selection object`

Represents the current selection in a window or pane. A selection represents either a selected or highlighted area in the document, or it represents the insertion point if nothing in the document is selected.

*Plural:* `selection objects`

**Contains:** `table`, `word`, `sentence`, `character`, `footnote`, `endnote`, `Word comment`, `cell`, `section`, `paragraph`, `field`, `form field`, `frame`, `bookmark`, `hyperlink object`, `column`, `row`, `inline shape`, `shape`, `math object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `IP at end of line` | `boolean` | r/o | Returns true if the insertion point is at the end of a line that wraps to the next line. False if the selection isn't collapsed, if the insertion point isn't at the end of a line, or if the insertion point is positioned before a paragraph mark. |
| `bookmark id` | `integer` | r/o | Returns the number of the bookmark that encloses the beginning of the selection. The number corresponds to the position of the bookmark in the document, 1 for the first bookmark, 2 for the second one, and so on. |
| `border options` | `border options` | r/o | Returns back border options object associated with the selection |
| `column options` | `column options` | r/o | Returns a column options object for this selection. |
| `column select mode` | `boolean` | r/w | Returns or set if column selection mode is active. When this mode is active, the letters COL appear on the status bar. |
| `content` | `text` | r/w | Returns or sets the text in the selection. |
| `document` | `document` | r/o | Returns the document object associated with the selection. |
| `endnote options` | `endnote options` | r/o | Returns the endnote options object for the selection. |
| `extend mode` | `boolean` | r/w | Returns or set if extend mode is active. |
| `find object` | `find` | r/o | Returns the find object associated with the selection. |
| `fit text width` | `real` | r/w | Returns or sets the width in the current measurement units in which Microsoft Word fits the text in the current selection. |
| `font object` | `font` | r/o | The font object associated with the selection. |
| `footnote options` | `footnote options` | r/o | Returns the footnote options object for the selection. |
| `formatted text` | `text range` | r/w | Returns or sets a text range object that includes the formatted text in the selection. |
| `header footer object` | `header footer` | r/o | Returns the header footer object associated with the selection. |
| `is end of row mark` | `boolean` | r/o | Returns true if the selection is collapsed and is located at the end-of-row mark in a table. |
| `language ID` | `WdLanguageID` | r/w | Returns or sets the language for the selection object |
| `language ID east asian` | `WdLanguageID` | r/w | Returns or sets an east asian language for the selection. |
| `no proofing` | `boolean` | r/w | Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores. |
| `orientation` | `WdTextOrientation` | r/w | Returns or sets the orientation of text in the selection when the text direction feature is enabled. |
| `page setup` | `page setup` | r/w | Returns or set the page setup object associated with the selection. |
| `paragraph format` | `paragraph format` | r/w | Returns or set the paragraph object associated with the selection. |
| `previous bookmark id` | `integer` | r/o | Returns the number of the last bookmark that starts before or at the same place as the selection, It returns zero if there's no corresponding bookmark. |
| `range endnote options` | `range endnote options` | r/o |  |
| `range footnote options` | `range footnote options` | r/o |  |
| `row options` | `row options` | r/o |  |
| `selection end` | `integer` | r/w | Returns or sets the ending character position of the selection. |
| `selection flags` | `WdSelectionFlags` | r/w | Returns or sets properties of the selection. |
| `selection is active` | `boolean` | r/o | Returns if the selection is active. |
| `selection start` | `integer` | r/w | Returns or sets the starting character position of the selection. |
| `selection type` | `WdSelectionType` | r/o | Returns the selection type. |
| `shading` | `shading` | r/o | Returns the shading object associated with the selection |
| `show Word comments by` | `text` | r/w | Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'. |
| `show hidden bookmarks` | `boolean` | r/w | Returns or sets if hidden bookmarks are included in the elements of the selection |
| `start is active` | `boolean` | r/w | Returns or sets if the beginning of the selection is active. If the selection is not collapsed to an insertion point, either the beginning or the end of the selection is active. |
| `story length` | `integer` | r/o | Returns the number of characters in the story that contains the selection |
| `story type` | `WdStoryType` | r/o | Returns the story type for the selection. |
| `style` | `WdBuiltinStyle` | r/w | Returns or sets the Word style associated with the selection object. |
| `supplemental language ID` | `WdLanguageID` | r/w | Returns or sets the language for the selection |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the selection |

### `subdocument`

Represents a subdocument within a document or range.

*Plural:* `subdocuments`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `has file` | `boolean` | r/o | Returns true if the specified subdocument has been saved to a file. |
| `level` | `integer` | r/o | Returns the heading level used to create the subdocument. |
| `locked` | `boolean` | r/w | Returns or sets if a subdocument in a master document is locked. |
| `name` | `text` | r/o | The name of the subdocument. |
| `path` | `text` | r/o | Returns the disk or Web path to the specified subdocument in HFS path style. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the subdocument. |

### `system object`

Contains information about the computer system.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `country` | `WdCountry` | r/o | Returns the country or region designation of the system |
| `cursor` | `WdCursorType` | r/w | Returns or sets the state of the pointer. |
| `horizontal resolution` | `integer` | r/o | Returns the horizontal display resolution, in pixels. |
| `operating system` | `text` | r/o | Returns the name of the current operating system. |
| `processor type` | `text` | r/o | Returns the type of processor that the system is using. |
| `system version` | `text` | r/o | Returns the version number of the operating system. |
| `vertical resolution` | `integer` | r/o | Returns the vertical display resolution, in pixels. |

### `tab stop`

Represents a single tab stop.

*Plural:* `tab stops`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `WdTabAlignment` | r/w | Returns or sets a constant that represents the alignment for the specified tab stop. |
| `custom tab` | `boolean` | r/o | Returns if the specified tab stop is a custom tab stop. |
| `next tab stop` | `tab stop` | r/o | Returns the next tab stop object |
| `previous tab stop` | `tab stop` | r/o | Returns the previous tab stop object |
| `tab leader` | `WdTabLeader` | r/w | Returns or sets the leader for the specified tab stop object |
| `tab stop position` | `real` | r/w | Returns or sets the position of a tab stop relative to the left margin. |

### `table of authorities`

Represents a single table of authorities in a document

*Plural:* `tables of authorities`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `category` | `integer` | r/w | Returns or sets the category of entries to be included in a table of authorities. |
| `entry separator` | `text` | r/w | Returns or sets the characters up to five that separate a table of authorities entry and its page number. The default is a tab character with a dotted leader. |
| `include category header` | `boolean` | r/w | Returns or sets if the category name for a group of entries appears in the table of authorities. |
| `include sequence name` | `text` | r/w | Returns or sets the Sequence SEQ field identifier for a table of authorities. |
| `keep entry formatting` | `boolean` | r/w | Returns or sets if formatting from table of authorities entries is applied to the entries in the specified table of authorities. |
| `page number separator` | `text` | r/w | Returns of sets the characters up to five that separate individual page references in a table of authorities. The default is a comma and a space. |
| `page range separator` | `text` | r/w | Returns or sets the characters up to five that separate a range of pages in a table of authorities. The default is an en dash. |
| `passim` | `boolean` | r/w | Returns or sets if five or more page references to the same authority are replaced with Passim. |
| `separator` | `text` | r/w | Returns or sets the characters up to five between the sequence number and the page number. A hyphen is the default character. |
| `tab leader` | `WdTabLeader` | r/w | Returns or sets the character between entries and their page numbers in a table of authorities. |
| `table of authorities bookmark` | `text` | r/w | Returns or sets the name of the bookmark from which to collect table of authorities entries. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the table of authorities object. |

### `table of contents`

Represents a single table of contents in a document.

*Plural:* `tables of contents`

**Contains:** `heading style`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `include page numbers` | `boolean` | r/w | Returns or sets if page numbers are included in the table of contents. |
| `lower heading level` | `integer` | r/w | Returns or sets the ending heading level for a table of contents. |
| `right align page numbers` | `boolean` | r/w | Returns or sets if page numbers are aligned with the right margin in a table of contents. |
| `tab leader` | `WdTabLeader` | r/w | Returns or sets the character between entries and their page numbers in a table of contents. |
| `table id` | `text` | r/w | Returns or sets a one-letter identifier that's used to build a table of contents or table of figures from TOC fields. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the table of contents object. |
| `upper heading level` | `integer` | r/w | Returns or sets the starting heading level for a table of contents. |
| `use fields` | `boolean` | r/w | Returns or sets if table of contents entry fields are used to create a table of contents. |
| `use heading styles` | `boolean` | r/w | Returns or sets if built-in heading styles are used to create a table of contents. |

### `table of figures`

Represents a single table of figures in a document.

*Plural:* `tables of figures`

**Contains:** `heading style`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `caption` | `text` | r/w | Returns or sets the label that identifies the items to be included in a table of figures. |
| `include label` | `boolean` | r/w | Returns or sets if the caption label and caption number are included in a table of figures. |
| `include page numbers` | `boolean` | r/w | Returns or sets if page numbers are included in the table of figures. |
| `lower heading level` | `integer` | r/w | Returns or sets the ending heading level for a table of figures. |
| `right align page numbers` | `boolean` | r/w | Returns or sets if page numbers are aligned with the right margin in a table of figures. |
| `tab leader` | `WdTabLeader` | r/w | Returns or sets the character between entries and their page numbers in a table of figures. |
| `table id` | `text` | r/w | Returns or sets a one-letter identifier that's used to build a table of figures from TOC fields. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the table of figures object. |
| `upper heading level` | `integer` | r/w | Returns or sets the starting heading level for a table of figures. |
| `use fields` | `boolean` | r/w | Returns or sets if table of figures entry fields are used to create a table of figures. |
| `use heading styles` | `boolean` | r/w | Returns or sets if built-in heading styles are used to create a table of figures. |

### `template`

Represents a document template.

*Plural:* `templates`

**Contains:** `auto text entry`, `document property`, `custom document property`, `list template`, `building block`, `building block type`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `east asian line break` | `WdFarEastLineBreakLevel` | r/w | Returns or sets the line break control level for the specified template. This property is ignored if the Ôøíeast asian line break controlÔøì property is set to false. Note that Ôøíeast asian line break controlÔøì is a paragraph format property. |
| `full name` | `text` | r/o | Specifies the name of the template including the drive or Web path in HFS path style. |
| `justification mode` | `WdJustificationMode` | r/w | Returns or sets the character spacing adjustment for the specified template. |
| `language ID` | `WdLanguageID` | r/w | Returns or sets the language for the template object |
| `language ID east asian` | `WdLanguageID` | r/w | Returns or sets an east asian language for the template. |
| `name` | `text` | r/o | Returns the name of the template. |
| `no line break after` | `text` | r/w | Returns or sets the kinsoku characters after which Microsoft Word will not break a line. |
| `no line break before` | `text` | r/w | Returns or sets the kinsoku characters before which Microsoft Word will not break a line. |
| `no proofing` | `boolean` | r/w | Returns or sets if the spelling and grammar checker should ignore documents based on this template. |
| `path` | `text` | r/o | Returns the disk or Web path to the template file in HFS path style. |
| `posix full name` | `text` | r/o | Specifies the name of the template including the drive or Web path in POSIX path style. |
| `saved` | `boolean` | r/w | Return or set if the template hasn't changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed. |
| `template type` | `WdTemplateType` | r/o | Returns the template type. |

### `text column`

Represents a single text column.

*Plural:* `text columns`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `space after` | `real` | r/w | Returns or sets the amount of spacing in points after the text column. |
| `width` | `real` | r/w | Returns or sets the width of the object. |

### `text input`

Represents a single text form field.

*Plural:* `text inputs`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `default text input` | `text` | r/w | Returns or sets the text that represents the default text box contents. |
| `format` | `text` | r/o | Returns the formatting string for this text input object. |
| `text input field type` | `WdTextFormFieldType` | r/o | Returns the type of text form field. |
| `valid` | `boolean` | r/o | Returns if the text input object is valid. |
| `width` | `integer` | r/w | Returns or sets the width of the object. |

### `text retrieval mode`

Represents options that control how text is retrieved from a text range object.

*Plural:* `text retrieval modes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `include field codes` | `boolean` | r/w | Returns or sets if the text retrieved from the specified range includes field codes. |
| `include hidden text` | `boolean` | r/w | Returns or sets if the text retrieved from the specified range includes hidden text. |
| `view type` | `WdViewType` | r/w | Returns or sets the view for the text retrieval mode object. |

### `variable`

Represents a variable stored as part of a document. Document variables are used to preserve macro settings in between macro sessions.

*Plural:* `variables`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | The name of the variable |
| `variable value` | `text` | r/w | Returns or sets the value of a variable as a string. |

### `view`

Contains the view attributes, show all, field shading, table gridlines, and so on, for a window or pane.

*Plural:* `views`

**Contains:** `reviewer`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `browse to window` | `boolean` | r/w | Returns or sets if lines wrap at the right edge of the window rather than at the right margin of the page. |
| `conflict mode` | `boolean` | r/w | Returns or sets the Conflict Mode. |
| `data merge data view` | `boolean` | r/w | Returns or sets if data merge data is displayed instead of data merge fields in the specified window. |
| `draft` | `boolean` | r/w | Returns or sets if all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display. |
| `enlarge fonts less than` | `integer` | r/w | Returns or sets the point size below which screen fonts are automatically scaled to the larger size. |
| `field shading` | `WdFieldShading` | r/w | Returns or sets on-screen shading for form fields. |
| `full screen` | `boolean` | r/w | Returns or sets if the window is in full-screen view. |
| `magnifier` | `boolean` | r/w | Returns or sets if the pointer is displayed as a magnifying glass in print preview, indicating that the user can click to zoom in on a particular area of the page or zoom out to see an entire page or spread of pages. |
| `revisions mode` | `WdRevisionsMode` | r/w | Returns or sets the way Microsoft Word shows revisions. |
| `revisions view` | `WdRevisionsView` | r/w | Returns or sets whether Microsoft Word shows revisions based on the final or the original document. |
| `seek view` | `WdSeekView` | r/w | Returns or sets the document element displayed in print layout view. |
| `show all` | `boolean` | r/w | Returns or sets if all nonprinting characters such as hidden text, tab marks, space marks, and paragraph marks are displayed. |
| `show bookmarks` | `boolean` | r/w | Returns or sets if square brackets are displayed at the beginning and end of each bookmark. |
| `show comments` | `boolean` | r/w | Returns or sets if Microsoft Word displays comments. |
| `show drawings` | `boolean` | r/w | Returns or sets if objects created with the drawing tools are displayed in print layout view. |
| `show field codes` | `boolean` | r/w | Returns or sets if field codes are displayed. |
| `show first line only` | `boolean` | r/w | Returns or sets if only the first line of body text is shown in outline view. |
| `show format` | `boolean` | r/w | Returns or set if character formatting is visible in outline view. |
| `show format changes` | `boolean` | r/w | Returns or sets if Microsoft Word displays formatting changes. |
| `show hidden text` | `boolean` | r/w | Returns or set if text formatted as hidden text is displayed. |
| `show highlight` | `boolean` | r/w | Returns or sets if highlight formatting is displayed and printed with a document. |
| `show hyphens` | `boolean` | r/w | Returns or sets if optional hyphens are displayed. An optional hyphen indicates where to break a word when it falls at the end of a line. |
| `show insertions and deletions` | `boolean` | r/w | Returns or sets if Microsoft Word displays insertions and deletions. |
| `show main text layer` | `boolean` | r/w | Returns or sets if the text in the specified document is visible when the header and footer areas are displayed. |
| `show object anchors` | `boolean` | r/w | Returns or set if object anchors are displayed next to items that can be positioned in print layout view. |
| `show optional breaks` | `boolean` | r/w | Returns or sets if Microsoft Word displays optional line breaks. |
| `show other authors` | `boolean` | r/w | Returns or sets whether Microsoft Word shows other authors who are also editing the document. |
| `show paragraphs` | `boolean` | r/w | Returns or sets if paragraph marks are displayed. |
| `show picture place holders` | `boolean` | r/w | Returns or sets if blank boxes are displayed as placeholders for pictures. |
| `show revisions and comments` | `boolean` | r/w | Returns or sets if Microsoft Word displays revisions and comments. |
| `show spaces` | `boolean` | r/w | Returns or sets if space characters are displayed. |
| `show tabs` | `boolean` | r/w | Returns or sets if tab characters are displayed. |
| `show text boundaries` | `boolean` | r/w | Returns or sets if dotted lines are displayed around page margins, text columns, objects, and frames in print layout view. |
| `split special` | `WdSpecialPane` | r/w | Returns or sets the active window pane. |
| `table gridlines` | `boolean` | r/w | Returns or sets if table gridlines are displayed. |
| `view type` | `WdViewType` | r/w | Returns or sets the view type. |
| `wrap to window` | `boolean` | r/w | Returns or sets if lines wrap at the right edge of the document window rather than at the right margin or the right column boundary. |
| `zoom` | `zoom` | r/o | Returns the zoom object associated with this view object. |

### `web options`

Contains document-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow png` | `boolean` | r/w | Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save a document as a Web page. |
| `encoding` | `MsoEncoding` | r/w | Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document |
| `pixels per inch` | `integer` | r/w | Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. |
| `screen size` | `MsoScreenSize` | r/w | Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser. |
| `use long file names` | `boolean` | r/w | Returns or sets if long file names are used when you save the document as a Web page. |

### `window`

Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.

*Plural:* `windows`

**Contains:** `pane`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `IME mode` | `WdIMEMode` | r/w | Returns or sets the default start-up mode for the Japanese Input Method Editor, IME |
| `active` | `boolean` | r/o | Returns if the window is active. |
| `active pane` | `pane` | r/o | Returns a pane object that represents the active pane for the window |
| `caption` | `text` | r/w | Returns or sets the caption text for the document window. |
| `display horizontal scroll bar` | `boolean` | r/w | Returns or sets if the horizontal scroll bar is visible |
| `display rulers` | `boolean` | r/w | Returns or sets if rulers are displayed for the window |
| `display screen tips` | `boolean` | r/w | Returns or set if comments, footnotes, endnotes, and hyperlinks are displayed as tips.  Text marked as having comments is highlighted. |
| `display vertical ruler` | `boolean` | r/w | Returns or sets if vertical rulers are displayed for the window |
| `display vertical scroll bar` | `boolean` | r/w | Returns or sets if the vertical scroll bar is visible |
| `document` | `document` | r/o | Returns a document object associated with the pane. |
| `document map` | `boolean` | r/w | Returns or sets if the document map is visible. |
| `document map percent width` | `integer` | r/w | Returns or sets the width of the document map as a percentage of the width of the specified window. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `height` | `integer` | r/w | Returns or sets the height of the object. |
| `horizontal percent scrolled` | `integer` | r/w | Returns or sets the horizontal scroll position as a percentage of the document width. |
| `left position` | `integer` | r/w | Returns or sets the horizontal position of the object. |
| `next window` | `window` | r/o | Returns the next window object |
| `previous window` | `window` | r/o | Returns the previous window object |
| `selection` | `selection object` | r/o | Returns the selection object that represents a selected text range or the insertion point. |
| `split vertical` | `integer` | r/w | Returns or sets the vertical split percentage for the specified window. To remove the split, set this property to zero or set the split window property to false. |
| `split window` | `boolean` | r/w | Returns or sets if the window is split into multiple panes. When setting this to true the window is split into two equal-sized window panes. |
| `style area width` | `real` | r/w | Returns or sets the width of the style area in points. |
| `top` | `integer` | r/w | Returns or sets the vertical position of the object. |
| `vertical percent scrolled` | `integer` | r/w | Returns or sets the vertical scroll position as a percentage of the document width. |
| `view` | `view` | r/o | Returns a view object that represents the view for the window. |
| `width` | `integer` | r/w | Returns or sets the width of the object. |
| `window number` | `integer` | r/o | Returns the window number of the document displayed in the specified window. For example, if the caption of the window is Sales.doc:2, this property returns the number 2. |
| `window state` | `WdWindowState` | r/w | Returns or sets the state of the window. |
| `window type` | `WdWindowType` | r/o | Returns the window type. |

### `zoom`

Contains magnification options, for example, the zoom percentage for a window or pane.

*Plural:* `zooms`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `page columns` | `integer` | r/w | Returns or sets the number of pages to be displayed side by side on-screen at the same time in print layout view or print preview. |
| `page fit` | `WdPageFit` | r/w | Returns or sets the view magnification of a window so that either the entire page is visible or the entire width of the page is visible. |
| `page rows` | `integer` | r/w | Returns or sets the number of pages to be displayed one above the other on-screen at the same time in print layout view or print preview. |
| `percentage` | `integer` | r/w | Returns or sets the magnification for a window as a percentage. |

### Enumerations

### `WdMoveSelection`

| Value | Description |
|-------|-------------|
| `towards the end` |  |
| `towards the start` |  |
| `towards the left` |  |
| `towards the right` |  |
| `towards the top` |  |
| `towards the bottom` |  |

### `WdMoveUntilWhile`

| Value | Description |
|-------|-------------|
| `towards the tail` |  |
| `towards the beginning` |  |

### `WdInsertAboveBelow`

| Value | Description |
|-------|-------------|
| `above` |  |
| `below` |  |

### `WdInsertBeforeAfter`

| Value | Description |
|-------|-------------|
| `before the object` |  |
| `after the object` |  |

### `WdInsertRightLeft`

| Value | Description |
|-------|-------------|
| `insert on the right` |  |
| `insert on the left` |  |

### `WdInsertionPoint`

| Value | Description |
|-------|-------------|
| `offset` |  |
| `location reference` |  |

### `WdInsertionPoint`

| Value | Description |
|-------|-------------|
| `text range` |  |
| `insertion point` |  |

### `WdTemplateType`

| Value | Description |
|-------|-------------|
| `type normal template` |  |
| `type global template` |  |
| `type attached template` |  |

### `WdContinue`

| Value | Description |
|-------|-------------|
| `continue disabled` |  |
| `reset list` |  |
| `continue list` |  |

### `WdIMEMode`

| Value | Description |
|-------|-------------|
| `IME mode no control` |  |
| `IME mode on` |  |
| `IME mode off` |  |
| `IME mode hiragana` |  |
| `IME mode katakana` |  |
| `IME mode katakana half` |  |
| `IME mode alpha full` |  |
| `IME mode alpha` |  |
| `IME mode hangul full` |  |
| `IME mode hangul` |  |

### `WdBaselineAlignment`

| Value | Description |
|-------|-------------|
| `baseline align top` |  |
| `baseline align center` |  |
| `baseline align baseline` |  |
| `baseline align east asian50` |  |
| `baseline align auto` |  |

### `WdIndexFilter`

| Value | Description |
|-------|-------------|
| `index filter none` |  |
| `index filter aiueo` |  |
| `index filter akasatana` |  |
| `index filter chosung` |  |
| `index filter low` |  |
| `index filter medium` |  |
| `index filter full` |  |

### `WdIndexSortBy`

| Value | Description |
|-------|-------------|
| `index sort by stroke` |  |
| `index sort by syllable` |  |

### `WdJustificationMode`

| Value | Description |
|-------|-------------|
| `justification mode expand` |  |
| `justification mode compress` |  |
| `justification mode compress kana` |  |

### `WdFarEastLineBreakLevel`

| Value | Description |
|-------|-------------|
| `east asian line break level normal` |  |
| `east asian line break level strict` |  |
| `east asian line break level custom` |  |

### `WdMultipleWordConversionsMode`

| Value | Description |
|-------|-------------|
| `hangul to hanja` |  |
| `hanja to hangul` |  |

### `WdColorIndex`

| Value | Description |
|-------|-------------|
| `auto` |  |
| `black` |  |
| `blue` |  |
| `turquoise` |  |
| `bright green` |  |
| `pink` |  |
| `red` |  |
| `yellow` |  |
| `white` |  |
| `dark blue` |  |
| `teal` |  |
| `green` |  |
| `violet` |  |
| `dark red` |  |
| `dark yellow` |  |
| `gray50` |  |
| `gray25` |  |
| `by author` |  |
| `no highlight` |  |

### `WdColor`

| Value | Description |
|-------|-------------|
| `color automatic` |  |
| `color black` |  |
| `color blue` |  |
| `color pink` |  |
| `color red` |  |
| `color yellow` |  |
| `color turquoise` |  |
| `color bright green` |  |
| `color green` |  |
| `color white` |  |
| `color dark blue` |  |
| `color teal` |  |
| `color violet` |  |
| `color dark green` |  |
| `color dark red` |  |
| `color dark yellow` |  |
| `color brown` |  |
| `color olive green` |  |
| `color dark teal` |  |
| `color indigo` |  |
| `color orange` |  |
| `color blue gray` |  |
| `color light orange` |  |
| `color lime` |  |
| `color sea green` |  |
| `color aqua` |  |
| `color light blue` |  |
| `color gold` |  |
| `color sky blue` |  |
| `color plum` |  |
| `color rose` |  |
| `color tan` |  |
| `color light yellow` |  |
| `color light green` |  |
| `color light turquoise` |  |
| `color pale blue` |  |
| `color lavender` |  |
| `color gray05` |  |
| `color gray10` |  |
| `color gray125` |  |
| `color gray15` |  |
| `color gray20` |  |
| `color gray25` |  |
| `color gray30` |  |
| `color gray35` |  |
| `color gray375` |  |
| `color gray40` |  |
| `color gray45` |  |
| `color gray50` |  |
| `color gray55` |  |
| `color gray60` |  |
| `color gray625` |  |
| `color gray65` |  |
| `color gray70` |  |
| `color gray75` |  |
| `color gray80` |  |
| `color gray85` |  |
| `color gray875` |  |
| `color gray90` |  |
| `color gray95` |  |

### `WdTextureIndex`

| Value | Description |
|-------|-------------|
| `texture none` |  |
| `texture2 pt5 percent` |  |
| `texture5 percent` |  |
| `texture7 pt5 percent` |  |
| `texture10 percent` |  |
| `texture12 pt5 percent` |  |
| `texture15 percent` |  |
| `texture17 pt5 percent` |  |
| `texture20 percent` |  |
| `texture22 pt5 percent` |  |
| `texture25 percent` |  |
| `texture27 pt5 percent` |  |
| `texture30 percent` |  |
| `texture32 pt5 percent` |  |
| `texture35 percent` |  |
| `texture37 pt5 percent` |  |
| `texture40 percent` |  |
| `texture42 pt5 percent` |  |
| `texture45 percent` |  |
| `texture47 pt5 percent` |  |
| `texture50 percent` |  |
| `texture52 pt5 percent` |  |
| `texture55 percent` |  |
| `texture57 pt5 percent` |  |
| `texture60 percent` |  |
| `texture62 pt5 percent` |  |
| `texture65 percent` |  |
| `texture67 pt5 percent` |  |
| `texture70 percent` |  |
| `texture72 pt5 percent` |  |
| `texture75 percent` |  |
| `texture77 pt5 percent` |  |
| `texture80 percent` |  |
| `texture82 pt5 percent` |  |
| `texture85 percent` |  |
| `texture87 pt5 percent` |  |
| `texture90 percent` |  |
| `texture92 pt5 percent` |  |
| `texture95 percent` |  |
| `texture97 pt5 percent` |  |
| `texture solid` |  |
| `texture dark horizontal` |  |
| `texture dark vertical` |  |
| `texture dark diagonal down` |  |
| `texture dark diagonal up` |  |
| `texture dark cross` |  |
| `texture dark diagonal cross` |  |
| `texture horizontal` |  |
| `texture vertical` |  |
| `texture diagonal down` |  |
| `texture diagonal up` |  |
| `texture cross` |  |
| `texture diagonal cross` |  |

### `WdUnderline`

| Value | Description |
|-------|-------------|
| `underline none` |  |
| `underline single` |  |
| `underline words` |  |
| `underline double` |  |
| `underline dotted` |  |
| `underline thick` |  |
| `underline dash` |  |
| `underline dot dash` |  |
| `underline dot dot dash` |  |
| `underline wavy` |  |
| `underline wavy heavy` |  |
| `underline dotted heavy` |  |
| `underline dash heavy` |  |
| `underline dot dash heavy` |  |
| `underline dot dot dash heavy` |  |
| `underline dash long` |  |
| `underline dash long heavy` |  |
| `underline wavy double` |  |

### `WdEmphasisMark`

| Value | Description |
|-------|-------------|
| `emphasis mark none` |  |
| `emphasis mark over solid circle` |  |
| `emphasis mark over comma` |  |
| `emphasis mark over white circle` |  |
| `emphasis mark under solid circle` |  |

### `WdInternationalIndex`

| Value | Description |
|-------|-------------|
| `list separator` |  |
| `decimal separator` |  |
| `thousands separator` |  |
| `currency code` |  |
| `twenty four hour clock` |  |
| `international am` |  |
| `international pm` |  |
| `time separator` |  |
| `date separator` |  |
| `product language ID` |  |

### `WdAutoMacros`

| Value | Description |
|-------|-------------|
| `auto exec` |  |
| `auto new` |  |
| `auto open` |  |
| `auto close` |  |
| `auto exit` |  |
| `auto sync` |  |

### `WdCaptionPosition`

| Value | Description |
|-------|-------------|
| `caption position above` |  |
| `caption position below` |  |

### `WdCountry`

| Value | Description |
|-------|-------------|
| `us` |  |
| `canada` |  |
| `latin america` |  |
| `netherlands` |  |
| `france` |  |
| `spain` |  |
| `italy` |  |
| `uk` |  |
| `denmark` |  |
| `sweden` |  |
| `norway` |  |
| `germany` |  |
| `peru` |  |
| `mexico` |  |
| `argentina` |  |
| `brazil` |  |
| `chile` |  |
| `venezuela` |  |
| `japan` |  |
| `taiwan` |  |
| `china` |  |
| `korea` |  |
| `finland` |  |
| `iceland` |  |

### `WdHeadingSeparator`

| Value | Description |
|-------|-------------|
| `heading separator none` |  |
| `heading separator blank line` |  |
| `heading separator letter` |  |
| `heading separator letter low` |  |
| `heading separator letter full` |  |

### `WdSeparatorType`

| Value | Description |
|-------|-------------|
| `separator hyphen` |  |
| `separator period` |  |
| `separator colon` |  |
| `separator em dash` |  |
| `separator en dash` |  |

### `WdPageNumberAlignment`

| Value | Description |
|-------|-------------|
| `align page number left` |  |
| `align page number center` |  |
| `align page number right` |  |
| `align page number inside` |  |
| `align page number outside` |  |

### `WdBorderType`

| Value | Description |
|-------|-------------|
| `border top` |  |
| `border left` |  |
| `border bottom` |  |
| `border right` |  |
| `border horizontal` |  |
| `border vertical` |  |
| `border diagonal down` |  |
| `border diagonal up` |  |

### `WdFramePosition`

| Value | Description |
|-------|-------------|
| `frame top` |  |
| `frame left` |  |
| `frame bottom` |  |
| `frame right` |  |
| `frame center` |  |
| `frame inside` |  |
| `frame outside` |  |

### `WdAnimation`

| Value | Description |
|-------|-------------|
| `animation none` |  |
| `animation las vegas lights` |  |
| `animation blinking background` |  |
| `animation sparkle text` |  |
| `animation marching black ants` |  |
| `animation marching red ants` |  |
| `animation shimmer` |  |

### `WdCharacterCase`

| Value | Description |
|-------|-------------|
| `next case` |  |
| `lower case` |  |
| `upper case` |  |
| `title word` |  |
| `title sentence` |  |
| `toggle case` |  |
| `case half width` |  |
| `case full width` |  |
| `case katakana` |  |
| `case hiragana` |  |

### `WdSummaryLength`

| Value | Description |
|-------|-------------|
| ` 10 sentences` |  |
| ` 20 sentences` |  |
| ` 100 words` |  |
| ` 500 words` |  |
| ` 10 percent` |  |
| ` 25 percent` |  |
| ` 50 percent` |  |
| ` 75 percent` |  |

### `WdStyleType`

| Value | Description |
|-------|-------------|
| `style type paragraph` |  |
| `style type character` |  |
| `style type table` |  |
| `style type list` |  |
| `style type paragraph only` |  |
| `style type linked` |  |

### `WdUnits`

| Value | Description |
|-------|-------------|
| `a character item` |  |
| `a word item` |  |
| `a sentence item` |  |
| `a paragraph item` |  |
| `a line item` |  |
| `a story item` |  |
| `a screen` |  |
| `a section` |  |
| `a column` |  |
| `a row` |  |
| `a window` |  |
| `a cell` |  |
| `a character formatting` |  |
| `a paragraph formatting` |  |
| `a table` |  |
| `an item unit` |  |

### `WdGoToItem`

| Value | Description |
|-------|-------------|
| `goto a bookmark item` |  |
| `goto a section item` |  |
| `goto a page item` |  |
| `goto a table item` |  |
| `goto a line item` |  |
| `goto a footnote item` |  |
| `goto an endnote item` |  |
| `goto a comment item` |  |
| `goto a field item` |  |
| `goto a graphic` |  |
| `goto an object item` |  |
| `goto an equation` |  |
| `goto a heading item` |  |
| `goto a percent` |  |
| `goto a spelling error` |  |
| `goto a grammatical error` |  |
| `goto a proofreading error` |  |

### `WdGoToDirection`

| Value | Description |
|-------|-------------|
| `the first item` |  |
| `the last item` |  |
| `the next item` |  |
| `relative` |  |
| `the previous item` |  |
| `absolute` |  |

### `WdCollapseDirection`

| Value | Description |
|-------|-------------|
| `collapse start` |  |
| `collapse end` |  |

### `WdRowHeightRule`

| Value | Description |
|-------|-------------|
| `row height auto` |  |
| `row height at least` |  |
| `row height exactly` |  |

### `WdFrameSizeRule`

| Value | Description |
|-------|-------------|
| `frame auto` |  |
| `frame at least` |  |
| `frame exact` |  |

### `WdInsertCells`

| Value | Description |
|-------|-------------|
| `insert cells shift right` |  |
| `insert cells shift down` |  |
| `insert cells entire row` |  |
| `insert cells entire column` |  |

### `WdDeleteCells`

| Value | Description |
|-------|-------------|
| `delete cells shift left` |  |
| `delete cells shift up` |  |
| `delete cells entire row` |  |
| `delete cells entire column` |  |

### `WdListApplyTo`

| Value | Description |
|-------|-------------|
| `list apply to whole list` |  |
| `list apply to this point forward` |  |
| `list apply to selection` |  |

### `WdAlertLevel`

| Value | Description |
|-------|-------------|
| `alerts none` |  |
| `alerts message box` |  |
| `alerts all` |  |

### `WdCursorType`

| Value | Description |
|-------|-------------|
| `cursor wait` |  |
| `cursor ibeam` |  |
| `cursor normal` |  |
| `cursor northwest arrow` |  |

### `WdRulerStyle`

| Value | Description |
|-------|-------------|
| `adjust none` |  |
| `adjust proportional` |  |
| `adjust first column` |  |
| `adjust same width` |  |

### `WdParagraphAlignment`

| Value | Description |
|-------|-------------|
| `align paragraph left` |  |
| `align paragraph center` |  |
| `align paragraph right` |  |
| `align paragraph justify` |  |
| `align paragraph distribute` |  |
| `align paragraph justify med` |  |
| `align paragraph justify hi` |  |
| `align paragraph justify low` |  |
| `align paragraph thai justify` |  |

### `WdListLevelAlignment`

| Value | Description |
|-------|-------------|
| `list level align left` |  |
| `list level align center` |  |
| `list level align right` |  |

### `WdRowAlignment`

| Value | Description |
|-------|-------------|
| `align row left` |  |
| `align row center` |  |
| `align row right` |  |

### `WdTabAlignment`

| Value | Description |
|-------|-------------|
| `align tab left` |  |
| `align tab center` |  |
| `align tab right` |  |
| `align tab decimal` |  |
| `align tab bar` |  |
| `align tab list` |  |

### `WdVerticalAlignment`

| Value | Description |
|-------|-------------|
| `align vertical top` |  |
| `align vertical center` |  |
| `align vertical justify` |  |
| `align vertical bottom` |  |

### `WdCellVerticalAlignment`

| Value | Description |
|-------|-------------|
| `cell align vertical top` |  |
| `cell align vertical center` |  |
| `cell align vertical bottom` |  |

### `WdRangeAnnotationAlignment`

| Value | Description |
|-------|-------------|
| `range annotation alignment center` |  |
| `zero one zero` |  |
| `one two one` |  |
| `range annotation alignment left` |  |
| `range annotation alignment right` |  |

### `WdTrailingCharacter`

| Value | Description |
|-------|-------------|
| `trailing tab` |  |
| `trailing space` |  |
| `trailing none` |  |

### `WdListGalleryType`

| Value | Description |
|-------|-------------|
| `bullet gallery` |  |
| `number gallery` |  |
| `outline number gallery` |  |

### `WdListNumberStyle`

| Value | Description |
|-------|-------------|
| `list number style arabic` |  |
| `list number style uppercase roman` |  |
| `list number style lowercase roman` |  |
| `list number style uppercase letter` |  |
| `list number style lowercase letter` |  |
| `list number style ordinal` |  |
| `list number style cardinal text` |  |
| `list number style ordinal text` |  |
| `list number style kanji` |  |
| `list number style kanji digit` |  |
| `list number style aiueo half width` |  |
| `list number style iroha half width` |  |
| `list number style arabic full width` |  |
| `list number style kanji traditional` |  |
| `list number style kanji traditional2` |  |
| `list number style number in circle` |  |
| `list number style aiueo` |  |
| `list number style iroha` |  |
| `list number style arabic lz` |  |
| `list number style bullet` |  |
| `list number style ganada` |  |
| `list number style chosung` |  |
| `list number style gbnum1` |  |
| `list number style gbnum2` |  |
| `list number style gbnum3` |  |
| `list number style gbnum4` |  |
| `list number style zodiac1` |  |
| `list number style zodiac2` |  |
| `list number style zodiac3` |  |
| `list number style trad chin num1` |  |
| `list number style trad chin num2` |  |
| `list number style trad chin num3` |  |
| `list number style trad chin num4` |  |
| `list number style simp chin num1` |  |
| `list number style simp chin num2` |  |
| `list number style simp chin num3` |  |
| `list number style simp chin num4` |  |
| `list number style hanja read` |  |
| `list number style hanja read digit` |  |
| `list number style hangul` |  |
| `list number style hanja` |  |
| `list number style hebrew1` |  |
| `list number style arabic1` |  |
| `list number style hebrew2` |  |
| `list number style arabic2` |  |
| `list number style hindi letter1` |  |
| `list number style hindi letter2` |  |
| `list number style hindi arabic` |  |
| `list number style hindi cardinal text` |  |
| `list number style thai letter` |  |
| `list number style thai arabic` |  |
| `list number style thai cardinal text` |  |
| `list number style viet cardinal text` |  |
| `list number style lowercase russian` |  |
| `list number style uppercase russian` |  |
| `list number style lowercase greek` |  |
| `list number style uppercase greek` |  |
| `list number style arabic lz2` |  |
| `list number style arabic lz3` |  |
| `list number style arabic lz4` |  |
| `list number style lowercase turkish` |  |
| `list number style uppercase turkish` |  |
| `list number style lowercase bulgarian` |  |
| `list number style uppercase bulgarian` |  |
| `list number style picture bullet` |  |
| `list number style legal` |  |
| `list number style legal lz` |  |
| `list number style none` |  |

### `WdNoteNumberStyle`

| Value | Description |
|-------|-------------|
| `note number style arabic` |  |
| `note number style uppercase roman` |  |
| `note number style lowercase roman` |  |
| `note number style uppercase letter` |  |
| `note number style lowercase letter` |  |
| `note number style symbol` |  |
| `note number style arabic full width` |  |
| `note number style kanji` |  |
| `note number style kanji digit` |  |
| `note number style kanji traditional` |  |
| `note number style number in circle` |  |
| `note number style hanja read` |  |
| `note number style hanja read digit` |  |
| `note number style trad chin num1` |  |
| `note number style trad chin num2` |  |
| `note number style simp chin num1` |  |
| `note number style simp chin num2` |  |
| `note number style hebrew letter1` |  |
| `note number style arabic letter1` |  |
| `note number style hebrew letter2` |  |
| `note number style arabic letter2` |  |
| `note number style hindi letter1` |  |
| `note number style hindi letter2` |  |
| `note number style hindi arabic` |  |
| `note number style hindi cardinal text` |  |
| `note number style thai letter` |  |
| `note number style thai arabic` |  |
| `note number style thai cardinal text` |  |
| `note number style viet cardinal text` |  |

### `WdCaptionNumberStyle`

| Value | Description |
|-------|-------------|
| `caption number style arabic` |  |
| `caption number style uppercase roman` |  |
| `caption number style lowercase roman` |  |
| `caption number style uppercase letter` |  |
| `caption number style lowercase letter` |  |
| `caption number style arabic full width` |  |
| `caption number style kanji` |  |
| `caption number style kanji digit` |  |
| `caption number style kanji traditional` |  |
| `caption number style number in circle` |  |
| `caption number style ganada` |  |
| `caption number style chosung` |  |
| `caption number style zodiac1` |  |
| `caption number style zodiac2` |  |
| `caption number style hanja read` |  |
| `caption number style hanja read digit` |  |
| `caption number style trad chin num2` |  |
| `caption number style trad chin num3` |  |
| `caption number style simp chin num2` |  |
| `caption number style simp chin num3` |  |
| `caption number style hebrew letter1` |  |
| `caption number style arabic letter1` |  |
| `caption number style hebrew letter2` |  |
| `caption number style arabic letter2` |  |
| `caption number style hindi letter1` |  |
| `caption number style hindi letter2` |  |
| `caption number style hindi arabic` |  |
| `caption number style hindi cardinal text` |  |
| `caption number style thai letter` |  |
| `caption number style thai arabic` |  |
| `caption number style thai cardinal text` |  |
| `caption number style viet cardinal text` |  |

### `WdPageNumberStyle`

| Value | Description |
|-------|-------------|
| `page number style arabic` |  |
| `page number style uppercase roman` |  |
| `page number style lowercase roman` |  |
| `page number style uppercase letter` |  |
| `page number style lowercase letter` |  |
| `page number style arabic full width` |  |
| `page number style kanji` |  |
| `page number style kanji digit` |  |
| `page number style kanji traditional` |  |
| `page number style number in circle` |  |
| `page number style hanja read` |  |
| `page number style hanja read digit` |  |
| `page number style trad chin num1` |  |
| `page number style trad chin num2` |  |
| `page number style simp chin num1` |  |
| `page number style simp chin num2` |  |
| `page number style hebrew letter1` |  |
| `page number style arabic letter1` |  |
| `page number style hebrew letter2` |  |
| `page number style arabic letter2` |  |
| `page number style hindi letter1` |  |
| `page number style hindi letter2` |  |
| `page number style hindi arabic` |  |
| `page number style hindi cardinal text` |  |
| `page number style thai letter` |  |
| `page number style thai arabic` |  |
| `page number style thai cardinal text` |  |
| `page number style viet cardinal text` |  |
| `page number style number in dash` |  |

### `WdStatistic`

| Value | Description |
|-------|-------------|
| `statistic words` |  |
| `statistic lines` |  |
| `statistic pages` |  |
| `statistic characters` |  |
| `statistic paragraphs` |  |
| `statistic characters with spaces` |  |
| `statistic east asian characters` |  |

### `WdBuiltInProperty`

| Value | Description |
|-------|-------------|
| `property title` |  |
| `property subject` |  |
| `property author` |  |
| `property keywords` |  |
| `property comments` |  |
| `property template` |  |
| `property last author` |  |
| `property revision` |  |
| `property app name` |  |
| `property time last printed` |  |
| `property time created` |  |
| `property time last saved` |  |
| `property vbatotal edit` |  |
| `property pages` |  |
| `property words` |  |
| `property characters` |  |
| `property security` |  |
| `property category` |  |
| `property format` |  |
| `property manager` |  |
| `property company` |  |
| `property bytes` |  |
| `property lines` |  |
| `property paras` |  |
| `property slides` |  |
| `property notes` |  |
| `property hidden slides` |  |
| `property mmclips` |  |
| `property hyperlink base` |  |
| `property chars wspaces` |  |

### `WdLineSpacing`

| Value | Description |
|-------|-------------|
| `line space single` |  |
| `line space1 pt5` |  |
| `line space double` |  |
| `line space at least` |  |
| `line space exactly` |  |
| `line space multiple` |  |

### `WdNumberType`

| Value | Description |
|-------|-------------|
| `number paragraph` |  |
| `number listnum` |  |
| `number all numbers` |  |

### `WdListType`

| Value | Description |
|-------|-------------|
| `list no numbering` |  |
| `list listnum only` |  |
| `list bullet` |  |
| `list simple numbering` |  |
| `list outline numbering` |  |
| `list mixed numbering` |  |
| `list picture bullet` |  |

### `WdStoryType`

| Value | Description |
|-------|-------------|
| `main text story` |  |
| `footnotes story` |  |
| `endnotes story` |  |
| `comments story` |  |
| `text frame story` |  |
| `even pages header story` |  |
| `primary header story` |  |
| `even pages footer story` |  |
| `primary footer story` |  |
| `first page header story` |  |
| `first page footer story` |  |
| `footnote separator story` |  |
| `footnote continuation separator  story` |  |
| `footnote continuation notice story` |  |
| `endnote separator story` |  |
| `endnote continuation separator  story` |  |
| `endnote continuation notice story` |  |

### `WdSaveFormat`

| Value | Description |
|-------|-------------|
| `format document97` |  |
| `format template97` |  |
| `format text` |  |
| `format text line breaks` |  |
| `format dostext` |  |
| `format dostext line breaks` |  |
| `format rtf` |  |
| `format Unicode text` |  |
| `format HTML` |  |
| `format web archive` |  |
| `format filtered HTML` |  |
| `format xml` |  |
| `format document` |  |
| `format documentME` |  |
| `format template` |  |
| `format templateME` |  |
| `format document default` |  |
| `format PDF` |  |
| `format` |  |
| `format flat document` |  |
| `format flat documentME` |  |
| `format flat template` |  |
| `format flat templateME` |  |
| `format ODT` |  |
| `format strict open xmldocument` |  |

### `WdOpenFormat`

| Value | Description |
|-------|-------------|
| `open format auto` |  |
| `open format document` |  |
| `open format document97` |  |
| `open format template` |  |
| `open format template97` |  |
| `open format rtf` |  |
| `open format text` |  |
| `open format unicode text` |  |
| `open format encoded text` |  |
| `open format all word` |  |
| `open format web pages` |  |
| `open format xml` |  |
| `open format xmldocument` |  |
| `open format xmldocument macro enabled` |  |
| `open format xmltemplate` |  |
| `open format xmltemplate macro enabled` |  |
| `open format all word templates` |  |
| `open format xmldocument serialized` |  |
| `open format xmldocument macro enabled serialized` |  |
| `open format xmltemplate serialized` |  |
| `open format xmltemplate macro enabled serialized` |  |
| `open format open document text` |  |

### `WdHeaderFooterIndex`

| Value | Description |
|-------|-------------|
| `header footer primary` |  |
| `header footer first page` |  |
| `header footer even pages` |  |

### `WdTocFormat`

| Value | Description |
|-------|-------------|
| `toctemplate` |  |
| `tocclassic` |  |
| `tocdistinctive` |  |
| `tocfancy` |  |
| `tocmodern` |  |
| `tocformal` |  |
| `tocsimple` |  |

### `WdTofFormat`

| Value | Description |
|-------|-------------|
| `toftemplate` |  |
| `tofclassic` |  |
| `tofdistinctive` |  |
| `tofcentered` |  |
| `tofformal` |  |
| `tofsimple` |  |

### `WdToaFormat`

| Value | Description |
|-------|-------------|
| `toatemplate` |  |
| `toaclassic` |  |
| `toadistinctive` |  |
| `toaformal` |  |
| `toasimple` |  |

### `WdLineStyle`

| Value | Description |
|-------|-------------|
| `line style none` |  |
| `line style single` |  |
| `line style dot` |  |
| `line style dash small gap` |  |
| `line style dash large gap` |  |
| `line style dash dot` |  |
| `line style dash dot dot` |  |
| `line style double` |  |
| `line style triple` |  |
| `line style thin thick small gap` |  |
| `line style thick thin small gap` |  |
| `line style thin thick thin small gap` |  |
| `line style thin thick med gap` |  |
| `line style thick thin med gap` |  |
| `line style thin thick thin med gap` |  |
| `line style thin thick large gap` |  |
| `line style thick thin large gap` |  |
| `line style thin thick thin large gap` |  |
| `line style single wavy` |  |
| `line style double wavy` |  |
| `line style dash dot stroked` |  |
| `line style emboss_3D` |  |
| `line style engrave_3D` |  |
| `line style outset` |  |
| `line style inset` |  |

### `WdLineWidth`

| Value | Description |
|-------|-------------|
| `line width25 point` |  |
| `line width50 point` |  |
| `line width75 point` |  |
| `line width100 point` |  |
| `line width150 point` |  |
| `line width225 point` |  |
| `line width300 point` |  |
| `line width450 point` |  |
| `line width600 point` |  |

### `WdBreakType`

| Value | Description |
|-------|-------------|
| `section break next page` |  |
| `section break continuous` |  |
| `section break even page` |  |
| `section break odd page` |  |
| `line break` |  |
| `page break` |  |
| `column break` |  |
| `line break clear left` |  |
| `line break clear right` |  |
| `text wrapping break` |  |

### `WdPageType`

| Value | Description |
|-------|-------------|
| `continued master` |  |

### `WdTabLeader`

| Value | Description |
|-------|-------------|
| `tab leader spaces` |  |
| `tab leader dots` |  |
| `tab leader dashes` |  |
| `tab leader lines` |  |
| `tab leader heavy` |  |
| `tab leader middle dot` |  |

### `WdMeasurementUnits`

| Value | Description |
|-------|-------------|
| `inches` |  |
| `centimeters` |  |
| `millimeters` |  |
| `points` |  |
| `picas` |  |

### `WdDropPosition`

| Value | Description |
|-------|-------------|
| `drop none` |  |
| `drop normal` |  |
| `drop margin` |  |

### `WdNumberingRule`

| Value | Description |
|-------|-------------|
| `restart continuous` |  |
| `restart section` |  |
| `restart page` |  |

### `WdFootnoteLocation`

| Value | Description |
|-------|-------------|
| `bottom of page` |  |
| `beneath text` |  |

### `WdEndnoteLocation`

| Value | Description |
|-------|-------------|
| `end_of_section` |  |
| `end_of_document` |  |

### `WdSortSeparator`

| Value | Description |
|-------|-------------|
| `sort separate by tabs` |  |
| `sort separate by commas` |  |
| `sort separate by default table separator` |  |

### `WdTableFieldSeparator`

| Value | Description |
|-------|-------------|
| `separate by paragraphs` |  |
| `separate by tabs` |  |
| `separate by commas` |  |
| `separate by default list separator` |  |

### `WdSortFieldType`

| Value | Description |
|-------|-------------|
| `sort field alphanumeric` |  |
| `sort field numeric` |  |
| `sort field date` |  |
| `sort field syllable` |  |
| `sort field japan jis` |  |
| `sort field stroke` |  |
| `sort field korea ks` |  |

### `WdSortOrder`

| Value | Description |
|-------|-------------|
| `sort order ascending` |  |
| `sort order descending` |  |

### `WdTableFormat`

| Value | Description |
|-------|-------------|
| `table format none` |  |
| `table format simple1` |  |
| `table format simple2` |  |
| `table format simple3` |  |
| `table format classic1` |  |
| `table format classic2` |  |
| `table format classic3` |  |
| `table format classic4` |  |
| `table format colorful1` |  |
| `table format colorful2` |  |
| `table format colorful3` |  |
| `table format columns1` |  |
| `table format columns2` |  |
| `table format columns3` |  |
| `table format columns4` |  |
| `table format columns5` |  |
| `table format grid1` |  |
| `table format grid2` |  |
| `table format grid3` |  |
| `table format grid4` |  |
| `table format grid5` |  |
| `table format grid6` |  |
| `table format grid7` |  |
| `table format grid8` |  |
| `table format list1` |  |
| `table format list2` |  |
| `table format list3` |  |
| `table format list4` |  |
| `table format list5` |  |
| `table format list6` |  |
| `table format list7` |  |
| `table format list8` |  |
| `table format3D effects1` |  |
| `table format3D effects2` |  |
| `table format3D effects3` |  |
| `table format contemporary` |  |
| `table format elegant` |  |
| `table format professional` |  |
| `table format subtle1` |  |
| `table format subtle2` |  |
| `table format web1` |  |
| `table format web2` |  |
| `table format web3` |  |

### `WdTableFormatApply`

| Value | Description |
|-------|-------------|
| `table format apply borders` |  |
| `table format apply shading` |  |
| `table format apply font` |  |
| `table format apply color` |  |
| `table format apply auto fit` |  |
| `table format apply heading rows` |  |
| `table format apply last row` |  |
| `table format apply first column` |  |
| `table format apply last column` |  |

### `WdLanguageID`

| Value | Description |
|-------|-------------|
| `language none` |  |
| `language no proofing` |  |
| `afrikaans` |  |
| `albanian` |  |
| `amharic` |  |
| `arabic algeria` |  |
| `arabic bahrain` |  |
| `arabic egypt` |  |
| `arabic iraq` |  |
| `arabic jordan` |  |
| `arabic kuwait` |  |
| `arabic lebanon` |  |
| `arabic libya` |  |
| `arabic morocco` |  |
| `arabic oman` |  |
| `arabic qatar` |  |
| `arabic` |  |
| `arabic syria` |  |
| `arabic tunisia` |  |
| `arabic uae` |  |
| `arabic yemen` |  |
| `armenian` |  |
| `assamese` |  |
| `azeri cyrillic` |  |
| `azeri latin` |  |
| `basque` |  |
| `byelorussian` |  |
| `bengali` |  |
| `bulgarian` |  |
| `burmese` |  |
| `catalan` |  |
| `cherokee` |  |
| `chinese hong kong` |  |
| `chinese macao` |  |
| `simplified chinese` |  |
| `chinese singapore` |  |
| `traditional chinese` |  |
| `croatian` |  |
| `czech` |  |
| `danish` |  |
| `divehi` |  |
| `belgian dutch` |  |
| `dutch` |  |
| `edo` |  |
| `english aus` |  |
| `english belize` |  |
| `english canadian` |  |
| `english caribbean` |  |
| `english ireland` |  |
| `english jamaica` |  |
| `english new zealand` |  |
| `english philippines` |  |
| `english south africa` |  |
| `english trinidad tobago` |  |
| `english uk` |  |
| `english us` |  |
| `english zimbabwe` |  |
| `english indonesia` |  |
| `estonian` |  |
| `faeroese` |  |
| `persian` |  |
| `filipino` |  |
| `finnish` |  |
| `fulfulde` |  |
| `belgian french` |  |
| `french cameroon` |  |
| `french canadian` |  |
| `french coted ivoire` |  |
| `french` |  |
| `french luxembourg` |  |
| `french mali` |  |
| `french monaco` |  |
| `french reunion` |  |
| `french senegal` |  |
| `french morocco` |  |
| `french haiti` |  |
| `swiss french` |  |
| `french west indies` |  |
| `french congo` |  |
| `frisian netherlands` |  |
| `gaelic ireland` |  |
| `gaelic scotland` |  |
| `galician` |  |
| `georgian` |  |
| `german austria` |  |
| `german` |  |
| `german liechtenstein` |  |
| `german luxembourg` |  |
| `swiss german` |  |
| `greek` |  |
| `guarani` |  |
| `gujarati` |  |
| `hausa` |  |
| `hawaiian` |  |
| `hebrew` |  |
| `hindi` |  |
| `hungarian` |  |
| `ibibio` |  |
| `icelandic` |  |
| `igbo` |  |
| `indonesian` |  |
| `inuktitut` |  |
| `italian` |  |
| `swiss italian` |  |
| `japanese` |  |
| `kannada` |  |
| `kanuri` |  |
| `kashmiri` |  |
| `kazakh` |  |
| `khmer` |  |
| `kirghiz` |  |
| `konkani` |  |
| `korean` |  |
| `kyrgyz` |  |
| `lao` |  |
| `latin` |  |
| `latvian` |  |
| `lithuanian` |  |
| `macedonian` |  |
| `malaysian` |  |
| `malay brunei darussalam` |  |
| `malayalam` |  |
| `maltese` |  |
| `manipuri` |  |
| `marathi` |  |
| `mongolian` |  |
| `nepali` |  |
| `norwegian bokmol` |  |
| `norwegian nynorsk` |  |
| `oriya` |  |
| `oromo` |  |
| `pashto` |  |
| `polish` |  |
| `portuguese brazil` |  |
| `portuguese` |  |
| `punjabi` |  |
| `rhaeto romanic` |  |
| `romanian moldova` |  |
| `romanian` |  |
| `russian moldova` |  |
| `russian` |  |
| `sami lappish` |  |
| `sanskrit` |  |
| `serbian cyrillic` |  |
| `serbian latin` |  |
| `sinhalese` |  |
| `sindhi` |  |
| `sindhi pakistan` |  |
| `slovak` |  |
| `slovenian` |  |
| `somali` |  |
| `sorbian` |  |
| `spanish argentina` |  |
| `spanish bolivia` |  |
| `spanish chile` |  |
| `spanish colombia` |  |
| `spanish costa rica` |  |
| `spanish dominican republic` |  |
| `spanish ecuador` |  |
| `spanish el salvador` |  |
| `spanish guatemala` |  |
| `spanish honduras` |  |
| `mexican spanish` |  |
| `spanish nicaragua` |  |
| `spanish panama` |  |
| `spanish paraguay` |  |
| `spanish peru` |  |
| `spanish puerto rico` |  |
| `spanish modern sort` |  |
| `spanish` |  |
| `spanish uruguay` |  |
| `spanish venezuela` |  |
| `sesotho` |  |
| `sutu` |  |
| `swahili` |  |
| `swedish finland` |  |
| `swedish` |  |
| `syriac` |  |
| `tajik` |  |
| `tamazight` |  |
| `tamazight latin` |  |
| `tamil` |  |
| `tatar` |  |
| `telugu` |  |
| `thai` |  |
| `tibetan` |  |
| `tigrigna ethiopic` |  |
| `tigrigna eritrea` |  |
| `tsonga` |  |
| `tswana` |  |
| `turkish` |  |
| `turkmen` |  |
| `ukrainian` |  |
| `urdu` |  |
| `uzbek cyrillic` |  |
| `uzbek latin` |  |
| `venda` |  |
| `vietnamese` |  |
| `welsh` |  |
| `xhosa` |  |
| `yi` |  |
| `yiddish` |  |
| `yoruba` |  |
| `zulu` |  |

### `WdFieldType`

| Value | Description |
|-------|-------------|
| `field empty` |  |
| `field ref` |  |
| `field index entry` |  |
| `field footnote ref` |  |
| `field set` |  |
| `field if` |  |
| `field index` |  |
| `field toc entry` |  |
| `field style ref` |  |
| `field ref doc` |  |
| `field sequence` |  |
| `field toc` |  |
| `field info` |  |
| `field title` |  |
| `field subject` |  |
| `field author` |  |
| `field key word` |  |
| `field comments` |  |
| `field last saved by` |  |
| `field create date` |  |
| `field save date` |  |
| `field print date` |  |
| `field revision num` |  |
| `field edit time` |  |
| `field num pages` |  |
| `field num words` |  |
| `field num chars` |  |
| `field file name` |  |
| `field template` |  |
| `field date` |  |
| `field time` |  |
| `field page` |  |
| `field expression` |  |
| `field quote` |  |
| `field include` |  |
| `field page ref` |  |
| `field ask` |  |
| `field fill in` |  |
| `field data` |  |
| `field next` |  |
| `field next if` |  |
| `field skip if` |  |
| `field merge rec` |  |
| `field dde` |  |
| `field ddeauto` |  |
| `field glossary` |  |
| `field print` |  |
| `field formula` |  |
| `field go to button` |  |
| `field macro button` |  |
| `field auto num outline` |  |
| `field auto num legal` |  |
| `field auto num` |  |
| `field import` |  |
| `field link` |  |
| `field symbol` |  |
| `field embed` |  |
| `field merge field` |  |
| `field user name` |  |
| `field user initials` |  |
| `field user address` |  |
| `field bar code` |  |
| `field doc variable` |  |
| `field section` |  |
| `field section pages` |  |
| `field include picture` |  |
| `field include text` |  |
| `field file size` |  |
| `field form text input` |  |
| `field form check box` |  |
| `field note ref` |  |
| `field toa` |  |
| `field toaentry` |  |
| `field merge seq` |  |
| `field private` |  |
| `field database` |  |
| `field auto text` |  |
| `field compare` |  |
| `field addin` |  |
| `field subscriber` |  |
| `field form drop down` |  |
| `field advance` |  |
| `field doc property` |  |
| `field ocx` |  |
| `field hyperlink` |  |
| `field auto text list` |  |
| `field listnum` |  |
| `field htmlactive x` |  |
| `field bidi outline` |  |
| `field address block` |  |
| `field greeting line` |  |
| `field shape` |  |
| `field citation` |  |
| `field bibliography` |  |
| `field merge barcode` |  |
| `field display barcode` |  |

### `WdBuiltinStyle`

| Value | Description |
|-------|-------------|
| `style normal` |  |
| `style heading1` |  |
| `style heading2` |  |
| `style heading3` |  |
| `style heading4` |  |
| `style heading5` |  |
| `style heading6` |  |
| `style heading7` |  |
| `style heading8` |  |
| `style heading9` |  |
| `style index1` |  |
| `style index2` |  |
| `style index3` |  |
| `style index4` |  |
| `style index5` |  |
| `style index6` |  |
| `style index7` |  |
| `style index8` |  |
| `style index9` |  |
| `style toc1` |  |
| `style toc2` |  |
| `style toc3` |  |
| `style toc4` |  |
| `style toc5` |  |
| `style toc6` |  |
| `style toc7` |  |
| `style toc8` |  |
| `style toc9` |  |
| `style normal indent` |  |
| `style footnote text` |  |
| `style comment text` |  |
| `style header` |  |
| `style footer` |  |
| `style index heading` |  |
| `style caption` |  |
| `style table of figures` |  |
| `style envelope address` |  |
| `style envelope return` |  |
| `style footnote reference` |  |
| `style comment reference` |  |
| `style line number` |  |
| `style page number` |  |
| `style endnote reference` |  |
| `style endnote text` |  |
| `style table of authorities` |  |
| `style macro text` |  |
| `style toa heading` |  |
| `style list` |  |
| `style list bullet` |  |
| `style list number` |  |
| `style list2` |  |
| `style list3` |  |
| `style list4` |  |
| `style list5` |  |
| `style list bullet2` |  |
| `style list bullet3` |  |
| `style list bullet4` |  |
| `style list bullet5` |  |
| `style list number2` |  |
| `style list number3` |  |
| `style list number4` |  |
| `style list number5` |  |
| `style title` |  |
| `style closing` |  |
| `style signature` |  |
| `style default paragraph font` |  |
| `style body text` |  |
| `style body text indent` |  |
| `style list continue` |  |
| `style list continue2` |  |
| `style list continue3` |  |
| `style list continue4` |  |
| `style list continue5` |  |
| `style message header` |  |
| `style subtitle` |  |
| `style salutation` |  |
| `style date` |  |
| `style body text first indent` |  |
| `style body text first indent2` |  |
| `style note heading` |  |
| `style body text2` |  |
| `style body text3` |  |
| `style body text indent2` |  |
| `style body text indent3` |  |
| `style block quotation` |  |
| `style hyperlink` |  |
| `style hyperlink followed` |  |
| `style strong` |  |
| `style emphasis` |  |
| `style nav pane` |  |
| `style plain text` |  |
| `style html normal` |  |
| `style html acronym` |  |
| `style html address` |  |
| `style html cite` |  |
| `style html code` |  |
| `style html define` |  |
| `style html keyboard` |  |
| `style html pre` |  |
| `style html samp` |  |
| `style html tt` |  |
| `style html var` |  |
| `style normal table` |  |
| `style normal object` |  |
| `style table light shading` |  |
| `style table light list` |  |
| `style table light grid` |  |
| `style table medium shading1` |  |
| `style table medium shading2` |  |
| `style table medium list1` |  |
| `style table medium list2` |  |
| `style table medium grid1` |  |
| `style table medium grid2` |  |
| `style table medium grid3` |  |
| `style table dark list` |  |
| `style table colorful shading` |  |
| `style table colorful list` |  |
| `style table colorful grid` |  |
| `style table light shading accent1` |  |
| `style table light list accent1` |  |
| `style table light grid accent1` |  |
| `style table medium shading1 accent1` |  |
| `style table medium shading2 accent1` |  |
| `style table medium list1 accent1` |  |
| `style list paragraph` |  |
| `style quote` |  |
| `style intense quote` |  |
| `style subtle emphasis` |  |
| `style intense emphasis` |  |
| `style subtle reference` |  |
| `style intense reference` |  |
| `style book title` |  |
| `style bibliography` |  |
| `style toc heading` |  |

### `WdWordDialogTab`

| Value | Description |
|-------|-------------|
| `dialog file document layout tab margins` |  |
| `dialog file document layout tab layout` |  |
| `dialog file page setup tab margins` |  |
| `dialog file page setup tab paper size` |  |
| `dialog file page setup tab paper source` |  |
| `dialog file page setup tab layout` |  |
| `dialog file page setup tab chars lines` |  |
| `dialog insert symbol tab symbols` |  |
| `dialog insert symbol tab special characters` |  |
| `dialog note options tab all footnotes` |  |
| `dialog note options tab all endnotes` |  |
| `dialog insert index and tables tab index` |  |
| `dialog insert index and tables tab table of contents` |  |
| `dialog insert index and tables tab table of figures` |  |
| `dialog insert index and tables tab table of authorities` |  |
| `dialog organizer tab styles` |  |
| `dialog organizer tab auto text` |  |
| `dialog organizer tab command bars` |  |
| `dialog organizer tab macros` |  |
| `dialog format font tab font` |  |
| `dialog format font tab character spacing` |  |
| `dialog format borders and shading tab borders` |  |
| `dialog format borders and shading tab page border` |  |
| `dialog format borders and shading tab shading` |  |
| `dialog tools envelopes and labels tab envelopes` |  |
| `dialog tools envelopes and labels tab labels` |  |
| `dialog format paragraph tab indents and spacing` |  |
| `dialog format paragraph tab text flow` |  |
| `dialog format paragraph tab teisai` |  |
| `dialog format drawing object tab colors and lines` |  |
| `dialog format drawing object tab size` |  |
| `dialog format drawing object tab position` |  |
| `dialog format drawing object tab wrapping` |  |
| `dialog format drawing object tab picture` |  |
| `dialog format drawing object tab textbox` |  |
| `dialog format drawing object tab hr` |  |
| `dialog tools autocorrect exceptions tab first letter` |  |
| `dialog tools autocorrect exceptions tab initial caps` |  |
| `dialog tools autocorrect exceptions tab hangul and alphabet` |  |
| `dialog tools autocorrect exceptions tab iac` |  |
| `dialog format bullets and numbering tab bulleted` |  |
| `dialog format bullets and numbering tab numbered` |  |
| `dialog format bullets and numbering tab outline numbered` |  |
| `dialog letter wizard tab letter format` |  |
| `dialog letter wizard tab recipient info` |  |
| `dialog letter wizard tab other elements` |  |
| `dialog letter wizard tab sender info` |  |
| `dialog tools auto manager tab autocorrect` |  |
| `dialog tools auto manager tab math autocorrect` |  |
| `dialog tools auto manager tab auto format as you type` |  |
| `dialog tools auto manager tab auto text` |  |
| `dialog tools auto manager tab auto format` |  |
| `dialog web options general` |  |
| `dialog web options files` |  |
| `dialog web options pictures` |  |
| `dialog web options encoding` |  |
| `dialog web options fonts` |  |
| `dialog format drawing object tab alt text` |  |

### `WdWordDialog`

| Value | Description |
|-------|-------------|
| `dialog help about` |  |
| `dialog help word perfect help` |  |
| `dialog document statistics` |  |
| `dialog file new` |  |
| `dialog file open` |  |
| `dialog data merge open data source` |  |
| `dialog data merge open header source` |  |
| `dialog file save as` |  |
| `dialog file summary info` |  |
| `dialog tools templates` |  |
| `dialog file print` |  |
| `dialog file print setup` |  |
| `dialog file find` |  |
| `dialog format address fonts` |  |
| `dialog edit paste special` |  |
| `dialog edit find` |  |
| `dialog edit replace` |  |
| `dialog edit style` |  |
| `dialog edit links` |  |
| `dialog edit object` |  |
| `dialog table to text` |  |
| `dialog text to table` |  |
| `dialog table insert table` |  |
| `dialog table insert cells` |  |
| `dialog table insert row` |  |
| `dialog table delete cells` |  |
| `dialog table split cells` |  |
| `dialog table row height` |  |
| `dialog table column width` |  |
| `dialog tools customize` |  |
| `dialog insert break` |  |
| `dialog insert symbol` |  |
| `dialog insert picture` |  |
| `dialog insert file` |  |
| `dialog insert date time` |  |
| `dialog insert field` |  |
| `dialog insert merge field` |  |
| `dialog insert bookmark` |  |
| `dialog mark index entry` |  |
| `dialog insert index` |  |
| `dialog insert table of contents` |  |
| `dialog insert object` |  |
| `dialog tools create envelope` |  |
| `dialog format font` |  |
| `dialog format paragraph` |  |
| `dialog format section layout` |  |
| `dialog format columns` |  |
| `dialog file document layout` |  |
| `dialog file page setup` |  |
| `dialog format tabs` |  |
| `dialog format style` |  |
| `dialog format define style font` |  |
| `dialog format define style para` |  |
| `dialog format define style tabs` |  |
| `dialog format define style frame` |  |
| `dialog format define style borders` |  |
| `dialog format define style lang` |  |
| `dialog format picture` |  |
| `dialog tools language` |  |
| `dialog format borders and shading` |  |
| `dialog format frame` |  |
| `dialog tools thesaurus` |  |
| `dialog tools hyphenation` |  |
| `dialog tools bullets numbers` |  |
| `dialog tools highlight changes` |  |
| `dialog tools revisions` |  |
| `dialog tools compare documents` |  |
| `dialog table sort` |  |
| `dialog tools options general` |  |
| `dialog tools options view` |  |
| `dialog tools advanced settings` |  |
| `dialog tools options print` |  |
| `dialog tools options save` |  |
| `dialog tools options spelling and grammar` |  |
| `dialog tools options user info` |  |
| `dialog tools macro record` |  |
| `dialog tools macro` |  |
| `dialog window activate` |  |
| `dialog format ret addr fonts` |  |
| `dialog organizer` |  |
| `dialog tools options edit` |  |
| `dialog tools options file locations` |  |
| `dialog tools word count` |  |
| `dialog control run` |  |
| `dialog insert page numbers` |  |
| `dialog format page number` |  |
| `dialog copy file` |  |
| `dialog format change case` |  |
| `dialog update` |  |
| `dialog insert database` |  |
| `dialog table formula` |  |
| `dialog form field options` |  |
| `dialog insert caption` |  |
| `dialog insert caption numbering` |  |
| `dialog insert auto caption` |  |
| `dialog form field help` |  |
| `dialog insert cross reference` |  |
| `dialog insert footnote` |  |
| `dialog note options` |  |
| `dialog tools auto correct` |  |
| `dialog tools options track changes` |  |
| `dialog convert object` |  |
| `dialog insert add caption` |  |
| `dialog connect` |  |
| `dialog tools customize keyboard` |  |
| `dialog tools customize menus` |  |
| `dialog tools merge documents` |  |
| `dialog mark table of contents entry` |  |
| `dialog file print one copy` |  |
| `dialog edit frame` |  |
| `dialog mark citation` |  |
| `dialog table of contents options` |  |
| `dialog insert table of authorities` |  |
| `dialog insert table of figures` |  |
| `dialog insert index and tables` |  |
| `dialog insert form field` |  |
| `dialog format drop cap` |  |
| `dialog tools create labels` |  |
| `dialog tools protect document` |  |
| `dialog format style gallery` |  |
| `dialog tools accept reject changes` |  |
| `dialog help word perfect help options` |  |
| `dialog tools unprotect document` |  |
| `dialog tools options compatibility` |  |
| `dialog table of captions options` |  |
| `dialog table auto format` |  |
| `dialog mail merge find record` |  |
| `dialog review afmt revisions` |  |
| `dialog view zoom` |  |
| `dialog tools protect section` |  |
| `dialog font substitution` |  |
| `dialog insert subdocument` |  |
| `dialog new toolbar` |  |
| `dialog tools envelopes and labels` |  |
| `dialog format callout` |  |
| `dialog table format cell` |  |
| `dialog tools customize menu bar` |  |
| `dialog file routing slip` |  |
| `dialog edit category` |  |
| `dialog tools manage fields` |  |
| `dialog draw snap to grid` |  |
| `dialog draw align` |  |
| `dialog mail merge create data source` |  |
| `dialog mail merge create header source` |  |
| `dialog mail merge` |  |
| `dialog mail merge check` |  |
| `dialog mail merge helper` |  |
| `dialog mail merge query options` |  |
| `dialog file mac page setup` |  |
| `dialog list commands` |  |
| `dialog edit create publisher` |  |
| `dialog edit subscribe to` |  |
| `dialog edit publish options` |  |
| `dialog edit subscribe options` |  |
| `dialog file mac custom page setup` |  |
| `dialog tools options typography` |  |
| `dialog tools auto correct exceptions` |  |
| `dialog tools options auto format as you type` |  |
| `dialog mail merge use address book` |  |
| `dialog tools hangul hanja conversion` |  |
| `dialog tools options fuzzy` |  |
| `dialog edit go to old` |  |
| `dialog insert number` |  |
| `dialog letter wizard` |  |
| `dialog format bullets and numbering` |  |
| `dialog tools spelling and grammar` |  |
| `dialog tools create directory` |  |
| `dialog table wrapping` |  |
| `dialog format theme` |  |
| `dialog table properties` |  |
| `dialog email options` |  |
| `dialog create auto text` |  |
| `dialog tools auto summarize` |  |
| `dialog tools grammar settings` |  |
| `dialog edit go to` |  |
| `dialog web options` |  |
| `dialog insert hyperlink` |  |
| `dialog tools auto manager` |  |
| `dialog file versions` |  |
| `dialog tools options auto format` |  |
| `dialog format drawing object` |  |
| `dialog tools options` |  |
| `dialog fit text` |  |
| `dialog edit auto text` |  |
| `dialog phonetic guide` |  |
| `dialog tools dictionary` |  |
| `dialog file save version` |  |
| `dialog tools options bidi` |  |
| `dialog frame set properties` |  |
| `dialog table table options` |  |
| `dialog table cell options` |  |
| `dialog set default` |  |
| `dialog translator` |  |
| `dialog horizontal in vertical` |  |
| `dialog two lines in one` |  |
| `dialog format enclose characters` |  |
| `dialog consistency checker` |  |
| `dialog tools options smart tag` |  |
| `dialog format styles custom` |  |
| `dialog links` |  |
| `dialog insert web component` |  |
| `dialog tools options edit copy paste` |  |
| `dialog tools options security` |  |
| `dialog search` |  |
| `dialog show repairs` |  |
| `dialog mail merge insert ask` |  |
| `dialog mail merge insert fill in` |  |
| `dialog mail merge insert if` |  |
| `dialog mail merge insert next if` |  |
| `dialog mail merge insert set` |  |
| `dialog mail merge insert skip if` |  |
| `dialog mail merge field mapping` |  |
| `dialog mail merge insert address block` |  |
| `dialog mail merge insert greeting line` |  |
| `dialog mail merge insert fields` |  |
| `dialog mail merge recipients` |  |
| `dialog mail merge find recipient` |  |
| `dialog mail merge set document type` |  |
| `dialog label options` |  |
| `dialog element attributes` |  |
| `dialog schema library` |  |
| `dialog permission` |  |
| `dialog my permission` |  |
| `dialog options` |  |
| `dialog formatting restrictions` |  |
| `dialog source manager` |  |
| `dialog create source` |  |
| `dialog document inspector` |  |
| `dialog style management` |  |
| `dialog insert source` |  |
| `dialog math recognized functions` |  |
| `dialog insert placeholder` |  |
| `dialog building block organizer` |  |
| `dialog content control properties` |  |
| `dialog compatibility checker` |  |
| `dialog export as fixed format` |  |
| `dialog file new` |  |

### `WdFieldKind`

| Value | Description |
|-------|-------------|
| `field kind none` |  |
| `field kind hot` |  |
| `field kind warm` |  |
| `field kind cold` |  |

### `WdTextFormFieldType`

| Value | Description |
|-------|-------------|
| `regular text` |  |
| `number text` |  |
| `date text` |  |
| `current date text` |  |
| `current time text` |  |
| `calculation text` |  |

### `WdChevronConvertRule`

| Value | Description |
|-------|-------------|
| `never convert` |  |
| `always convert` |  |
| `ask to not convert` |  |
| `ask to convert` |  |

### `WdMailMergeMainDocType`

| Value | Description |
|-------|-------------|
| `not a merge document` |  |
| `document type form letters` |  |
| `document type mailing labels` |  |
| `document type envelopes` |  |
| `document type catalog` |  |
| `email` |  |
| `fax` |  |
| `directory` |  |

### `WdMailMergeState`

| Value | Description |
|-------|-------------|
| `normal document` |  |
| `main document only` |  |
| `main and data source` |  |
| `main and header` |  |
| `main and source and header` |  |
| `data source` |  |

### `WdMailMergeDestination`

| Value | Description |
|-------|-------------|
| `send to new document` |  |
| `send to printer` |  |
| `send to email` |  |
| `send to fax` |  |

### `WdMailMergeActiveRecord`

| Value | Description |
|-------|-------------|
| `no active record` |  |
| `next data record` |  |
| `previous data record` |  |
| `first data record` |  |
| `last data record` |  |
| `first data source record` |  |
| `last data source record` |  |
| `next data source record` |  |
| `previous data source record` |  |

### `WdMailMergeDefaultRecord`

| Value | Description |
|-------|-------------|
| `default first record` |  |
| `default last record` |  |

### `WdMailMergeDataSource`

| Value | Description |
|-------|-------------|
| `no merge info` |  |
| `merge info from word` |  |
| `merge info from access dde` |  |
| `merge info from excel dde` |  |
| `merge info from msquery dde` |  |
| `merge info from odbc` |  |

### `WdMailMergeComparison`

| Value | Description |
|-------|-------------|
| `merge if equal` |  |
| `merge if not equal` |  |
| `merge if less than` |  |
| `merge if greater than` |  |
| `merge if less than or equal` |  |
| `merge if greater than or equal` |  |
| `merge if is blank` |  |
| `merge if is not blank` |  |

### `WdBookmarkSortBy`

| Value | Description |
|-------|-------------|
| `sort by name` |  |
| `sort by location` |  |

### `WdWindowState`

| Value | Description |
|-------|-------------|
| `window state normal` |  |
| `window state maximize` |  |
| `window state minimize` |  |

### `WdPictureLinkType`

| Value | Description |
|-------|-------------|
| `link none` |  |
| `link data in doc` |  |
| `link data on disk` |  |

### `WdLinkType`

| Value | Description |
|-------|-------------|
| `link type ole` |  |
| `link type picture` |  |
| `link type text` |  |
| `link type reference` |  |
| `link type include` |  |
| `link type import` |  |
| `link type dde` |  |
| `link type ddeauto` |  |
| `link type chart` |  |

### `WdWindowType`

| Value | Description |
|-------|-------------|
| `window document` |  |
| `window template` |  |

### `WdViewType`

| Value | Description |
|-------|-------------|
| `normal view` |  |
| `draft view` |  |
| `outline view` |  |
| `print view` |  |
| `page view` |  |
| `print preview view` |  |
| `master view` |  |
| `web view` |  |
| `online view` |  |
| `reading view` |  |
| `conflict view` |  |

### `WdSeekView`

| Value | Description |
|-------|-------------|
| `seek main document` |  |
| `seek primary header` |  |
| `seek first page header` |  |
| `seek even pages header` |  |
| `seek primary footer` |  |
| `seek first page footer` |  |
| `seek even pages footer` |  |
| `seek footnotes` |  |
| `seek endnotes` |  |
| `seek current page header` |  |
| `seek current page footer` |  |

### `WdSpecialPane`

| Value | Description |
|-------|-------------|
| `pane none` |  |
| `pane primary header` |  |
| `pane first page header` |  |
| `pane even pages header` |  |
| `pane primary footer` |  |
| `pane first page footer` |  |
| `pane even pages footer` |  |
| `pane footnotes` |  |
| `pane endnotes` |  |
| `pane footnote continuation notice` |  |
| `pane footnote continuation separator` |  |
| `pane footnote separator` |  |
| `pane endnote continuation notice` |  |
| `pane endnote continuation separator` |  |
| `pane endnote separator` |  |
| `pane comments` |  |
| `pane current page header` |  |
| `pane current page footer` |  |
| `pane revisions` |  |
| `pane revisions horiz` |  |
| `pane revisions vert` |  |

### `WdPageFit`

| Value | Description |
|-------|-------------|
| `page fit none` |  |
| `page fit full page` |  |
| `page fit best fit` |  |
| `page fit text fit` |  |

### `WdBrowseTarget`

| Value | Description |
|-------|-------------|
| `browse page` |  |
| `browse section` |  |
| `browse comment` |  |
| `browse footnote` |  |
| `browse endnote` |  |
| `browse field` |  |
| `browse table` |  |
| `browse graphic` |  |
| `browse heading` |  |
| `browse edit` |  |
| `browse find` |  |
| `browse go to` |  |

### `WdPaperTray`

| Value | Description |
|-------|-------------|
| `printer default bin` |  |
| `printer upper bin` |  |
| `printer only bin` |  |
| `printer lower bin` |  |
| `printer middle bin` |  |
| `printer manual feed` |  |
| `printer envelope feed` |  |
| `printer manual envelope feed` |  |
| `printer automatic sheet feed` |  |
| `printer tractor feed` |  |
| `printer small format bin` |  |
| `printer large format bin` |  |
| `printer large capacity bin` |  |
| `printer paper cassette` |  |
| `printer form source` |  |

### `WdOrientation`

| Value | Description |
|-------|-------------|
| `orient portrait` |  |
| `orient landscape` |  |

### `WdSelectionType`

| Value | Description |
|-------|-------------|
| `no selection` |  |
| `selection ip` |  |
| `selection normal` |  |
| `selection frame` |  |
| `selection column` |  |
| `selection row` |  |
| `selection block` |  |
| `selection inline shape` |  |
| `selection shape` |  |

### `WdCaptionLabelID`

| Value | Description |
|-------|-------------|
| `caption figure` |  |
| `caption table` |  |
| `caption equation` |  |

### `WdReferenceType`

| Value | Description |
|-------|-------------|
| `reference type numbered item` |  |
| `reference type heading` |  |
| `reference type bookmark` |  |
| `reference type footnote` |  |
| `reference type endnote` |  |

### `WdReferenceKind`

| Value | Description |
|-------|-------------|
| `reference content text` |  |
| `reference number relative context` |  |
| `reference number no context` |  |
| `reference number full context` |  |
| `reference entire caption` |  |
| `reference only label and number` |  |
| `reference only caption text` |  |
| `reference footnote number` |  |
| `reference endnote number` |  |
| `reference page number` |  |
| `reference position` |  |
| `reference footnote number formatted` |  |
| `reference endnote number formatted` |  |

### `WdIndexFormat`

| Value | Description |
|-------|-------------|
| `index template` |  |
| `index classic` |  |
| `index fancy` |  |
| `index modern` |  |
| `index bulleted` |  |
| `index formal` |  |
| `index simple` |  |

### `WdIndexType`

| Value | Description |
|-------|-------------|
| `index indent` |  |
| `index runin` |  |

### `WdRevisionsWrap`

| Value | Description |
|-------|-------------|
| `wrap never` |  |
| `wrap always` |  |
| `wrap ask` |  |

### `WdRevisionType`

| Value | Description |
|-------|-------------|
| `no revision` |  |
| `revision insert` |  |
| `revision delete` |  |
| `revision property` |  |
| `revision paragraph number` |  |
| `revision display field` |  |
| `revision reconcile` |  |
| `revision conflict` |  |
| `revision style` |  |
| `revision replace` |  |
| `revision paragraph property` |  |
| `revision table property` |  |
| `revision section property` |  |
| `revision style definition` |  |
| `revision move from` |  |
| `revision move to` |  |
| `revision cell insertion` |  |
| `revision cell deletion` |  |
| `revision cell merge` |  |
| `revision cell split` |  |
| `revision conflict insert` |  |
| `revision conflict delete` |  |

### `WdRoutingSlipDelivery`

| Value | Description |
|-------|-------------|
| `one after another` |  |
| `all at once` |  |

### `WdRoutingSlipStatus`

| Value | Description |
|-------|-------------|
| `not yet routed` |  |
| `route in progress` |  |
| `route complete` |  |

### `WdSectionStart`

| Value | Description |
|-------|-------------|
| `section continuous` |  |
| `section new column` |  |
| `section new page` |  |
| `section even page` |  |
| `section odd page` |  |

### `WdSaveOptions`

| Value | Description |
|-------|-------------|
| `do not save changes` |  |
| `save changes` |  |
| `prompt to save changes` |  |

### `WdDocumentKind`

| Value | Description |
|-------|-------------|
| `document not specified` |  |
| `document letter` |  |
| `document email` |  |

### `WdDocumentType`

| Value | Description |
|-------|-------------|
| `type document` |  |
| `type template` |  |
| `type frameset` |  |

### `WdOriginalFormat`

| Value | Description |
|-------|-------------|
| `word document` |  |
| `original document format` |  |
| `prompt user` |  |

### `WdRelocate`

| Value | Description |
|-------|-------------|
| `relocate up` |  |
| `relocate down` |  |

### `WdInsertedTextMark`

| Value | Description |
|-------|-------------|
| `inserted text mark none` |  |
| `inserted text mark bold` |  |
| `inserted text mark italic` |  |
| `inserted text mark underline` |  |
| `inserted text mark double underline` |  |
| `inserted text mark color only` |  |
| `inserted text mark strike through` |  |
| `inserted text mark double strike through` |  |

### `WdRevisedLinesMark`

| Value | Description |
|-------|-------------|
| `revised lines mark none` |  |
| `revised lines mark left border` |  |
| `revised lines mark right border` |  |
| `revised lines mark outside border` |  |

### `WdDeletedTextMark`

| Value | Description |
|-------|-------------|
| `deleted text mark hidden` |  |
| `deleted text mark strike through` |  |
| `deleted text mark caret` |  |
| `deleted text mark pound` |  |
| `deleted text mark none` |  |
| `deleted text mark bold` |  |
| `deleted text mark italic` |  |
| `deleted text mark underline` |  |
| `deleted text mark double underline` |  |
| `deleted text mark color only` |  |
| `deleted text mark double strike through` |  |

### `WdRevisedPropertiesMark`

| Value | Description |
|-------|-------------|
| `revised properties mark none` |  |
| `revised properties mark bold` |  |
| `revised properties mark italic` |  |
| `revised properties mark underline` |  |
| `revised properties mark double underline` |  |
| `revised properties mark color only` |  |
| `revised properties mark strike through` |  |
| `revised properties mark double strike through` |  |

### `WdMoveToTextMark`

| Value | Description |
|-------|-------------|
| `move to text mark none` |  |
| `move to text mark bold` |  |
| `move to text mark italic` |  |
| `move to text mark underline` |  |
| `move to text mark double underline` |  |
| `move to text mark color only` |  |
| `move to text mark strike through` |  |
| `move to text mark double strike through` |  |

### `WdOMathFunctionType`

| Value | Description |
|-------|-------------|
| `math accent type` |  |
| `math function bar type` |  |
| `math box type` |  |
| `math boxed formula type` |  |
| `math delimiters type` |  |
| `math equation array type` |  |
| `math fraction type` |  |
| `math function apply type` |  |
| `math stretch stack type` |  |
| `math lower limit type` |  |
| `math upper limit type` |  |
| `math matrix type` |  |
| `math nary operator type` |  |
| `math phantom type` |  |
| `math left sub superscript type` |  |
| `math radical type` |  |
| `math subscript type` |  |
| `math sub superscript type` |  |
| `math superscript type` |  |
| `math text type` |  |
| `math normal text type` |  |
| `math literal text type` |  |

### `WdOMathHorizAlignType`

| Value | Description |
|-------|-------------|
| `equation horizontal align center` |  |
| `equation horizontal align left` |  |
| `equation horizontal align right` |  |

### `WdOMathVertAlignType`

| Value | Description |
|-------|-------------|
| `equation vertical align center` |  |
| `equation vertical align top` |  |
| `equation vertical align bottom` |  |

### `WdOMathFracType`

| Value | Description |
|-------|-------------|
| `math fraction type bar` |  |
| `math fraction type stack` |  |
| `math fraction type slashed` |  |
| `math fraction type linear` |  |

### `WdOMathSpacingRule`

| Value | Description |
|-------|-------------|
| `equation spacing single` |  |
| `equation spacing one and a half` |  |
| `equation spacing double` |  |
| `equation spacing exactly` |  |
| `equation spacing multiple` |  |

### `WdOMathType`

| Value | Description |
|-------|-------------|
| `equation display professional` | Specifies an equation is displayed in professional/built up format. |
| `equation display inline` | Specifies an equation is displayed in linear/built down style. |

### `WdOMathShapeType`

| Value | Description |
|-------|-------------|
| `math delimiters center vertically` |  |
| `math delimiters match content size` |  |

### `WdOMathJc`

| Value | Description |
|-------|-------------|
| `equation justification center group` |  |
| `equation justification center` |  |
| `equation justification left` |  |
| `equation justification right` |  |
| `equation justification inline` |  |

### `WdOMathBreakBin`

| Value | Description |
|-------|-------------|
| `math binary operator break before` |  |
| `math binary operator break after` |  |
| `math binary operator repeat` |  |

### `WdOMathBreakSub`

| Value | Description |
|-------|-------------|
| `minus minus` |  |
| `plus minus` |  |
| `minus plus` |  |

### `WdBuildingBlockTypes`

| Value | Description |
|-------|-------------|
| `building block equations` |  |

### `WdDocPartInsertOptions`

| Value | Description |
|-------|-------------|
| `inline building block` |  |
| `page level building block` |  |
| `paragraph level building block` |  |

### `WdMoveFromTextMark`

| Value | Description |
|-------|-------------|
| `move from text mark hidden` |  |
| `move from text mark double strike through` |  |
| `move from text mark strike through` |  |
| `move from text mark caret` |  |
| `move from text mark pound` |  |
| `move from text mark none` |  |
| `move from text mark bold` |  |
| `move from text mark italic` |  |
| `move from text mark underline` |  |
| `move from text mark double underline` |  |
| `move from text mark color only` |  |

### `WdCellColor`

| Value | Description |
|-------|-------------|
| `cell color by author` |  |
| `cell color no highlight` |  |
| `cell color pink` |  |
| `cell color light blue` |  |
| `cell color light yellow` |  |
| `cell color light purple` |  |
| `cell color light orange` |  |
| `cell color light green` |  |
| `cell color light gray` |  |

### `WdFieldShading`

| Value | Description |
|-------|-------------|
| `field shading never` |  |
| `field shading always` |  |
| `field shading when selected` |  |

### `WdDefaultFilePath`

| Value | Description |
|-------|-------------|
| `documents path` |  |
| `pictures path` |  |
| `user templates path` |  |
| `workgroup templates path` |  |
| `user options path` |  |
| `auto recover path` |  |
| `tools path` |  |
| `tutorial path` |  |
| `startup path` |  |
| `program path` |  |
| `graphics filters path` |  |
| `text converters path` |  |
| `proofing tools path` |  |
| `temp file path` |  |
| `current folder path` |  |
| `style gallery path` |  |
| `border art path` |  |

### `WdCompatibility`

| Value | Description |
|-------|-------------|
| `no tab hanging indent` |  |
| `no space for raised or lowered characters` |  |
| `print colors black` |  |
| `wrap trail spaces` |  |
| `no column balance` |  |
| `convert data merge escapes` |  |
| `suppress space before after page break` |  |
| `suppress top spacing` |  |
| `original word table rules` |  |
| `transparent metafiles` |  |
| `show breaks in frames` |  |
| `swap borders facing pages` |  |
| `leave backslash alone` |  |
| `expand shift return` |  |
| `do not underline trailing spaces` |  |
| `do not balance SBCS and DBCS characters` |  |
| `suppress top spacing Mac Word5` |  |
| `spacing in whole points` |  |
| `print body text before header` |  |
| `no extra space between rows of text` |  |
| `no space for underlines` |  |
| `use larger small caps` |  |
| `no extra line spacing` |  |
| `truncate font height` |  |
| `substitute font by size` |  |
| `use printer metrics` |  |
| `Word6 border rules` |  |
| `exact on top` |  |
| `suppress bottom spacing` |  |
| `WordPerfect space width` |  |
| `WordPerfect justification` |  |
| `Word6 line wrap` |  |
| `Word96 shape layout` |  |
| `Word98 footnote layout` |  |
| `do not use html paragraph auto spacing` |  |
| `do not adjust line height in table` |  |
| `forget last tab alignment` |  |
| `Word95 auto space` |  |
| `align tables row by row` |  |
| `layout raw table width` |  |
| `layout table rows apart` |  |
| `use word97 line breaking rules` |  |
| `do not break wrapped tables` |  |
| `do not snap text to grid in table with objects` |  |
| `select field with first or last character` |  |
| `apply breaking rules` |  |
| `do not wrap text with punctuation` |  |
| `do not use asian break rules in grid` |  |
| `use word 2002 table style rules` |  |
| `grow autofit` |  |
| `use normal style for list` |  |
| `do not use indent as numbering tab stop` |  |
| `line break` |  |
| `allow space of same style in table` |  |
| `indent rules` |  |
| `dont autofit constrained tables` |  |
| `autofit like` |  |
| `underline tab in num list` |  |
| `hangul width like` |  |
| `split pg break and para mark` |  |
| `dont vert align cell with shape` |  |
| `dont break constrained forced tables` |  |
| `dont vert align in textbox` |  |
| `word kerning pairs` |  |
| `cached col balance` |  |
| `disable kerning` |  |
| `flip mirror indents` |  |
| `dont override table style font sz and justification` |  |
| `use word table style rules` |  |
| `delayable floating table` |  |
| `allow hyphenation at track bottom` |  |
| `use word track bottom hyphenation` |  |
| `use pre 2018 word mac font layout` |  |

### `WdCompatibilityVersion`

| Value | Description |
|-------|-------------|
| `default compatibility settings` | Microsoft Word 2007-2008 |
| `word 2000 2004 and X` | Microsoft Word 2000-2004 and X |
| `word 97 and 98` | Microsoft Word 97-98 |
| `word 6 and 95` | Microsoft Word 6.0/95 |
| `win word 1` | Word for Windows 1.0 |
| `win word 2` | Word for Windows 2.0 |
| `mac word 5` | Word for the Macintosh 5.x |
| `dos word` | Word for MS-DOS |
| `wordperfect 5` | WordPerfect 5.x |
| `win wordperfect 6` | WordPerfect 6.x for Windows |
| `dos wordperfect 6` | WordPerfect 6.0 for DOS |
| `asian word 97 and 98` | Microsoft Word Asian Versions 97-98 |
| `us word 6 and 95` | US Microsoft Word for Windows 6.0/95 |
| `custom compatibility settings` | Custom settings |

### `WdPaperSize`

| Value | Description |
|-------|-------------|
| `paper ten X fourteen` |  |
| `paper eleven X seventeen` |  |
| `paper letter` |  |
| `paper letter small` |  |
| `paper legal` |  |
| `paper executive` |  |
| `paper a3` |  |
| `paper a4` |  |
| `paper a4 small` |  |
| `paper a5` |  |
| `paper b4` |  |
| `paper b5` |  |
| `paper csheet` |  |
| `paper dsheet` |  |
| `paper esheet` |  |
| `paper fanfold legal german` |  |
| `paper fanfold std german` |  |
| `paper fanfold us` |  |
| `paper folio` |  |
| `paper ledger` |  |
| `paper note` |  |
| `paper quarto` |  |
| `paper statement` |  |
| `paper tabloid` |  |
| `paper envelope9` |  |
| `paper envelope10` |  |
| `paper envelope11` |  |
| `paper envelope12` |  |
| `paper envelope14` |  |
| `paper envelope b4` |  |
| `paper envelope b5` |  |
| `paper envelope b6` |  |
| `paper envelope c3` |  |
| `paper envelope c4` |  |
| `paper envelope c5` |  |
| `paper envelope c6` |  |
| `paper envelope c65` |  |
| `paper envelope dl` |  |
| `paper envelope italy` |  |
| `paper envelope monarch` |  |
| `paper envelope personal` |  |
| `paper custom` |  |

### `WdCustomLabelPageSize`

| Value | Description |
|-------|-------------|
| `custom label letter` |  |
| `custom label letter landscape` |  |
| `custom label A4` |  |
| `custom label A4 landscape` |  |
| `custom label A5` |  |
| `custom label A5 landscape` |  |
| `custom label B5` |  |
| `custom label mini` |  |
| `custom label fanfold` |  |
| `custom label vert half sheet` |  |
| `custom label vert half sheet ls` |  |
| `custom label higaki` |  |
| `custom label higaki ls` |  |
| `custom label b4 jis` |  |

### `WdProtectionType`

| Value | Description |
|-------|-------------|
| `no document protection` |  |
| `allow only revisions` |  |
| `allow only comments` |  |
| `allow only form fields` |  |
| `allow only reading` |  |

### `WdPartOfSpeech`

| Value | Description |
|-------|-------------|
| `adjective` |  |
| `noun` |  |
| `adverb` |  |
| `verb` |  |
| `pronoun` |  |
| `conjunction` |  |
| `preposition` |  |
| `interjection` |  |
| `idiom` |  |
| `other` |  |

### `WdRelativeHorizontalPosition`

| Value | Description |
|-------|-------------|
| `relative horizontal position margin` |  |
| `relative horizontal position page` |  |
| `relative horizontal position column` |  |
| `relative horizontal position character` |  |
| `relative horizontal position left margin` |  |
| `relative horizontal position right margin` |  |
| `relative horizontal position inner margin` |  |
| `relative horizontal position outer margin` |  |

### `WdRelativeVerticalPosition`

| Value | Description |
|-------|-------------|
| `relative vertical position margin` |  |
| `relative vertical position page` |  |
| `relative vertical position paragraph` |  |
| `relative vertical position line` |  |
| `relative vertical position top margin` |  |
| `relative vertical position bottom margin` |  |
| `relative vertical position inner margin` |  |
| `relative vertical position outer margin` |  |

### `WdHelpType`

| Value | Description |
|-------|-------------|
| `help` |  |
| `help about` |  |
| `help contents` |  |
| `help index` |  |
| `help psshelp` |  |
| `help search` |  |

### `WdKeyCategory`

| Value | Description |
|-------|-------------|
| `key category nil` |  |
| `key category disable` |  |
| `key category command` |  |
| `key category macro` |  |
| `key category font` |  |
| `key category auto text` |  |
| `key category style` |  |
| `key category symbol` |  |
| `key category prefix` |  |

### `WdKey`

| Value | Description |
|-------|-------------|
| `no_key` |  |
| `shift_key` |  |
| `control_key` |  |
| `command_key` |  |
| `option_key` |  |
| `key alt` |  |
| `a_key` |  |
| `b_key` |  |
| `c_key` |  |
| `d_key` |  |
| `e_key` |  |
| `f_key` |  |
| `g_key` |  |
| `h_key` |  |
| `i_Key` |  |
| `j_key` |  |
| `k_key` |  |
| `l_key` |  |
| `m_key` |  |
| `n_key` |  |
| `o_key` |  |
| `p_key` |  |
| `q_key` |  |
| `r_key` |  |
| `s_key` |  |
| `t_key` |  |
| `u_kwy` |  |
| `v_key` |  |
| `w_key` |  |
| `x_key` |  |
| `y_key` |  |
| `z_key` |  |
| `key_number_0` |  |
| `key_number_1` |  |
| `key_numbe_2` |  |
| `key_numbe_3` |  |
| `key_number_4` |  |
| `key_number_5` |  |
| `key_number_6` |  |
| `key number_7` |  |
| `key_number_8` |  |
| `key_number_9` |  |
| `backspace_key` |  |
| `tab_key` |  |
| `key_numeric_5_special` |  |
| `return_key` |  |
| `pause_key` |  |
| `esc_key` |  |
| `spacebar_key` |  |
| `page_up_key` |  |
| `page_down_key` |  |
| `end_key` |  |
| `home_key` |  |
| `insert_key` |  |
| `delete_key` |  |
| `key_numeric_0` |  |
| `key_numeric_1` |  |
| `key_numeric_2` |  |
| `key_numeric_3` |  |
| `key_numeric_4` |  |
| `key_numeric_5` |  |
| `key_numeric_6` |  |
| `key_numeric_7` |  |
| `key_numeric_8` |  |
| `key_numeric_9` |  |
| `key_numeric_multiply` |  |
| `key_numeric_add` |  |
| `key_numeric_subtract` |  |
| `key_numeric_decimal` |  |
| `key_numeric_divide` |  |
| `f1_key` |  |
| `f2_key ` |  |
| `f3_key` |  |
| `f4_key` |  |
| `f5_key` |  |
| `f6_key` |  |
| `f7_key` |  |
| `f8_key` |  |
| `f9_key` |  |
| `f10_key` |  |
| `f11_key` |  |
| `f12_key` |  |
| `f13_key` |  |
| `f14_key` |  |
| `f15_key` |  |
| `f16_key` |  |
| `scroll_lock_key` |  |
| `semi_colon_key` |  |
| `equals_key` |  |
| `comma_key` |  |
| `hyphen_key` |  |
| `period_key` |  |
| `slash_key ` |  |
| `back_single_quote_key` |  |
| `open_square_brace_key` |  |
| `back_slash_key` |  |
| `close_square_brace_key` |  |
| `single_quote_key` |  |

### `WdOLEType`

| Value | Description |
|-------|-------------|
| `olelink` |  |
| `oleembed` |  |
| `olecontrol` |  |

### `WdOLEVerb`

| Value | Description |
|-------|-------------|
| `oleverb primary` |  |
| `oleverb show` |  |
| `oleverb open` |  |
| `oleverb hide` |  |
| `oleverb uiactivate` |  |
| `oleverb in place activate` |  |
| `oleverb discard undo state` |  |

### `WdOLEPlacement`

| Value | Description |
|-------|-------------|
| `in line` |  |
| `float over text` |  |

### `WdEnvelopeOrientation`

| Value | Description |
|-------|-------------|
| `left portrait` |  |
| `center portrait` |  |
| `right portrait` |  |
| `left landscape` |  |
| `center landscape` |  |
| `right landscape` |  |
| `left clockwise` |  |
| `center clockwise` |  |
| `right clockwise` |  |

### `WdLetterStyle`

| Value | Description |
|-------|-------------|
| `full block` |  |
| `modified block` |  |
| `semi block` |  |

### `WdLetterheadLocation`

| Value | Description |
|-------|-------------|
| `letter top` |  |
| `letter bottom` |  |
| `letter left` |  |
| `letter right` |  |

### `WdSalutationType`

| Value | Description |
|-------|-------------|
| `salutation informal` |  |
| `salutation formal` |  |
| `salutation business` |  |
| `salutation other` |  |

### `WdSalutationGender`

| Value | Description |
|-------|-------------|
| `gender female` |  |
| `gender male` |  |
| `gender neutral` |  |
| `gender unknown` |  |

### `WdMovementType`

| Value | Description |
|-------|-------------|
| `by moving` |  |
| `by selecting` |  |

### `WdConstants`

| Value | Description |
|-------|-------------|
| `undefined constant` |  |
| `toggle constant` |  |
| `forward constant` |  |
| `backward constant` |  |
| `auto position` |  |
| `first constant` |  |
| `creator code` |  |

### `WdPasteDataType`

| Value | Description |
|-------|-------------|
| `paste oleobject` |  |
| `paste rtf` |  |
| `paste text` |  |
| `paste metafile picture` |  |
| `paste bitmap` |  |
| `paste device independent bitmap` |  |
| `paste hyperlink` |  |
| `paste shape` |  |
| `paste enhanced metafile` |  |
| `paste html` |  |

### `WdPrintOutItem`

| Value | Description |
|-------|-------------|
| `print document content` |  |
| `print properties` |  |
| `print comments` |  |
| `print markup` |  |
| `print styles` |  |
| `print auto text entries` |  |
| `print key assignments` |  |
| `print envelope` |  |
| `print document with markup` |  |

### `WdPrintOutPages`

| Value | Description |
|-------|-------------|
| `print all pages` |  |
| `print odd pages only` |  |
| `print even pages only` |  |

### `WdPrintOutRange`

| Value | Description |
|-------|-------------|
| `print all document` |  |
| `print selection` |  |
| `print current page` |  |
| `print from to` |  |
| `print range of pages` |  |

### `WdDictionaryType`

| Value | Description |
|-------|-------------|
| `spelling` |  |
| `grammar` |  |
| `thesaurus` |  |
| `hyphenation` |  |
| `spelling complete` |  |
| `spelling custom` |  |
| `spelling legal` |  |
| `spelling medical` |  |
| `hangul hanja conversion` |  |
| `hangul hanja conversion custom` |  |

### `WdSpellingWordType`

| Value | Description |
|-------|-------------|
| `spelling word type spell word` |  |
| `spelling word type wildcard` |  |
| `spelling word type anagram` |  |

### `WdSpellingErrorType`

| Value | Description |
|-------|-------------|
| `spelling correct` |  |
| `spelling not in dictionary` |  |
| `spelling capitalization` |  |

### `WdProofreadingErrorType`

| Value | Description |
|-------|-------------|
| `spelling error` |  |
| `grammatical error` |  |

### `WdInlineShapeType`

| Value | Description |
|-------|-------------|
| `inline shape embedded oleobject` |  |
| `inline shape linked oleobject` |  |
| `inline shape picture` |  |
| `inline shape linked picture` |  |
| `inline shape olecontrol object` |  |
| `inline shape horizontal line` |  |
| `inline shape picture horizontal line` |  |
| `inline shape linked picture horizontal line` |  |
| `inline shape picture bullet` |  |
| `inline shape script anchor` |  |
| `inline shape OWS anchor` |  |
| `inline shape chart` |  |
| `inline shape diagram` |  |
| `inline shape locked canvas` |  |
| `inline shape smartart graphic` |  |
| `inline shape web video` |  |
| `inline shape graphic` |  |
| `inline shape linked graphic` |  |
| `inline shape 3dmodel` |  |
| `inline shape linked 3dmodel` |  |

### `WdArrangeStyle`

| Value | Description |
|-------|-------------|
| `tiled` |  |
| `icons` |  |

### `WdSelectionFlags`

| Value | Description |
|-------|-------------|
| `selection start active` |  |
| `selection at eol` |  |
| `selection overtype` |  |
| `selection active` |  |
| `selection replace` |  |
| `selection inactive` |  |
| `selection start active and at eol` |  |
| `selection start active and overtype` |  |
| `selection at eol and overtype` |  |
| `selection start active and at eol and overtype` |  |
| `selection start active and active` |  |
| `selection at eol and active` |  |
| `selection start active and at eol and active` |  |
| `selection overtype and active` |  |
| `selection start active and overtype and active` |  |
| `selection at eol and overtype and active` |  |
| `selection start active and at eol and overtype and active` |  |
| `selection start active and replace` |  |
| `selection at eol and replace` |  |
| `selection start active and at eol and replace` |  |
| `selection overtype and replace` |  |
| `selection start active and overtype and replace` |  |
| `selection at eol and overtype and replace` |  |
| `selection start active and at eol and overtype and replace` |  |
| `selection active and replace` |  |
| `selection start active and active and replace` |  |
| `selection at eol and active and replace` |  |
| `selection start active and at eol and active and replace` |  |
| `selection overtype and active and replace` |  |
| `selection start active and overtype and active and replace` |  |
| `selection at eol and overtype and active and replace` |  |
| `selection start active and at eol and overtype and active and replace` |  |

### `WdAutoVersions`

| Value | Description |
|-------|-------------|
| `auto version off` |  |
| `auto version on close` |  |

### `WdOrganizerObject`

| Value | Description |
|-------|-------------|
| `organizer object styles` |  |
| `organizer object auto text` |  |
| `organizer object command bars` |  |
| `organizer object project items` |  |

### `WdFindMatch`

| Value | Description |
|-------|-------------|
| `match paragraph mark` |  |
| `match tab character` |  |
| `match comment mark` |  |
| `match any character` |  |
| `match any digit` |  |
| `match any letter` |  |
| `match caret character` |  |
| `match column break` |  |
| `match em dash` |  |
| `match en dash` |  |
| `match endnote mark` |  |
| `match field` |  |
| `match footnote mark` |  |
| `match graphic` |  |
| `match manual line break` |  |
| `match manual page break` |  |
| `match nonbreaking hyphen` |  |
| `match nonbreaking space` |  |
| `match optional hyphen` |  |
| `match section break` |  |
| `match white space` |  |

### `WdFindWrap`

| Value | Description |
|-------|-------------|
| `find stop` |  |
| `find continue` |  |
| `find ask` |  |

### `WdInformation`

| Value | Description |
|-------|-------------|
| `active end adjusted page number` |  |
| `active end section number` |  |
| `active end page number` |  |
| `number of pages in document` |  |
| `horizontal position relative to page` |  |
| `vertical position relative to page` |  |
| `horizontal position relative to text boundary` |  |
| `vertical position relative to text boundary` |  |
| `first character column number` |  |
| `first character line number` |  |
| `frame is selected` |  |
| `with in table` |  |
| `start of range row number` |  |
| `end_of range row number` |  |
| `maximum number of rows` |  |
| `start of range column number` |  |
| `end_of range column number` |  |
| `maximum number of columns` |  |
| `zoom percentage` |  |
| `selection mode` |  |
| `info caps lock` |  |
| `info num lock` |  |
| `over type` |  |
| `revision marking` |  |
| `in footnote endnote pane` |  |
| `in comment pane` |  |
| `in header footer` |  |
| `at end of row marker` |  |
| `reference of type` |  |
| `header footer type` |  |
| `in master document` |  |
| `in footnote` |  |
| `in endnote` |  |
| `in word mail` |  |
| `in clipboard` |  |
| `in cover page` |  |
| `in bibliography` |  |
| `in citation` |  |
| `in field code` |  |
| `in field result` |  |
| `in content control` |  |

### `WdWrapType`

| Value | Description |
|-------|-------------|
| `wrap square` |  |
| `wrap tight` |  |
| `wrap through` |  |
| `wrap none` |  |
| `wrap top bottom` |  |
| `wrap behind` |  |
| `wrap front` |  |
| `wrap inline` |  |

### `WdWrapTypeMerged`

| Value | Description |
|-------|-------------|
| `picture wrap type inline with text` |  |
| `picture wrap type square` |  |
| `picture wrap type tight` |  |
| `picture wrap type behind text` |  |
| `picture wrap type in front of text` |  |
| `picture wrap type through` |  |
| `picture wrap type top and bottom` |  |

### `WdWrapSideType`

| Value | Description |
|-------|-------------|
| `wrap both` |  |
| `wrap left` |  |
| `wrap right` |  |
| `wrap largest` |  |

### `WdOutlineLevel`

| Value | Description |
|-------|-------------|
| `outline level1` |  |
| `outline level2` |  |
| `outline level3` |  |
| `outline level4` |  |
| `outline level5` |  |
| `outline level6` |  |
| `outline level7` |  |
| `outline level8` |  |
| `outline level9` |  |
| `outline level body text` |  |

### `WdTextOrientation`

| Value | Description |
|-------|-------------|
| `text orientation horizontal` |  |
| `text orientation upward` |  |
| `text orientation downward` |  |
| `text orientation vertical east asian` |  |
| `text orientation horizontal rotated east asian` |  |
| `text orientation vertical` |  |

### `WdPageBorderArt`

| Value | Description |
|-------|-------------|
| `art none` |  |
| `art apples` |  |
| `art maple muffins` |  |
| `art cake slice` |  |
| `art candy corn` |  |
| `art ice cream cones` |  |
| `art champagne bottle` |  |
| `art party glass` |  |
| `art christmas tree` |  |
| `art trees` |  |
| `art palms color` |  |
| `art balloons3 colors` |  |
| `art balloons hot air` |  |
| `art party favor` |  |
| `art confetti streamers` |  |
| `art hearts` |  |
| `art heart balloon` |  |
| `art stars3D` |  |
| `art stars shadowed` |  |
| `art stars` |  |
| `art sun` |  |
| `art earth2` |  |
| `art earth1` |  |
| `art people hats` |  |
| `art sombrero` |  |
| `art pencils` |  |
| `art packages` |  |
| `art clocks` |  |
| `art firecrackers` |  |
| `art rings` |  |
| `art map pins` |  |
| `art confetti` |  |
| `art creatures butterfly` |  |
| `art creatures lady bug` |  |
| `art creatures fish` |  |
| `art birds flight` |  |
| `art scared cat` |  |
| `art bats` |  |
| `art flowers roses` |  |
| `art flowers red rose` |  |
| `art poinsettias` |  |
| `art holly` |  |
| `art flowers tiny` |  |
| `art flowers pansy` |  |
| `art flowers modern2` |  |
| `art flowers modern1` |  |
| `art white flowers` |  |
| `art vine` |  |
| `art flowers daisies` |  |
| `art flowers block print` |  |
| `art deco arch color` |  |
| `art fans` |  |
| `art film` |  |
| `art lightning1` |  |
| `art compass` |  |
| `art double d` |  |
| `art classical wave` |  |
| `art shadowed squares` |  |
| `art twisted lines1` |  |
| `art waveline` |  |
| `art quadrants` |  |
| `art checked bar color` |  |
| `art swirligig` |  |
| `art push pin note1` |  |
| `art push pin note2` |  |
| `art pumpkin1` |  |
| `art eggs black` |  |
| `art cup` |  |
| `art heart gray` |  |
| `art gingerbread man` |  |
| `art baby pacifier` |  |
| `art baby rattle` |  |
| `art cabins` |  |
| `art house funky` |  |
| `art stars black` |  |
| `art snowflakes` |  |
| `art snowflake fancy` |  |
| `art skyrocket` |  |
| `art seattle` |  |
| `art music notes` |  |
| `art palms black` |  |
| `art maple leaf` |  |
| `art paper clips` |  |
| `art shorebird tracks` |  |
| `art people` |  |
| `art people waving` |  |
| `art eclipsing squares2` |  |
| `art hypnotic` |  |
| `art diamonds gray` |  |
| `art deco arch` |  |
| `art deco blocks` |  |
| `art circles lines` |  |
| `art papyrus` |  |
| `art woodwork` |  |
| `art weaving braid` |  |
| `art weaving ribbon` |  |
| `art weaving angles` |  |
| `art arched scallops` |  |
| `art safari` |  |
| `art celtic knotwork` |  |
| `art crazy maze` |  |
| `art eclipsing squares1` |  |
| `art birds` |  |
| `art flowers teacup` |  |
| `art northwest` |  |
| `art southwest` |  |
| `art tribal6` |  |
| `art tribal4` |  |
| `art tribal3` |  |
| `art tribal2` |  |
| `art tribal5` |  |
| `art xillusions` |  |
| `art zany triangles` |  |
| `art pyramids` |  |
| `art pyramids above` |  |
| `art confetti grays` |  |
| `art confetti outline` |  |
| `art confetti white` |  |
| `art mosaic` |  |
| `art lightning2` |  |
| `art heebie jeebies` |  |
| `art light bulb` |  |
| `art gradient` |  |
| `art triangle party` |  |
| `art twisted lines2` |  |
| `art moons` |  |
| `art ovals` |  |
| `art double diamonds` |  |
| `art chain link` |  |
| `art triangles` |  |
| `art tribal1` |  |
| `art marquee toothed` |  |
| `art sharks teeth` |  |
| `art sawtooth` |  |
| `art sawtooth gray` |  |
| `art postage stamp` |  |
| `art weaving strips` |  |
| `art zig zag` |  |
| `art cross stitch` |  |
| `art gems` |  |
| `art circles rectangles` |  |
| `art corner triangles` |  |
| `art creatures insects` |  |
| `art zig zag stitch` |  |
| `art checkered` |  |
| `art checked bar black` |  |
| `art marquee` |  |
| `art basic white dots` |  |
| `art basic wide midline` |  |
| `art basic wide outline` |  |
| `art basic wide inline` |  |
| `art basic thin lines` |  |
| `art basic white dashes` |  |
| `art basic white squares` |  |
| `art basic black squares` |  |
| `art basic black dashes` |  |
| `art basic black dots` |  |
| `art stars top` |  |
| `art certificate banner` |  |
| `art handmade1` |  |
| `art handmade2` |  |
| `art torn paper` |  |
| `art torn paper black` |  |
| `art coupon cutout dashes` |  |
| `art coupon cutout dots` |  |

### `WdBorderDistanceFrom`

| Value | Description |
|-------|-------------|
| `border distance from text` |  |
| `border distance from page edge` |  |

### `WdReplace`

| Value | Description |
|-------|-------------|
| `replace none` |  |
| `replace one` |  |
| `replace all` |  |

### `WdFontBias`

| Value | Description |
|-------|-------------|
| `font bias do not care` |  |
| `font bias default` |  |
| `font bias east asian` |  |

### `WdBrowserLevel`

| Value | Description |
|-------|-------------|
| `browser level v4` |  |
| `browser level microsoft internet explorer5` |  |
| `browser level microsoft internet explorer6` |  |

### `WdEnclosureType`

| Value | Description |
|-------|-------------|
| `enclosure circle` |  |
| `enclosure square` |  |
| `enclosure triangle` |  |
| `enclosure diamond` |  |

### `WdEncloseStyle`

| Value | Description |
|-------|-------------|
| `enclose style none` |  |
| `enclose style small` |  |
| `enclose style large` |  |

### `WdLayoutMode`

| Value | Description |
|-------|-------------|
| `layout mode default` |  |
| `layout mode grid` |  |
| `layout mode line grid` |  |
| `layout mode genko` |  |

### `WdDocumentMedium`

| Value | Description |
|-------|-------------|
| `for a email message` |  |
| `for a document` |  |
| `for a web page` |  |

### `WdTwoLinesInOneType`

| Value | Description |
|-------|-------------|
| `two lines in one none` |  |
| `two lines in one no brackets` |  |
| `two lines in one parentheses` |  |
| `two lines in one square brackets` |  |
| `two lines in one angle brackets` |  |
| `two lines in one curly brackets` |  |

### `WdHorizontalInVerticalType`

| Value | Description |
|-------|-------------|
| `horizontal in vertical none` |  |
| `horizontal in vertical fit in line` |  |
| `horizontal in vertical resize line` |  |

### `WdHorizontalLineAlignment`

| Value | Description |
|-------|-------------|
| `horizontal line align left` |  |
| `horizontal line align center` |  |
| `horizontal line align right` |  |

### `WdHorizontalLineWidthType`

| Value | Description |
|-------|-------------|
| `horizontal line percent width` |  |
| `horizontal line fixed width` |  |

### `WdPhoneticGuideAlignmentType`

| Value | Description |
|-------|-------------|
| `phonetic guide alignment center` |  |
| `phonetic guide alignment zero one zero` |  |
| `phonetic guide alignment one two one` |  |
| `phonetic guide alignment left` |  |
| `phonetic guide alignment right` |  |
| `phonetic guide alignment right vertical` |  |

### `WdNumberStyleWordBasicBiDi`

| Value | Description |
|-------|-------------|
| `list number style bidi 1` |  |
| `list number style bidi 2` |  |
| `caption number style bidi letter 1` |  |
| `caption number style bidi letter 2` |  |
| `note number style bidi letter 1` |  |
| `note number style bidi letter 2` |  |
| `page number style bidi letter 1` |  |
| `page number style bidi letter 2` |  |

### `WdTableDirection`

| Value | Description |
|-------|-------------|
| `table direction rtl` |  |
| `table direction ltr` |  |

### `WdGutterStyle`

| Value | Description |
|-------|-------------|
| `gutter position left` |  |
| `gutter position top` |  |
| `gutter position right` |  |

### `WdGutterStyleOld`

| Value | Description |
|-------|-------------|
| `gutter style latin` |  |
| `gutter style bidi` |  |
| `gutter style none` |  |

### `WdShapePosition`

| Value | Description |
|-------|-------------|
| `shape top` |  |
| `shape left` |  |
| `shape bottom` |  |
| `shape right` |  |
| `shape center` |  |
| `shape inside` |  |
| `shape outside` |  |

### `WdTablePosition`

| Value | Description |
|-------|-------------|
| `table top` |  |
| `table left` |  |
| `table bottom` |  |
| `table right` |  |
| `table center` |  |
| `table inside` |  |
| `table outside` |  |

### `WdDefaultTableBehavior`

| Value | Description |
|-------|-------------|
| `word8 table behavior` |  |
| `word9 table behavior` |  |

### `WdAutoFitBehavior`

| Value | Description |
|-------|-------------|
| `auto fit fixed` |  |
| `auto fit content` |  |
| `auto fit window` |  |

### `WdDefaultListBehavior`

| Value | Description |
|-------|-------------|
| `word8 list behavior` |  |
| `word9 list behavior` |  |
| `word10 list behavior` |  |

### `WdPreferredWidthType`

| Value | Description |
|-------|-------------|
| `preferred width auto` |  |
| `preferred width percent` |  |
| `preferred width points` |  |

### `WdNewDocumentType`

| Value | Description |
|-------|-------------|
| `new blank document` |  |
| `new web page` |  |
| `new xml document` |  |

### `WdRecoveryType`

| Value | Description |
|-------|-------------|
| `paste default` |  |
| `single cell text` |  |
| `single cell table` |  |
| `list continue numbering` |  |
| `list restart numbering` |  |
| `table insert as rows` |  |
| `table append table` |  |
| `table original formatting` |  |
| `chart picture` |  |
| `chart` |  |
| `chart linked` |  |
| `format original formatting` |  |
| `format surrounding formatting with emphasis` |  |
| `format plain text` |  |
| `table overwrite cells` |  |
| `list combine with existing list` |  |
| `list dont merge` |  |
| `use destination styles recovery` |  |

### `WdRecoveryType`

| Value | Description |
|-------|-------------|
| `go forward` |  |
| `go backward` |  |
| `a numeric constant` |  |

### `WdLineEndingType`

| Value | Description |
|-------|-------------|
| `line ending cr lf` |  |
| `line ending cr only` |  |
| `line ending lf only` |  |
| `line ending lf cr` |  |
| `line ending ls ps` |  |

### `WdConditionCode`

| Value | Description |
|-------|-------------|
| `condition first row` |  |
| `condition last row` |  |
| `condition odd row banding` |  |
| `condition even row banding` |  |
| `condition first column` |  |
| `condition last column` |  |
| `condition odd column banding` |  |
| `condition even column banding` |  |
| `condition top right cell` |  |
| `condition top left cell` |  |
| `condition bottom right cell` |  |
| `condition bottom left cell` |  |

### `WdUnits`

| Value | Description |
|-------|-------------|
| `unit a line` |  |
| `unit a story` |  |
| `unit a screen` |  |
| `unit a section` |  |
| `unit a column` |  |
| `unit a row` |  |

### `WdHighlightStuff`

| Value | Description |
|-------|-------------|
| `highlight on` |  |
| `highlight off` |  |
| `a numeric constant` |  |

### `WdCompareTarget`

| Value | Description |
|-------|-------------|
| `compare target selected` |  |
| `compare target current` |  |
| `compare target new` |  |

### `WdMergeTarget`

| Value | Description |
|-------|-------------|
| `merge target selected` |  |
| `merge target current` |  |
| `merge target new` |  |

### `WdRevisionsView`

| Value | Description |
|-------|-------------|
| `revisions view final` |  |
| `revisions view original` |  |

### `WdRevisionsMode`

| Value | Description |
|-------|-------------|
| `balloon revisions` |  |
| `in line revisions` |  |
| `mixed revisions` |  |

### `WdRevisionsBalloonPrintOrientation`

| Value | Description |
|-------|-------------|
| `balloon print orientation automatic` |  |
| `balloon print orientation preserve` |  |
| `balloon print orientation landscape` |  |

### `WdRevisionsBalloonMargin`

| Value | Description |
|-------|-------------|
| `balloon margin left` |  |
| `balloon margin right` |  |

### `WdCheckInVersionType`

| Value | Description |
|-------|-------------|
| `minor version` |  |
| `major version` |  |
| `overwrite current version` |  |

### `WdLockType`

| Value | Description |
|-------|-------------|
| `lock none` |  |
| `lock reservation` |  |
| `lock ephemeral` |  |
| `lock changed` |  |

### `unlink`

| Value | Description |
|-------|-------------|
| `field options` |  |
| `field` |  |

### `update source`

| Value | Description |
|-------|-------------|
| `field options` |  |
| `field` |  |

### `accept`

| Value | Description |
|-------|-------------|
| `revision` |  |
| `conflict` |  |

### `reject`

| Value | Description |
|-------|-------------|
| `revision` |  |
| `conflict` |  |

### `reset separator`

| Value | Description |
|-------|-------------|
| `footnote options` |  |
| `endnote options` |  |

### `reset continuation separator`

| Value | Description |
|-------|-------------|
| `footnote options` |  |
| `endnote options` |  |

### `reset continuation notice`

| Value | Description |
|-------|-------------|
| `footnote options` |  |
| `endnote options` |  |

### `activate object`

| Value | Description |
|-------|-------------|
| `document` |  |
| `window` |  |
| `pane` |  |

### `get border`

| Value | Description |
|-------|-------------|
| `font` |  |
| `frame` |  |
| `selection object` |  |

### `cut object`

| Value | Description |
|-------|-------------|
| `field` |  |
| `frame` |  |
| `form field` |  |
| `data merge field` |  |
| `selection object` |  |
| `page number` |  |

### `can continue previous list`

| Value | Description |
|-------|-------------|
| `list format` |  |
| `Word list` |  |

### `check grammar`

| Value | Description |
|-------|-------------|
| `application` |  |
| `document` |  |

### `check spelling`

| Value | Description |
|-------|-------------|
| `application` |  |
| `document` |  |

### `clear formatting`

| Value | Description |
|-------|-------------|
| `find` |  |
| `replacement` |  |
| `selection object` |  |

### `clear`

| Value | Description |
|-------|-------------|
| `drop cap` |  |
| `tab stop` |  |
| `text input` |  |
| `key binding` |  |

### `count numbered items`

| Value | Description |
|-------|-------------|
| `document` |  |
| `list format` |  |
| `Word list` |  |

### `copy object`

| Value | Description |
|-------|-------------|
| `field` |  |
| `frame` |  |
| `form field` |  |
| `data merge field` |  |
| `selection object` |  |
| `page number` |  |

### `large scroll`

| Value | Description |
|-------|-------------|
| `window` |  |
| `pane` |  |

### `convert numbers to text`

| Value | Description |
|-------|-------------|
| `document` |  |
| `list format` |  |
| `Word list` |  |

### `page scroll`

| Value | Description |
|-------|-------------|
| `window` |  |
| `pane` |  |

### `print out`

| Value | Description |
|-------|-------------|
| `application` |  |
| `document` |  |
| `window` |  |

### `remove numbers`

| Value | Description |
|-------|-------------|
| `document` |  |
| `list format` |  |
| `Word list` |  |

### `small scroll`

| Value | Description |
|-------|-------------|
| `window` |  |
| `pane` |  |

### `update page numbers`

| Value | Description |
|-------|-------------|
| `table of figures` |  |
| `table of contents` |  |

### `update`

| Value | Description |
|-------|-------------|
| `link format` |  |
| `field options` |  |
| `table of figures` |  |
| `table of contents` |  |
| `table of authorities` |  |
| `dialog` |  |
| `index` |  |

---

## Drawing Suite

Classes and Methods used for Graphic Objects

### Commands

### `apply`

Applies to the specified shape formatting that has been copied using the pick up method.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `automatic length`

Specifies that the first segment of the callout line the segment attached to the text callout box be scaled automatically when the callout is moved. Applies only to callouts whose lines consist of more than one segment.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4027` | no |  |

### `break forward link`

Breaks the forward link for the specified text frame, if such a link exists.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text frame` | no |  |

### `convert to frame`

Converts the specified shape to a frame. Returns a frame object that represents the new frame.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

**Returns:** `frame`

### `convert to inline shape`

Converts the specified shape in the drawing layer of a document to an inline shape in the text layer. You can convert only picture shapes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

**Returns:** `inline shape`

### `convert to shape`

Converts an inline shape to a free-floating shape.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `inline shape` | no |  |

**Returns:** `shape` — The newly converted shape object.

### `custom drop`

Sets the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4028` | no |  |
| `drop` | `real` | no |  |

### `custom length`

Specifies that the first segment of the callout, line the segment attached to the text callout box, retain a fixed length whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4029` | no |  |
| `length` | `real` | no |  |

### `flip`

Flips a shape horizontally or vertically.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |
| `flip command` | `MsoFlipCmd` | no | Specifies if the shape will be flipped horizontally or vertically. |

### `get border`

Returns the specified border object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `inline shape` | no |  |
| `which border` | `WdBorderType` | no |  |

**Returns:** `border`

### `get custom color`

Returns the custom color for the specified Microsoft Office theme.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme color scheme` | no |  |
| `name` | `text` | no |  |

**Returns:** `integer`

### `load theme color scheme`

Loads the color scheme of a Microsoft Office theme from a file

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme color scheme` | no |  |
| `file name` | `text` | no | The name of the color theme file. |

### `load theme effect scheme`

Loads the effects scheme of a Microsoft Office theme from a file

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme effect scheme` | no |  |
| `file name` | `text` | no | The name of the effect scheme file. |

### `load theme font scheme`

Loads the font scheme of a Microsoft Office theme from a file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme font scheme` | no |  |
| `file name` | `text` | no | The name of the font scheme file. |

### `one color gradient`

Sets the specified fill to a one-color gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4031` | no |  |
| `gradient style` | `MsoGradientStyle` | no | Specifies the gradient style. |
| `gradient variant` | `integer` | no | The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the gradient tab in the fill effects dialog box. If gradient style is from center gradient, this argument can be either 1 or 2. |
| `gradient degree` | `real` | no | The gradient degree. Can be a value from 0.0 which is dark to 1.0 which is light. |

### `patterned`

Sets the specified fill to a pattern.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4032` | no |  |
| `pattern` | `MsoPatternType` | no | The pattern to be used for the specified fill. |

### `pick up`

Copies the formatting of the specified shape. Use the apply method to apply the copied formatting to another shape.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `preset drop`

Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4030` | no |  |
| `DropType` | `MsoCalloutDropType` | no |  |

### `preset gradient`

Sets the specified fill to a preset gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4033` | no |  |
| `style` | `MsoGradientStyle` | no | Specifies the gradient style. |
| `gradient variant` | `integer` | no | The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the gradient tab in the fill effects dialog box. If gradient style is from center gradient, this argument can be either 1 or 2. |
| `preset gradient type` | `MsoPresetGradientType` | no | The gradient type. |

### `preset textured`

Sets the specified fill to a preset texture

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4034` | no |  |
| `preset texture` | `MsoPresetTexture` | no | The preset texture |

### `reroute connections`

Reroutes the connections between shapes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `reset`

Removes changes that were made to an inline shape.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `inline shape` | no |  |

### `reset rotation`

Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4039` | no |  |

### `save as picture`

Saves the shape in the requested file using the stated graphic format

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4026` | no |  |
| `picture type` | `MsoPictureType` | yes | Specifies the graphic format in which the file is saved |
| `file name` | `text` | yes | The name and path for the picture being saved |

### `save theme color scheme`

Saves the color scheme of a Microsoft Office theme to a file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme color scheme` | no |  |
| `file name` | `text` | no | The name of the color theme file. |

### `save theme font scheme`

Saves the font scheme of a Microsoft Office theme to a file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme font scheme` | no |  |
| `file name` | `text` | no | The name of the font scheme file. |

### `scale height`

Scales the height of the picture shape by a specified factor.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `picture` | no |  |
| `factor` | `real` | no | Specifies the ratio between the height of the shape after you resize it and the current or original height. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument. |
| `relative to original size` | `boolean` | no | Set to true to scale the shape relative to its original size. False to scale it relative to its current size. |
| `scale` | `MsoScaleFrom` | no | Specifies the part of the shape that retains its position when the shape is scaled. |

### `scale width`

Scales the width of the shape by a specified factor.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `picture` | no |  |
| `factor` | `real` | no | Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument. |
| `relative to original size` | `boolean` | no | Set to true to scale the shape relative to its original size. False to scale it relative to its current size. |
| `scale` | `MsoScaleFrom` | no | Specifies the part of the shape that retains its position when the shape is scaled. |

### `set shapes default properties`

Applies the formatting of the specified shape to a default shape for that document. New shapes inherit many of their attributes from the default shape.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `solid`

Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4035` | no |  |

### `toggle vertical text`

Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `word art format` | no |  |

### `two color gradient`

Sets the specified fill to a two-color gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4036` | no |  |
| `gradient style` | `MsoGradientStyle` | no | Specifies the gradient style. |
| `gradient variant` | `integer` | no | The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the gradient tab in the fill effects dialog box. If gradient style is from center gradient, this argument can be either 1 or 2. |

### `user picture`

Fills the specified shape with one large image. If you want to fill the shape with small tiles of an image, use the user textured method.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4037` | no |  |
| `picture file` | `text` | no | The name of the picture file. |

### `user textured`

Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the user picture method.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4038` | no |  |
| `texture file` | `text` | no | The name of the picture file. |

### `valid link target`

Determines whether the text frame of one shape can be linked to the text frame of another shape. Returns false if target textframe already contains text or is already linked, or if the shape doesn't support attached text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text frame` | no |  |
| `target textframe` | `text frame` | no | The target text frame that you'd like to link the specified text frame to. |

**Returns:** `boolean`

### `z order`

Moves the specified shape in front of or behind other shapes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |
| `z order command` | `MsoZOrderCmd` | no |  |

### Classes

### `adjustment`

*Plural:* `adjustments`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `adjustment_value` | `real` | r/w | Returns or sets a floating point adjustment for a shape. |

### `callout format`

Contains properties and methods that apply to line callouts.

*Plural:* `callout formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accent` | `boolean` | r/w | Returns or sets if a vertical accent bar separates the callout text from the callout line. |
| `angle` | `MsoCalloutAngleType` | r/w | Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. |
| `auto attach` | `boolean` | r/w | Returns or sets if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line, where the callout points to, is to the left or right of the callout text box. |
| `auto length` | `boolean` | r/o | Returns if the length of the callout line is automatically set. Use the automatic length method to set this property to true, and use the custom length method to set this property to false. |
| `callout format length` | `real` | r/o | When the AutoLength property of the specified callout is set to False, the Length property returns the length in points of the first segment of the callout line, the segment attached to the text callout box. |
| `callout has border` | `boolean` | r/w | Returns or sets whether the text in the specified callout is surrounded by a border. |
| `callout type` | `MsoCalloutType` | r/w | Returns or sets the callout type. |
| `drop` | `real` | r/o | For callouts with an explicitly set drop value, this property returns the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box. |
| `drop type` | `MsoCalloutDropType` | r/o | Returns a value that indicates where the callout line attaches to the callout text box. |
| `gap` | `real` | r/w | Returns or sets the horizontal distance in points between the end of the callout line and the text bounding box. |

### `callout`

Represents an callout object in the drawing layer.

*Plural:* `callouts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `callout format` | `callout format` | r/o | Returns the callout format object associated with this callout shape object. |
| `callout type` | `MsoCalloutType` | r/o | Return the type of this callout |

### `fill format`

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.

*Plural:* `fill formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `back color` | `integer` | r/w | Returns or sets a RGB color that represents the background color for the specified fill or patterned line. |
| `back color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the specified fill background color. |
| `fill type` | `MsoFillType` | r/o | Returns the shape fill format type. |
| `fore color` | `integer` | r/w | Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow. |
| `fore color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the specified foreground fill or solid color. |
| `gradient color type` | `MsoGradientColorType` | r/o | Returns the gradient color type for the specified fill. |
| `gradient degree` | `real` | r/o | Returns a value that indicates how dark or light a one-color gradient fill is. A value of zero means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in. Values between 1 and zero blend. |
| `gradient style` | `MsoGradientStyle` | r/o | Returns the gradient style for the specified fill. |
| `gradient variant` | `integer` | r/o | Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. If the gradient style is from center gradient, this property returns either 1 or 2. |
| `pattern` | `MsoPatternType` | r/o | Returns the value that represents the pattern applied to the specified fill or line. |
| `preset gradient type` | `MsoPresetGradientType` | r/o | Returns the preset gradient type for the specified fill. |
| `preset texture` | `MsoPresetTexture` | r/o | Returns the preset texture for the specified fill. |
| `texture name` | `text` | r/o | Returns the name of the custom texture file for the specified fill. |
| `texture type` | `MsoTextureType` | r/o | Returns the texture type for the specified fill. |
| `transparency` | `real` | r/w | Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear. |
| `visible` | `boolean` | r/w | Returns or sets if the specified object, or the formatting applied to it, is visible. |

### `glow format`

Represents the glow formatting for a shape or range of shapes.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `integer` | r/w | Returns or sets the color for the specified glow format. |
| `color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified glow format. |
| `radius` | `real` | r/w | Returns or sets the length of the radius for the specified glow format. |

### `horizontal line format`

Represents horizontal line formatting.

*Plural:* `horizontal line formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `WdHorizontalLineAlignment` | r/w | Returns or sets a constant that represents the alignment for the specified horizontal line. |
| `no shade` | `boolean` | r/w | Returns or sets if Microsoft Word draws the specified horizontal line without 3-D shading. |
| `percent width` | `real` | r/w | Returns or sets the length of the specified horizontal line expressed as a percentage of the window width. |
| `width type` | `WdHorizontalLineWidthType` | r/w | Returns or sets the width type for the specified horizontal line format object. |

### `inline horizontal line`

Represents a horizontal line object in the text layer of a document.

*Plural:* `inline horizontal lines`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `text` | r/o | The file name that contains the picture of the horizontal line. |

### `inline picture bullet`

Represents a picture bullet object in the text layer of a document.

*Plural:* `inline picture bullets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `text` | r/o | The file name that contains the picture of the picture bullet. |

### `inline picture`

Represents a picture object in the text layer of a document.

*Plural:* `inline pictures`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `text` | r/o | The file name that contains the picture. |
| `link to file` | `boolean` | r/o | Returns true if the picture shape is linked to the file. |
| `picture format` | `picture format` | r/o | Returns the picture format object associated with this inline picture object. |
| `save with document` | `boolean` | r/o | Specifies if the picture should be saved with the document. |

### `inline shape`

Represents a graphic object in the text layer of a document.

*Plural:* `inline shapes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alternative text` | `text` | r/w | Returns or sets the alternative text associated with a shape in a Web page. |
| `anchorID` | `integer` | r/o | Returns the anchor id for the specified shape. |
| `border options` | `border options` | r/o | Returns back border options object associated with the inline shape |
| `editID` | `integer` | r/o | Returns the edit id for the specified shape. |
| `field` | `field` | r/o | Returns a field object that represents the field associated with the specified inline shape. |
| `fill format` | `fill format` | r/o | Returns the fill format object associated with this inline shape object. |
| `glow format` | `glow format` | r/o | Returns the formatting properties for a glow effect. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `horizontal line format` | `horizontal line format` | r/o | Returns the horizontal line format object associated with this inline shape object. |
| `hyperlink` | `hyperlink object` | r/o | Returns the hyperlink object associated with this inline shape object. |
| `inline shape scale height` | `real` | r/w | Returns or sets the height of the specified inline shape relative to its original size. |
| `inline shape scale width` | `real` | r/w | Returns or sets the width of the specified inline shape relative to its original size. |
| `inline shape type` | `WdInlineShapeType` | r/o | The type of this inline shape. |
| `is picture bullet` | `boolean` | r/o | Returns true if an InlineShape object is a picture bullet. |
| `line format` | `line format` | r/o | Returns the line format object associated with this inline shape object. |
| `link format` | `link format` | r/o | Returns the link format object associated with this inline shape object. |
| `lock aspect ratio` | `boolean` | r/w | Returns or set if the specified shape retains its original proportions when you resize it. |
| `picture format` | `picture format` | r/o | Returns the picture format object associated with this inline shape object. |
| `reflection format` | `reflection format` | r/o | Returns the formatting properties for a reflection effect. |
| `shadow` | `shadow format` | r/o | Returns the shadow format object associated with this shape object. |
| `soft edge format` | `soft edge format` | r/o | Returns the formatting properties for a soft edge effect. |
| `text object` | `text range` | r/o | Returns a text range object that represents the text for the inline shape. |
| `width` | `real` | r/w | Returns or sets the width of the object. |
| `word art format` | `word art format` | r/o | Returns the word art format object associated with the word art shape object. |

### `line format`

Represents line and arrowhead formatting. For a line, the line format object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.

*Plural:* `line formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `back color` | `integer` | r/w | Returns or sets a RGB color that represents the background color for the specified fill or patterned line. |
| `back color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the background color for a patterned line. |
| `begin arrowhead length` | `MsoArrowheadLength` | r/w | Returns or sets the length of the arrowhead at the beginning of the specified line. |
| `begin arrowhead style` | `MsoArrowheadStyle` | r/w | Returns or sets the style of the arrowhead at the beginning of the specified line. |
| `begin arrowhead width` | `MsoArrowheadWidth` | r/w | Returns or sets the width of the arrowhead at the beginning of the specified line. |
| `dash style` | `MsoLineDashStyle` | r/w | Returns or sets the dash style for the specified line. |
| `end arrowhead length` | `MsoArrowheadLength` | r/w | Returns or sets the length of the arrowhead at the end of the specified line. |
| `end arrowhead style` | `MsoArrowheadStyle` | r/w | Returns or sets the style of the arrowhead at the end of the specified line. |
| `end arrowhead width` | `MsoArrowheadWidth` | r/w | Returns or sets the width of the arrowhead at the end of the specified line. |
| `fore color` | `integer` | r/w | Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow. |
| `fore color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the foreground color for the line. |
| `line style` | `MsoLineStyle` | r/w | Returns or sets the line format style. |
| `pattern` | `MsoPatternType` | r/w | Returns or sets a value that represents the pattern applied to the specified fill or line. |
| `transparency` | `real` | r/w | Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear. |
| `visible` | `boolean` | r/w | Returns or sets if the specified object, or the formatting applied to it, is visible. |
| `weight` | `real` | r/w | Returns or sets the thickness of the specified line in points. |

### `line shape`

The line shape uses begin line X, begin line Y, end line X, and end line Y when created

*Plural:* `line shapes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `begin line X` | `real` | r/w | Returns or sets the starting X coordinate for the line shape. |
| `begin line Y` | `real` | r/w | Returns or sets the starting Y coordinate for the line shape. |
| `end line X` | `real` | r/w | Returns or sets the ending X coordinate for the line shape. |
| `end line Y` | `real` | r/w | Returns or sets the ending Y coordinate for the line shape. |

### `major theme font`

Represents a container for the font schemes of a Microsoft Office 2007 theme.

*Plural:* `major theme fonts`

### `minor theme font`

Represents a container for the font schemes of a Microsoft Office 2007 theme.

*Plural:* `minor theme fonts`

### `office theme`

Represents a Microsoft Office theme.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `theme color scheme` | `theme color scheme` | r/o | Returns the color scheme of a Microsoft Office theme. |
| `theme effect scheme` | `theme effect scheme` | r/o | Returns the effects scheme of a Microsoft Office theme. |
| `theme font scheme` | `theme font scheme` | r/o | Returns the font scheme of a Microsoft Office theme. |

### `picture format`

Contains properties and methods that apply to pictures.

*Plural:* `picture formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `brightness` | `real` | r/w | Returns or sets the brightness of the specified picture . The value for this property must be a number from 0.0, dimmest to 1.0, brightest. |
| `contrast` | `real` | r/w | Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast. |
| `crop bottom` | `real` | r/w | Returns or sets the number of points that are cropped off the bottom of the specified picture. |
| `crop left` | `real` | r/w | Returns or sets the number of points that are cropped off the left side of the specified picture. |
| `crop right` | `real` | r/w | Returns or sets the number of points that are cropped off the right side of the specified picture. |
| `crop top` | `real` | r/w | Returns or sets the number of points that are cropped off the top of the specified picture. |
| `transparency color` | `cRGBColor` | r/w | Returns or sets the transparent color for the specified picture as a RGB color. For this property to take effect, the transparent background property must be set to true. |
| `transparent background` | `boolean` | r/w | Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent. |

### `picture`

Represents an picture object in the drawing layer.

*Plural:* `pictures`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `text` | r/o | The name of the file containing the picture. |
| `link to file` | `boolean` | r/o | Returns true if the picture shape is linked to the file. |
| `picture format` | `picture format` | r/o | Returns the picture format object associated with this picture shape. |
| `save with document` | `boolean` | r/o | Specifies if the picture should be saved with the document. |

### `reflection format`

Represents the reflection effect in Office graphics.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `reflection type` | `MsoReflectionType` | r/w | Returns or sets the type of the reflection format object. |

### `shadow format`

Represents shadow formatting for a shape.

*Plural:* `shadow formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `blur` | `real` | r/w | Returns or sets the blur, in points, of the specified shadow. |
| `fore color` | `integer` | r/w | Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow. |
| `fore color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the foreground color for the shadow format. |
| `obscured` | `boolean` | r/w | Returns or sets if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. If false the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill. |
| `offset x` | `real` | r/w | Returns or sets the horizontal offset in points of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. |
| `offset y` | `real` | r/w | Returns or sets the vertical offset in points of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape. |
| `rotate with shape` | `boolean` | r/w | Returns or sets whether to rotate the shadow when rotating the shape. |
| `shadow style` | `MsoShadowStyle` | r/w | Returns or sets the style of shadow formatting to apply to a shape. |
| `shadow type` | `MsoShadowType` | r/w | Returns or sets the shape shadow type. |
| `size` | `real` | r/w | Returns or sets the width of the shadow. |
| `transparency` | `real` | r/w | Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear. |
| `visible` | `boolean` | r/w | Returns or sets if the specified object, or the formatting applied to it, is visible. |

### `shape`

Represents an object in the drawing layer.

*Plural:* `shapes`

**Contains:** `adjustment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `anchor` | `text range` | r/o | Returns a text range object that represents the anchoring range for the specified shape. |
| `anchorID` | `integer` | r/o | Returns the anchor id for the specified shape. |
| `auto shape type` | `MsoAutoShapeType` | r/w | Returns or sets the shape type for the specified shape object, which must represent an autoshape. |
| `child` | `boolean` | r/o | True if the shape is a child shape. |
| `editID` | `integer` | r/o | Returns the edit id for the specified shape. |
| `fill format` | `fill format` | r/o | Return the fill format object associated with this shape object. |
| `glow format` | `glow format` | r/o | Returns the formatting properties for a glow effect. |
| `has chart` | `boolean` | r/w | True if the specified shape has a chart. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `horizontal flip` | `boolean` | r/o | Returns true if the shape has been flipped horizontally. |
| `hyperlink` | `hyperlink object` | r/o | Returns the hyperlink object associated with this shape object. |
| `left position` | `real` | r/w | Returns or sets the horizontal position of the object. |
| `line format` | `line format` | r/o | Returns the line format object associated with this shape object. |
| `link format` | `link format` | r/o | Returns the link format object associated with this shape object. |
| `lock anchor` | `boolean` | r/w | Returns or sets if the specified shape object's anchor is locked to the anchoring range. When a shape has a locked anchor, you cannot move the shape's anchor by dragging it, the anchor doesn't move as the shape is moved. |
| `lock aspect ratio` | `boolean` | r/w | Returns or sets if the specified shape retains its original proportions when you resize it. If false, you can change the height and width of the shape independently of one another when you resize it. |
| `name` | `text` | r/w | Returns or sets the name of this shape object. |
| `reflection format` | `reflection format` | r/o | Returns the formatting properties for a reflection effect. |
| `relative horizontal position` | `WdRelativeHorizontalPosition` | r/w | Returns or sets if the horizontal position of the shape is relative. |
| `relative vertical position` | `WdRelativeVerticalPosition` | r/w | Returns or sets if the vertical position of the shape is relative. |
| `rotation` | `real` | r/w | Returns or sets the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. |
| `shadow` | `shadow format` | r/o | Returns the shadow format object associated with this shape object. |
| `shape type` | `MsoShapeType` | r/o | Returns the shape type. |
| `soft edge format` | `soft edge format` | r/o | Returns the formatting properties for a soft edge effect. |
| `text frame` | `text frame` | r/o | Returns the text frame object associated with this shape object. |
| `threeD format` | `threeD format` | r/o | Returns the threeD format object associated with this shape object. |
| `top` | `real` | r/w | Returns or sets the vertical position of the object. |
| `vertical flip` | `boolean` | r/o | Returns true if the shape has been flipped vertically. |
| `visible` | `boolean` | r/w | Returns or sets if the shape object is visible. |
| `width` | `real` | r/w | Returns or sets the width of the object. |
| `wrap format` | `wrap format` | r/o | Returns the wrap format object associated with this shape object. |
| `z order position` | `integer` | r/o | Returns the position of the specified shape in the z-order. |

### `soft edge format`

Represents the soft edge formatting for a shape or range of shapes.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `soft edge type` | `MsoSoftEdgeType` | r/w | Returns or sets the type soft edge format object. |

### `standard inline horizontal line`

Represents a standard horizontal line object in the text layer of a document.

*Plural:* `standard inline horizontal lines`

### `text box`

Represents an text box object in the drawing layer.

*Plural:* `text boxes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `text orientation` | `MsoTextOrientation` | r/o | Returns the orientation of the text inside the text shape. |

### `text frame`

Represents the text frame in a shape object. Contains the text in the text frame as well as the properties that control the margins and orientation of the text frame.

*Plural:* `text frames`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `containing range` | `text range` | r/o | Returns a text range object that represents the entire story in a series of shapes with linked text frames that the specified text frame belongs to. |
| `has text` | `boolean` | r/o | Returns true if the specified shape has text associated with it. |
| `margin bottom` | `real` | r/w | Returns or sets the distance in points between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. |
| `margin left` | `real` | r/w | Returns or sets the distance in points between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. |
| `margin right` | `real` | r/w | Returns or sets the distance in points between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. |
| `margin top` | `real` | r/w | Returns or sets the distance in points between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. |
| `next textframe` | `text frame` | r/w | Returns or sets the next text frame object. |
| `orientation` | `MsoTextOrientation` | r/w | Returns or sets the orientation of the text inside the frame. |
| `overflowing` | `boolean` | r/o | Returns if the text inside the specified text frame doesn't all fit within the frame. |
| `previous textframe` | `text frame` | r/w | Returns or sets the previous text frame object. |
| `text range` | `text range` | r/o | Returns a text range object that represents the text range for this text frame. |

### `theme color scheme`

Represents the color scheme of a Microsoft Office 2007 theme.

**Contains:** `theme color`

### `theme color`

Represents a color in the color scheme of a Microsoft Office 2007 theme.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `RGB` | `integer` | r/w | Returns or sets a value of a color in the color scheme of a Microsoft Office theme. |
| `theme color scheme index` | `MsoThemeColorSchemeIndex` | r/o | Returns the index value a color scheme of a Microsoft Office theme. |

### `theme effect scheme`

Represents the effect scheme of a Microsoft Office theme.

### `theme font scheme`

Represents the font scheme of a Microsoft Office theme.

**Contains:** `minor theme font`, `major theme font`

### `theme font`

Represents a container for the font schemes of a Microsoft Office 2007 theme.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | Returns or sets a value specifying the font to use for a selection. |

### `theme fonts`

Represents a collection of major and minor fonts in the font scheme of a Microsoft Office 2007 theme.

### `threeD format`

Represents a shape's three-dimensional formatting.

*Plural:* `threeD formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `Z distance` | `real` | r/w | Returns or sets a Single that represents the distance from the center of an object or text. |
| `bevel bottom depth` | `real` | r/w | Returns or sets the depth/height of the bottom bevel. |
| `bevel bottom inset` | `real` | r/w | Returns or sets the inset size/width for the bottom bevel. |
| `bevel bottom type` | `MsoBevelType` | r/w | Returns or sets the bevel type for the bottom bevel. |
| `bevel top depth` | `real` | r/w | Returns or sets the depth/height of the top bevel. |
| `bevel top inset` | `real` | r/w | Returns or sets the inset size/width for the top bevel. |
| `bevel top type` | `MsoBevelType` | r/w | Returns or sets the bevel type for the top bevel. |
| `contour color` | `integer` | r/w | Returns or sets the color of the contour of an object or text. |
| `contour color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified contour. |
| `contour width` | `real` | r/w | Returns or sets the width of the contour of an object or text. |
| `depth` | `real` | r/w | Returns or sets the depth of the shape's extrusion. |
| `extrusion color` | `integer` | r/w | Returns or sets a RGB color that represents the color of the shape's extrusion. |
| `extrusion color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified extrusion. |
| `extrusion color type` | `MsoExtrusionColorType` | r/w | Returns or sets a value that indicates what will determine the extrusion color. |
| `field of view` | `real` | r/w | Returns or sets the amount of perspective for an object or text. |
| `light angle` | `real` | r/w | Returns or sets the angle of the lighting. |
| `perspective` | `boolean` | r/w | Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point. |
| `preset camera` | `MsoPresetCamera` | r/w | Returns a constant that represents the camera preset. |
| `preset extrusion direction` | `MsoPresetExtrusionDirection` | r/w | Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion. |
| `preset lighting direction` | `MsoPresetLightingDirection` | r/w | Returns or sets the position of the light source relative to the extrusion. |
| `preset lighting rig` | `MsoLightRigType` | r/w | Returns a constant that represents the lighting preset. |
| `preset lighting softness` | `MsoPresetLightingSoftness` | r/w | Returns or sets the intensity of the extrusion lighting. |
| `preset material` | `MsoPresetMaterial` | r/w | Returns or sets the extrusion surface material. |
| `preset threeD format` | `MsoPresetThreeDFormat` | r/w | Returns or sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion. |
| `project text` | `boolean` | r/w | Returns or sets whether text on a shape rotates with shape. |
| `rotation X` | `real` | r/w | Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation. |
| `rotation Y` | `real` | r/w | Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right. |
| `rotation Z` | `real` | r/w | Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation. |
| `visible` | `boolean` | r/w | Returns or sets if the specified object, or the formatting applied to it, is visible. |

### `word art format`

Contains properties and methods that apply to WordArt objects.

*Plural:* `word art formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `MsoTextEffectAlignment` | r/w | Returns or sets a constant that represents the alignment for the specified text effect. |
| `bold` | `boolean` | r/w | Returns or sets if the text of the word art shape is formatted as bold. |
| `font name` | `text` | r/w | Returns or sets the font name of the font used by this word art shape. |
| `font size` | `real` | r/w | Returns or sets the font size of the font used by this word art shape. |
| `italic` | `boolean` | r/w | Returns or sets if the text of the word art shape is formatted as italic. |
| `kerned pairs` | `boolean` | r/w | Returns or sets if character pairs in a WordArt object have been kerned. |
| `normalized height` | `boolean` | r/w | Returns or sets if all characters, both uppercase and lowercase, in the specified WordArt are the same height. |
| `preset shape` | `MsoPresetTextEffectShape` | r/w | Returns or sets the shape of the specified WordArt. |
| `preset word art effect` | `MsoPresetTextEffect` | r/w | Returns or sets the style of the specified WordArt. |
| `rotated chars` | `boolean` | r/w | Returns or sets if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. If false, characters in the specified WordArt retain their original orientation relative to the bounding shape. |
| `tracking` | `real` | r/w | Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Can be a value from zero through 5. |
| `word art text` | `text` | r/w | Returns or sets the text associated with this word art shape. |

### `word art`

Represents an word art object in the drawing layer.

*Plural:* `word arts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bold` | `boolean` | r/o | Returns if the text of the word art shape is formatted as bold. |
| `font name` | `text` | r/o | Returns the font name of the font used by this word art shape. |
| `font size` | `real` | r/o | Returns the font size of the font used by this word art shape. |
| `italic` | `boolean` | r/o | Returns if the text of the word art shape is formatted as italic. |
| `preset word art effect` | `MsoPresetTextEffect` | r/o | Returns the style of the specified word art. |
| `word art format` | `word art format` | r/o | Returns the word art format object associated with the word art shape object. |
| `word art text` | `text` | r/o | The text associated with this word art shape. |

### `wrap format`

Represents all the properties for wrapping text around a shape.

*Plural:* `wrap formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow overlap` | `boolean` | r/w | Returns or sets a value that specifies whether a given shape can overlap other shapes. |
| `distance bottom` | `real` | r/w | Returns or sets the distance in points between the document text and the bottom edge of the text-free area surrounding the specified shape. |
| `distance left` | `real` | r/w | Returns or sets the distance in points between the document text and the left edge of the text-free area surrounding the specified shape. |
| `distance right` | `real` | r/w | Returns or sets the distance in points between the document text and the right edge of the text-free area surrounding the specified shape. |
| `distance top` | `real` | r/w | Returns or sets the distance in points between the document text and the top edge of the text-free area surrounding the specified shape. |
| `wrap side` | `WdWrapSideType` | r/w | Returns or sets a value that indicates whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin. |
| `wrap type` | `WdWrapType` | r/w | Returns or sets the wrap type for the specified shape. |

### Enumerations

### `WdThemeColorIndex`

| Value | Description |
|-------|-------------|
| `not theme color` |  |
| `theme color main dark1` |  |
| `theme color main light1` |  |
| `theme color main dark2` |  |
| `theme color main light2` |  |
| `theme color accent1` |  |
| `theme color accent2` |  |
| `theme color accent3` |  |
| `theme color accent4` |  |
| `theme color accent5` |  |
| `theme color accent6` |  |
| `theme color hyperlink` |  |
| `theme color hyperlink followed` |  |
| `theme color background1` |  |
| `theme color text1` |  |
| `theme color background2` |  |
| `theme color text2` |  |

### `automatic length`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `callout format` |  |

### `custom drop`

| Value | Description |
|-------|-------------|
| `callout` |  |
| `callout format` |  |

### `custom length`

| Value | Description |
|-------|-------------|
| `callout` |  |
| `callout format` |  |

### `preset drop`

| Value | Description |
|-------|-------------|
| `callout` |  |
| `callout format` |  |

### `one color gradient`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `patterned`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `preset gradient`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `preset textured`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `solid`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `two color gradient`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `user picture`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `user textured`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `reset rotation`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `threeD format` |  |

### `save as picture`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `inline shape` |  |

---

## Text Suite

Classes and types for scripting text manipulations

### Commands

### `auto format text range`

Automatically formats a text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `calculate range`

Calculates a mathematical expression within a text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

**Returns:** `real`

### `change end of range`

Moves or extends the ending character position of a range or selection to the end of the nearest specified text unit. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `by` | `WdUnits` | yes | The unit by which to move the ending character position. The default value is a word item. |
| `extend type` | `WdMovementType` | yes | Determines if the text range is moved or extended. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `change start of range`

Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `by` | `WdUnits` | yes | The unit by which to move the starting character position. The default value is a word item. |
| `extend type` | `WdMovementType` | yes | Determines if the text range is moved or extended. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `check grammar`

Begins a grammar check for the text range.  Returns true if the text range had no errors

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

**Returns:** `boolean`

### `check spelling`

Begins a spelling check for the text range.  Returns true if the text range had no errors

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `custom dictionary` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `ignore uppercase` | `boolean` | yes | Set to true to ignore capitalization. |
| `main dictionary` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 2` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 3` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 4` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 5` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 6` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 7` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 8` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 9` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary 10` | `dictionary` | yes | A dictionary object that should be used to check the text range. |

**Returns:** `boolean`

### `check synonyms`

Displays the thesaurus dialog box, which lists alternative word choices, or synonyms, for the text in the text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `close up`

Removes any spacing before the specified paragraphs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4041` | no |  |

### `collapse range`

Collapses this text range to the starting or ending position and returns a new text range object. After a text range is collapsed, the starting and ending points are equal.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `direction` | `WdCollapseDirection` | yes | The direction to collapse the text range object. |

**Returns:** `text range` — A new text range object that refers to the text range specified by the direction argument.

### `compute text range statistics`

Returns a statistic based on the contents of the specified text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `statistic` | `WdStatistic` | no | The statistic to be computed. |

**Returns:** `integer`

### `convert to table`

Converts text within a text range to a table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `separator` | `WdTableFieldSeparator` | yes | Specifies the character used to separate text into cells. |
| `number of rows` | `integer` | yes | The number of rows in the table. If this argument is omitted, Microsoft Word sets the number of rows, based on the contents of the text range. |
| `number of columns` | `integer` | yes | The number of columns in the table. If this argument is omitted, Microsoft Word sets the number of rows, based on the contents of the text range. |
| `initial column width` | `integer` | yes | The initial width of each column, in points. If this argument is omitted, Word calculates and adjusts the column width so that the table stretches from margin to margin. |
| `table format` | `WdTableFormat` | yes | Specifies one of the predefined formats |
| `apply borders` | `boolean` | yes | Set to true to apply the border properties of the specified format. |
| `apply shading` | `boolean` | yes | Set to true to apply the shading properties of the specified format. |
| `apply font` | `boolean` | yes | Set to true to apply the font properties of the specified format. |
| `apply color` | `boolean` | yes | Set to true to apply the color properties of the specified format. |
| `apply heading rows` | `boolean` | yes | Set to true to apply the heading-row properties of the specified format. |
| `apply last row` | `boolean` | yes | Set to true to apply the last-row properties of the specified format. |
| `apply first column` | `boolean` | yes | Set to true to apply the first-column properties of the specified format. |
| `apply last column` | `boolean` | yes | Set to true to apply the last-column properties of the specified format. |
| `auto fit` | `boolean` | yes | Set to true to decrease the width of the table columns as much as possible without changing the way text wraps in the cells. |
| `auto fit behavior` | `WdAutoFitBehavior` | yes | Sets the auto fit rules for how Word sizes a table. |
| `default table behavior` | `WdDefaultTableBehavior` | yes | Sets a value that specifies whether Microsoft Word automatically resizes cells in a table to fit the contents. |

**Returns:** `table` — The newly created table object.

### `copy as picture`

Copies the content of the text range as a picture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `copy object`

Copies the content of the text range to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `cut object`

Removes the content of the text range from the document and places it on the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `expand`

Expands the specified range. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `by` | `WdUnits` | yes | The unit by which to expand the range. The default value is a word item. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `get border`

Returns the specified border object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4040` | no |  |
| `which border` | `WdBorderType` | no | This specifies which border object should be retrieved. |

**Returns:** `border`

### `get footer`

Returns a specific footer object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section` | no |  |
| `index` | `WdHeaderFooterIndex` | no | Specifies which footer to retrieve. |

**Returns:** `header footer`

### `get header`

Returns a specific header object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section` | no |  |
| `index` | `WdHeaderFooterIndex` | no | Specifies which header to retrieve. |

**Returns:** `header footer`

### `get range information`

Returns requested information about the text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `information type` | `WdInformation` | no | The information to be returned. |

**Returns:** `text` — The requested information.

### `go to next`

Returns a text range object that refers to the start position of the next item or location specified by the what argument.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `what` | `WdGoToItem` | no | The item to go to. |

**Returns:** `text range`

### `go to previous`

Returns a text range object that refers to the start position of the previous item or location specified by the what argument.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `what` | `WdGoToItem` | no | The item to go to. |

**Returns:** `text range`

### `in range`

Returns true if the text range to which the method is applied is contained in the range specified by the text range argument.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `text range` | `text range` | no | The range to which you want to compare. |

**Returns:** `boolean`

### `in story`

Returns true if the text range to which the method is applied is in the same story as the range specified by the text range argument.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `text range` | `text range` | no | The range whose story is compared with the story that contains the target range. |

**Returns:** `boolean`

### `indent`

Indents one or more paragraphs by one level.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `paragraph` | no |  |

### `indent char width`

Indents one or more paragraphs by a specified number of characters.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4050` | no |  |
| `count` | `integer` | no | The number of characters by which the specified paragraphs are to be indented. |

### `indent first line char width`

Indents the first line of one or more paragraphs by a specified number of characters.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4051` | no |  |
| `count` | `integer` | no | The number of characters by which the first line of each specified paragraph is to be indented. |

### `is equivalent`

True if the selection or range to which this method is applied is equal to the range specified by the text range argument. This method compares the starting and ending character positions, as well as the story type.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `text range` | `text range` | no | The text range object to compare with. |

**Returns:** `boolean`

### `link to list template`

Links the specified style to a list template so that the style's formatting can be applied to lists.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `Word style` | no |  |
| `list template` | `list template` | no | The list template that the style is to be linked to. |
| `list level number` | `integer` | yes | An integer corresponding to the list level that the style is to be linked to. If this argument is omitted, then the level of the style is used. |

### `modify enclosure`

Adds, modifies, or removes an enclosure around the specified character or characters.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `enclosure style` | `WdEncloseStyle` | no | The style of the enclosure. |
| `enclosure type` | `WdEnclosureType` | yes | The type of the enclosure. |
| `enclosed text` | `text` | yes | The characters that you want to enclose. If you include this argument, Microsoft Word replaces the text range with the enclosed characters. If you don't specify text to enclose, Microsoft Word encloses all text in the specified range. |

### `move end of range`

Moves the ending character position of the range.  This method returns the new text range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `by` | `WdUnits` | yes | The unit by which to move the ending character position. The default value is a character item. |
| `count` | `integer` | yes | The number of units to move. A positive constant specifies moving forward, a negative constant specifies moving backward. The default value is one. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `move range`

Collapses the specified range to its start or end position and then moves the collapsed object by the specified number of units. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `by` | `WdUnits` | yes | The unit by which to move the collapsed range. |
| `count` | `integer` | yes | The number of units by which you want to move. The default value is one. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `move range end until`

Moves the end position of the specified range until any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `characters` | `text` | no | The set of characters for which to scan. |
| `count` | `WdRecoveryType` | yes | The maximum number of character positions to move. A positive constant specifies moving forward, a negative constant specifies moving backward. The default value is go forward. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `move range end while`

Moves the ending character position of a the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `characters` | `text` | no | The set of characters for which to scan. |
| `count` | `WdRecoveryType` | yes | The maximum number of character positions to move. A positive constant specifies moving forward, a negative constant specifies moving backward. The default value is go forward. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `move range start until`

Moves the start position of the specified range until one of the specified characters is found in the document. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `characters` | `text` | no | The set of characters for which to scan. |
| `count` | `WdRecoveryType` | yes | The maximum number of character positions to move. A positive constant specifies moving forward, a negative constant specifies moving backward. The default value is go forward. |

**Returns:** `text range` — A new text range object that refers to the specified text range

### `move range start while`

Moves the start position of the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `characters` | `text` | no | The set of characters for which to scan. |
| `count` | `WdRecoveryType` | yes | The maximum number of character positions to move. A positive constant specifies moving forward, a negative constant specifies moving backward. The default value is go forward. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `move range until`

Moves the specified range until one of the specified characters is found in the document. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `characters` | `text` | no | The set of characters for which to scan. |
| `count` | `WdRecoveryType` | yes | The maximum number of character positions to move. A positive constant specifies moving forward, a negative constant specifies moving backward. The default value is go forward. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `move range while`

Moves the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `characters` | `text` | yes | The set of characters for which to scan. |
| `count` | `WdRecoveryType` | yes | The maximum number of character positions to move. A positive constant specifies moving forward, a negative constant specifies moving backward. The default value is go forward. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `move start of range`

Moves the starting character position of the range. This method returns the new text range object or missing value if an error occurred.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `by` | `WdUnits` | yes | The unit by which to move the starting character position. The default value is a character item. |
| `count` | `integer` | yes | The number of units to move. A positive constant specifies moving forward, a negative constant specifies moving backward. The default value is one |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `navigate`

Returns a text range object that represents the start position of the specified item, such as a page, bookmark, or field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `to` | `WdGoToItem` | yes | The type of item to navigate to. |
| `position` | `WdGoToDirection` | yes | The type of navigation that should be performed. |
| `count` | `integer` | yes | The number of items by which to navigate. The default value is one. |
| `name` | `text` | yes | The name of the item to navigate to. |

**Returns:** `text range` — A new text range that refers to the specified item

### `next paragraph`

Returns the next paragraph object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `paragraph` | no |  |

**Returns:** `paragraph`

### `next range`

Returns a text range object relative to the specified text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `by` | `WdUnits` | yes | Relative relationship of the next text range. The default value is a character item. |
| `count` | `integer` | yes | The number of units by which to move ahead. The default value is one. |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `next subdocument`

Returns a new text range object to the next subdocument. If there isn't another subdocument, an error occurs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

**Returns:** `text range`

### `open or close up`

If spacing before the specified paragraphs is zero, this method sets spacing to 12 points. If spacing before the paragraphs is greater than zero, this method sets spacing to zero.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4043` | no |  |

### `open up`

Sets spacing before the specified paragraphs to 12 points.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4042` | no |  |

### `outdent`

Removes one level of indent for one or more paragraphs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `paragraph` | no |  |

### `outline demote`

Applies the next heading level style Heading 1 through Heading 8 to the specified paragraph or paragraphs. For example, if a paragraph is formatted with the Heading 2 style, this method demotes the paragraph by changing the style to Heading 3.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `paragraph` | no |  |

### `outline demote to body`

Demotes the specified paragraph or paragraphs to body text by applying the Normal style.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `paragraph` | no |  |

### `outline promote`

Applies the previous heading level style Heading 1 through Heading 8 to the specified paragraph or paragraphs. For example, if a paragraph is formatted with the Heading 2 style, this method promotes the paragraph by changing the style to Heading 1.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `paragraph` | no |  |

### `paste and format`

Pastes the selected table cells and formats them as specified.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `type` | `WdRecoveryType` | no | The type of formatting to use when pasting the selected table cells. |

### `paste append table`

Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `paste as nested table`

Pastes a cell or group of cells as a nested table into the text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `paste excel table`

Pastes and formats a Microsoft Excel table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `linked to Excel` | `boolean` | no | Set to true to link the pasted table to the original Excel file so that changes made to the Excel file are reflected in Microsoft Word. |
| `word formatting` | `boolean` | no | Set to true to format the table using the formatting in the Word document. False formats the table according to the original Excel file. |
| `RTF` | `boolean` | no | Set to true to paste the Excel table using Rich Text Format. False pastes the Excel table as HTML. |

### `paste object`

Inserts the contents of the clipboard at the specified text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `paste special`

Inserts the contents of the clipboard. Unlike with the paste method, with paste special you can control the format of the pasted information and optionally establish a link to the source file - for example, a Microsoft Excel worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `link` | `boolean` | yes | Set to true to create a link to the source file of the clipboard contents. The default value is false. |
| `placement` | `WdOLEPlacement` | yes | The placement for the clipboard contents when they're inserted into the document. The default value is 'in line' |
| `display as icon` | `boolean` | yes | Set to true to display the link as an icon. The default value is false. |
| `data type` | `WdPasteDataType` | yes | A format for the clipboard contents when they're inserted into the document. |
| `icon label` | `text` | yes | If display as icon is true, this argument is the text that appears below the icon. |

### `phonetic guide`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `text` | `text` | no |  |
| `alignment type` | `WdPhoneticGuideAlignmentType` | no |  |
| `raise` | `integer` | no |  |
| `font size` | `integer` | no |  |
| `font name` | `text` | no |  |

### `previous paragraph`

Returns the previous paragraph object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `paragraph` | no |  |

**Returns:** `paragraph`

### `previous range`

Returns a text range object relative to the specified text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `by` | `WdUnits` | yes | Relative relationship of the previous text range. The default value is a character item. |
| `count` | `integer` | yes | The number of units by which to move back. The default value is one. |

**Returns:** `text range` — A new text range object that refers to the text range specified by the by and count arguments.

### `previous subdocument`

Returns a new text range object to the previous subdocument. If there isn't another subdocument, an error occurs.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

**Returns:** `text range`

### `relocate`

In outline view, moves the paragraphs within the text range after the next visible paragraph or before the previous visible paragraph. Body text moves with a heading only if the body text is collapsed in outline view or if it's part of the text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `direction` | `WdRelocate` | no | The direction of the move. |

### `reset`

Removes paragraph formatting that differs from the underlying style. For example, if you manually right align a paragraph and the underlying style has a different alignment, the reset method changes the alignment to match the style formatting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4046` | no |  |

### `set range`

Returns a text range object by using the specified starting and ending character positions.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `start` | `integer` | no | The starting character position |
| `end` | `integer` | no | The ending character position |

**Returns:** `text range` — A new text range object that refers to the specified text range.

### `sort`

Sorts the paragraphs in the specified text range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `exclude header` | `boolean` | yes | Set to true to exclude the first row or paragraph header from the sort operation. The default value is false. |
| `field number` | `integer` | yes | The first field to sort by. |
| `sort field type` | `WdSortFieldType` | yes | The type of the first field |
| `sort order` | `WdSortOrder` | yes | The sorting order to use when sorting the first field. |
| `field number two` | `integer` | yes | The second field to sort by. |
| `sort field type two` | `WdSortFieldType` | yes | The type of the second field |
| `sort order two` | `WdSortOrder` | yes | The sorting order to use when sorting the second field. |
| `field number three` | `integer` | yes | The third field to sort by. |
| `sort field type three` | `WdSortFieldType` | yes | The type of the third field |
| `sort order three` | `WdSortOrder` | yes | The sorting order to use when sorting the third field. |
| `sort column` | `boolean` | yes | True to sort only the column in the specified text range object. |
| `separator` | `WdSortSeparator` | yes | The type of field separator. The default value is 'sort separate by commas'. If the text range is in a table, this argument is ignored |
| `case sensitive` | `boolean` | yes | Set to true to sort with case sensitivity. The default value is false |
| `language ID` | `WdLanguageID` | yes | Specifies the sorting language. |

### `sort ascending`

Sorts paragraphs or table rows in ascending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `sort descending`

Sorts paragraphs or table rows in descending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `space 1`

Single-spaces the specified paragraphs. The exact spacing is determined by the font size of the largest characters in each paragraph.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4047` | no |  |

### `space 15`

Formats the specified paragraphs with 1.5-line spacing. The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4048` | no |  |

### `space 2`

Double-spaces the specified paragraphs. The exact spacing is determined by adding 12 points to the font size of the largest character in each paragraph.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4049` | no |  |

### `tab hanging indent`

Sets a hanging indent to a specified number of tab stops. Can be used to remove tab stops from a hanging indent if the value of the count argument is a negative number.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4044` | no |  |
| `count` | `integer` | no | If positive, the number of tabs stops to indent or, if negative, the number of tab stops to remove from the indent. |

### `tab indent`

Sets the left indent for the specified paragraphs to a specified number of tab stops. Can also be used to remove the indent if the value of the count argument is a negative number.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4045` | no |  |
| `count` | `integer` | no | If positive, the number of tabs stops to indent or, if negative, the number of tab stops to remove from the indent. |

### `text range spelling suggestions`

Returns a record that contains the spelling error type and the list of words suggested as replacements for the first word in the specified range

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `custom dictionary` | `dictionary` | yes | A dictionary object that represents the primary custom dictionary. |
| `ignore uppercase` | `boolean` | yes | Set to true to ignore words in all uppercase letters. If this argument is omitted, the current value of the ignore uppercase property of the Word options is used. |
| `main dictionary` | `dictionary` | yes | A dictionary object that represents the main dictionary. If this argument is omitted, Word uses the main dictionary that corresponds to the language formatting of the first word in the text range. |
| `suggestion mode` | `WdSpellingWordType` | yes | Specifies the way Word makes spelling suggestions.  The default is spelling word type spell word. |
| `custom dictionary2` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary3` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary4` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary5` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary6` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary7` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary8` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary9` | `dictionary` | yes | A dictionary object that should be used to check the text range. |
| `custom dictionary10` | `dictionary` | yes | A dictionary object that should be used to check the text range. |

**Returns:** `record` — The spelling error type, named 'type', and the list of words, named 'list'. The spelling error types are 'spelling correct', 'spelling not in dictionary', and 'spelling capitalization'.

### Classes

### `Word style`

Represents a single built-in or user-defined style. The Word style object includes style attributes, font, font style, paragraph spacing, and so on, as properties of the Word style object.

*Plural:* `Word styles`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `automatically update` | `boolean` | r/w | True if the style is automatically redefined based on the selection. False if Word prompts for confirmation before redefining the style based on the selection. |
| `base style` | `WdBuiltinStyle` | r/w | Returns or sets an existing style on which you can base the formatting of another style. |
| `border options` | `border options` | r/o | Returns back border options object associated with this cell object. |
| `built in` | `boolean` | r/o | Returns true if the specified object is one of the built-in styles in Word. |
| `description` | `text` | r/o | Returns the description of the specified style. For example, a typical description for the Heading 2 style might be Normal, Font: Arial, 12 pt, Bold, Italic, Space Before 12 pt After 3 pt, KeepWithNext, Level 2. |
| `font object` | `font` | r/o | Returns the font object associated with this Word style object. |
| `frame` | `frame` | r/o | Returns the frame object associated with this Word style object. |
| `in use` | `boolean` | r/o | Returns true if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document. |
| `language ID` | `WdLanguageID` | r/w | Returns or sets the language for the Word style object |
| `language ID east asian` | `WdLanguageID` | r/w | Returns or sets an east asian language for the template. |
| `list level number` | `integer` | r/o | Returns the list level for the specified style. |
| `list template` | `list template` | r/o | Returns the list template object associated with this Word style object. |
| `name local` | `text` | r/w | Returns or sets the name of a built-in style in the language of the user. Setting this property renames a user-defined style or adds an alias to a built-in style. |
| `next paragraph style` | `WdBuiltinStyle` | r/w | Returns or sets the style to be applied automatically to a new paragraph inserted after a paragraph formatted with the specified style. |
| `no proofing` | `boolean` | r/w | Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores for this Word style |
| `no space between same` | `boolean` | r/w | Returns or sets if Microsoft Word suppresses space between paragraphs of the same style |
| `paragraph format` | `paragraph format` | r/w | Returns or sets the paragraph format object associated with this Word style object. |
| `shading` | `shading` | r/o | Returns the shading object associated with the Word style. |
| `style type` | `WdStyleType` | r/o | Returns the style type. |
| `table style` | `table style` | r/o | Returns table style properties for this style. |

### `character`

*Plural:* `characters`

### `grammatical error`

*Plural:* `grammatical errors`

### `paragraph format`

Represents all the formatting for a paragraph.

*Plural:* `paragraph formats`

**Contains:** `tab stop`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `add space between east asian and alpha` | `boolean` | r/w | Returns or sets if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. |
| `add space between east asian and digit` | `boolean` | r/w | Returns or sets if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. |
| `alignment` | `WdParagraphAlignment` | r/w | Returns or sets a constant that represents the alignment for the specified paragraphs. |
| `auto adjust right indent` | `boolean` | r/w | Returns or sets if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if youÔøïve specified a set number of characters per line. |
| `base line alignment` | `WdBaselineAlignment` | r/w | Returns or sets a constant that represents the vertical position of fonts on a line. |
| `border options` | `border options` | r/o | Returns back border options object associated with this text range object. |
| `character unit first line indent` | `real` | r/w | Returns or sets the value in characters for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. |
| `character unit left indent` | `real` | r/w | Returns or sets the left indent value in characters for the specified paragraphs. |
| `character unit right indent` | `real` | r/w | Returns or sets the right indent value in characters for the specified paragraphs. |
| `disable line height grid` | `boolean` | r/w | Returns or sets if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. |
| `east asian line break control` | `boolean` | r/w |  |
| `first line indent` | `real` | r/w | Returns or sets the value in points for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. |
| `half width punctuation on top of line` | `boolean` | r/w | Returns or sets if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. |
| `hanging punctuation` | `boolean` | r/w | Returns or sets if hanging punctuation is enabled for the specified paragraphs. |
| `hyphenation` | `boolean` | r/w | Returns or sets if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation. |
| `keep together` | `boolean` | r/w | Returns or sets if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. |
| `keep with next` | `boolean` | r/w | Returns or sets if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. |
| `line spacing` | `real` | r/w | Returns or sets the line spacing in points for the specified paragraphs. |
| `line spacing rule` | `WdLineSpacing` | r/w | Returns or sets the line spacing for the specified paragraphs. |
| `line unit after` | `real` | r/w | Returns or sets the amount of spacing in gridlines after the specified paragraphs. |
| `line unit before` | `real` | r/w | Returns or sets the amount of spacing in gridlines before the specified paragraphs. |
| `no line number` | `boolean` | r/w | Returns or set if line numbers are repressed for the specified paragraphs. |
| `outline level` | `WdOutlineLevel` | r/w | Returns or sets the outline level for the specified paragraphs. |
| `page break before` | `boolean` | r/w | Returns or sets if a page break is forced before the specified paragraphs. |
| `paragraph format left indent` | `real` | r/w | Returns or sets the left indent in points for the specified paragraphs. |
| `paragraph format right indent` | `real` | r/w | Returns or sets the right indent in points for the specified paragraphs. |
| `shading` | `shading` | r/o | Returns the shading object associated with the paragraph format object. |
| `space after` | `real` | r/w | Returns or sets the spacing in points after the specified paragraphs. |
| `space after auto` | `boolean` | r/w | Returns or sets if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. |
| `space before` | `real` | r/w | Returns or sets the spacing in points before the specified paragraphs. |
| `space before auto` | `boolean` | r/w | Returns or sets if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. |
| `style` | `WdBuiltinStyle` | r/w | Returns or sets the Word style associated with the replacement object. |
| `widow control` | `boolean` | r/w | Returns or sets if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. |
| `word wrap` | `boolean` | r/w | Returns or sets if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames. |

### `paragraph`

Represents a single paragraph in a selection, range, or document.

*Plural:* `paragraphs`

**Contains:** `tab stop`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `add space between east asian and alpha` | `boolean` | r/w | Returns or sets if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. |
| `add space between east asian and digit` | `boolean` | r/w | Returns or sets if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. |
| `alignment` | `WdParagraphAlignment` | r/w | Returns or sets a constant that represents the alignment for the specified paragraphs. |
| `auto adjust right indent` | `boolean` | r/w | Returns or sets if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if youÔøïve specified a set number of characters per line. |
| `base line alignment` | `WdBaselineAlignment` | r/w | Returns or sets a constant that represents the vertical position of fonts on a line. |
| `border options` | `border options` | r/o | Returns back border options object associated with this text range object. |
| `character unit first line indent` | `real` | r/w | Returns or sets the value in characters for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. |
| `character unit left indent` | `real` | r/w | Returns or sets the left indent value in characters for the specified paragraphs. |
| `character unit right indent` | `real` | r/w | Returns or sets the right indent value in characters for the specified paragraphs. |
| `disable line height grid` | `boolean` | r/w | Returns or sets if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. |
| `drop cap` | `drop cap` | r/o | Returns the drop cap object associated with this paragraph object. |
| `east asian line break control` | `boolean` | r/w |  |
| `first line indent` | `real` | r/w | Returns or sets the value in points for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. |
| `half width punctuation on top of line` | `boolean` | r/w | Returns or sets if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. |
| `hanging punctuation` | `boolean` | r/w | Returns or sets if hanging punctuation is enabled for the specified paragraphs. |
| `hyphenation` | `boolean` | r/w | Returns or sets if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation. |
| `keep together` | `boolean` | r/w | Returns or sets if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. |
| `keep with next` | `boolean` | r/w | Returns or sets if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. |
| `line spacing` | `real` | r/w | Returns or sets the line spacing in points for the specified paragraphs. |
| `line spacing rule` | `WdLineSpacing` | r/w | Returns or sets the line spacing for the specified paragraphs. |
| `line unit after` | `real` | r/w | Returns or sets the amount of spacing in gridlines after the specified paragraphs. |
| `line unit before` | `real` | r/w | Returns or sets the amount of spacing in gridlines before the specified paragraphs. |
| `no line number` | `boolean` | r/w | Returns or set if line numbers are repressed for the specified paragraphs. |
| `outline level` | `WdOutlineLevel` | r/w | Returns or sets the outline level for the specified paragraphs. |
| `page break before` | `boolean` | r/w | Returns or sets if a page break is forced before the specified paragraphs. |
| `paragraph format` | `paragraph format` | r/w | Returns or sets the paragraph format object associated with this paragraph object. |
| `paragraph left indent` | `real` | r/w | Returns or sets the left indent in points for the specified paragraphs. |
| `paragraph_id` | `integer` | r/w | Returns the paragraph id for the specified paragraphs. |
| `right indent` | `real` | r/w | Returns or sets the right indent in points for the specified paragraphs. |
| `shading` | `shading` | r/o | Returns the shading object associated with the paragraph object. |
| `space after` | `real` | r/w | Returns or sets the spacing in points after the specified paragraphs. |
| `space before` | `real` | r/w | Returns or sets the spacing in points before the specified paragraphs. |
| `style` | `WdBuiltinStyle` | r/w | Returns or sets the Word style associated with the replacement object. |
| `text object` | `text range` | r/o | Returns a text range object that represents the portion of a document that's contained in the section object. |
| `text_id` | `integer` | r/w | Returns the text id for the specified paragraphs. |
| `widow control` | `boolean` | r/w | Returns or sets if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. |
| `word wrap` | `boolean` | r/w | Returns or sets if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames. |

### `section`

Represents a single section in a selection, range, or document.

*Plural:* `sections`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border options` | `border options` | r/o |  |
| `page setup` | `page setup` | r/w | Returns or sets a page setup object associated with this section object |
| `protected for forms` | `boolean` | r/w | Returns or sets if the specified section is protected for forms. When a section is protected for forms, you can select and modify text only in form fields. |
| `section index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `text object` | `text range` | r/o | Returns a text range object that represents the portion of a document that's contained in the section object. |

### `sentence`

*Plural:* `sentences`

### `shading`

Contains shading attributes for an object.

*Plural:* `shadings`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `background pattern color` | `integer` | r/w | Returns or sets the RGB color that's applied to the background of the shading object. |
| `background pattern color index` | `WdColorIndex` | r/w | Returns or sets the color index that's applied to the background of the shading object. |
| `background pattern color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color that's applied to the background of the shading object. |
| `foreground pattern color` | `integer` | r/w | Returns or sets the RGB color that's applied to the foreground of the shading object. |
| `foreground pattern color index` | `WdColorIndex` | r/w | Returns or sets the color index that's applied to the foreground of the shading object. |
| `foreground pattern color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color that's applied to the foreground of the shading object. This color is applied to the dots and lines in the shading pattern. |
| `texture` | `WdTextureIndex` | r/w | Returns or sets the shading texture for the specified object. |

### `spelling error`

*Plural:* `spelling errors`

### `story range`

*Plural:* `story ranges`

### `text range`

Represents a contiguous area in a document. Each text range object is defined by a starting and ending character position.

*Plural:* `ranges`

**Contains:** `character`, `word`, `sentence`, `table`, `footnote`, `endnote`, `Word comment`, `cell`, `section`, `paragraph`, `field`, `form field`, `frame`, `bookmark`, `revision`, `hyperlink object`, `subdocument`, `column`, `row`, `shape`, `readability statistic`, `grammatical error`, `spelling error`, `inline shape`, `math object`, `coauth lock`, `coauth update`, `conflict`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bold` | `boolean` | r/w | Returns or sets if the text associated with the text range is formatted as bold. |
| `bookmark id` | `integer` | r/o | Returns the number of the bookmark that encloses the beginning of the text range. The number corresponds to the position of the bookmark in the document, 1 for the first bookmark, 2 for the second one, and so on. |
| `border options` | `border options` | r/o | Returns back border options object associated with this text range object. |
| `character_case` | `WdCharacterCase` | r/w | Synonym for case |
| `case` | `WdCharacterCase` | r/w | Returns or sets a constant that represents the case of the text in the text range. |
| `column options` | `column options` | r/o |  |
| `combine characters` | `boolean` | r/w |  |
| `content` | `text` | r/w | Returns or sets the text in the text range. |
| `disable character space grid` | `boolean` | r/w | Returns or sets if Microsoft Word ignores the number of characters per line for the corresponding text range object. |
| `emphasis mark` | `WdEmphasisMark` | r/w | Returns or sets the emphasis mark for a character or designated character string. |
| `end of content` | `integer` | r/o | Returns the ending character position of the text range. |
| `endnote options` | `endnote options` | r/o | Returns the endnote options object for this text range. |
| `field options` | `field options` | r/o |  |
| `find object` | `find` | r/o | Returns the find object associated with this text range. |
| `fit text width` | `real` | r/w | Returns or sets the width in which Microsoft Word fits the text in the text range. |
| `font object` | `font` | r/o | Returns the font object associated with this text range. |
| `footnote options` | `footnote options` | r/o | Returns the footnote options object for this text range. |
| `formatted text` | `text range` | r/w | Returns or sets a text range object that includes the formatted text in the text range. |
| `grammar checked` | `boolean` | r/w | True if a grammar check has been run on the text range. False if some of the text range hasn't been checked for grammar. To recheck the grammar in the document, set the grammar checked property to false. |
| `highlight color index` | `WdColorIndex` | r/w | Returns or sets the highlight color for the text range. |
| `horizontal in vertical` | `WdHorizontalInVerticalType` | r/w |  |
| `is end of row mark` | `boolean` | r/o | Returns true if the text range is collapsed and is located at the end-of-row mark in a table. |
| `italic` | `boolean` | r/w | Returns or sets if the text associated with the text range is formatted as italic. |
| `language ID` | `WdLanguageID` | r/w | Returns or sets the language for the text range object |
| `language ID east asian` | `WdLanguageID` | r/w | Returns or sets an east asian language for the template. |
| `list format` | `list format` | r/o | Returns the list format object associated with this text range. |
| `next story range` | `text range` | r/o | Returns a text range object that refers to the next story |
| `no proofing` | `boolean` | r/w | Returns or sets if the spelling and grammar checker should ignore documents based on this text range. |
| `orientation` | `WdTextOrientation` | r/w | Returns or sets the orientation of text in a text range when the text direction feature is enabled. |
| `page setup` | `page setup` | r/w | Returns or sets the page setup object associated with this text range. |
| `paragraph format` | `paragraph format` | r/w | Returns or sets the paragraph format object associated with this text range. |
| `previous bookmark id` | `integer` | r/o | Returns the number of the last bookmark that starts before or at the same place as the text range, It returns zero if there's no corresponding bookmark. |
| `range endnote options` | `range endnote options` | r/o |  |
| `range footnote options` | `range footnote options` | r/o |  |
| `row options` | `row options` | r/o |  |
| `shading` | `shading` | r/o | Returns the shading object associated with this text range. |
| `show Word comments by` | `text` | r/w | Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'. |
| `show hidden bookmarks` | `boolean` | r/w | Returns or sets if hidden bookmarks are included in the elements of this text range. |
| `spelling checked` | `boolean` | r/w | True if a spelling check has been run on the text range. False if some of the text range hasn't been checked for spelling. To see if the document contains spelling errors, get the count of spelling errors for the text range. |
| `start of content` | `integer` | r/o | Returns the starting character position of the text range. |
| `story length` | `integer` | r/o | Returns the number of characters in the story that contains the text range. |
| `story type` | `WdStoryType` | r/o | Returns the story type for the text range. |
| `style` | `WdBuiltinStyle` | r/w | Returns or sets the Word style for this range. |
| `subdocuments expanded` | `boolean` | r/w | Returns or sets if the subdocuments in the specified text range are expanded. |
| `supplemental language ID` | `WdLanguageID` | r/w | Returns or sets the language for the text range object |
| `text retrieval mode` | `text retrieval mode` | r/w | Returns or sets the text retrieval object that controls how text is retrieved from this text range. |
| `two lines in one` | `WdTwoLinesInOneType` | r/w | Returns or sets whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.  Read/write |
| `underline` | `WdUnderline` | r/w | Returns or sets the type of underline applied to the text range. |

### `word`

*Plural:* `words`

### Enumerations

### `reset`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `get border`

| Value | Description |
|-------|-------------|
| `text range` |  |
| `section` |  |
| `paragraph` |  |
| `paragraph format` |  |
| `Word style` |  |

### `close up`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `indent char width`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `indent first line char width`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `open or close up`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `open up`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `space 15`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `space 1`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `space 2`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `tab hanging indent`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

### `tab indent`

| Value | Description |
|-------|-------------|
| `paragraph` |  |
| `paragraph format` |  |

---

## Proofing Suite

Classes and types for scripting text proofing

### Commands

### `apply correction`

Replaces a range with the value of the specified autocorrect entry.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `autocorrect entry` | no |  |
| `to range` | `text range` | no | A reference to a text range object that will be replaced with this autocorrect entry. |

### `get synonym list for`

Get the list of synonyms for a particular word

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `synonym info` | no |  |
| `item to check` | `text` | no |  |

**Returns:** `text` — The list contains a list of strings

### `get synonym list from`

Get the list of synonyms using an index into the list of meanings

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `synonym info` | no |  |
| `meaning index` | `integer` | no |  |

**Returns:** `text` — A list of strings containing the synonyms.

### Classes

### `autocorrect entry`

Represents a single autocorrect entry.

*Plural:* `autocorrect entries`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `autocorrect value` | `text` | r/w | Returns or sets the value of the auto correct entry. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/w | Returns or sets the name of the auto correct entry. |
| `rich text` | `boolean` | r/o | Returns if formatting is stored with the autocorrect entry replacement text. |

### `autocorrect`

Represents the autocorrect functionality in Word.

**Contains:** `autocorrect entry`, `first letter exception`, `two initial caps exception`, `other corrections exception`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `correct days` | `boolean` | r/w | Returns or sets if Word automatically capitalizes the first letter of days of the week. |
| `correct initial caps` | `boolean` | r/w | Returns or sets if Word automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, WOrd is corrected to Word. |
| `correct sentence caps` | `boolean` | r/w | Returns or sets if Word automatically capitalizes the first letter in each sentence. |
| `correct table caps` | `boolean` | r/w | Returns or sets if Word automatically capitalizes the first letter in each table cell. |
| `first letter auto add` | `boolean` | r/w | Returns or sets if Word automatically adds abbreviations to the list of autocorrect first letter exceptions. |
| `other corrections auto add` | `boolean` | r/w | Returns or sets if Microsoft Word automatically adds words to the list of other autocorrect exceptions. Word adds a word to this list if you delete and then retype a word that you didn't want Word to correct. |
| `replace text` | `boolean` | r/w | Returns or sets if Microsoft Word automatically replaces specified text with entries from the autocorrect list. |
| `replace text from spelling checker` | `boolean` | r/w | Returns or sets if Microsoft Word automatically replaces misspelled text with suggestions from the spelling checker as the user types. |
| `show autocorrect smart button` | `boolean` | r/w | Returns or sets if Word shows the AutoCorrect smart button which allows you to review the correction when an automatic correction occurs. |
| `two initial caps auto add` | `boolean` | r/w | Returns or sets if Microsoft Word automatically adds words to the list of autocorrect initial caps exceptions. |

### `dictionary`

Represents a dictionary.

*Plural:* `dictionaries`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `dictionary type` | `WdDictionaryType` | r/o | Returns the dictionary type. |
| `language ID` | `WdLanguageID` | r/w | Returns or sets the language for the dictionary object |
| `language specific` | `boolean` | r/w | Returns or sets if the custom dictionary is to be used only with text formatted for a specific language. |
| `name` | `text` | r/o | Returns or sets the name of this dictionary object. |
| `path` | `text` | r/o | Returns the disk or Web path to the specified dictionary in HFS path style. |
| `read only` | `boolean` | r/o | Returns true if the specified dictionary cannot be changed. |

### `first letter exception`

Represents an abbreviation excluded from automatic correction.

*Plural:* `first letter exceptions`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | The name of this first letter exception object. |

### `language`

Represents a language used for proofing or formatting in Microsoft Word.

*Plural:* `languages`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active grammar dictionary` | `dictionary` | r/o | Returns a dictionary object that represents the active grammar dictionary for the specified language |
| `active hyphenation dictionary` | `dictionary` | r/o | Returns a dictionary object that represents the active hyphenation dictionary for the specified language. |
| `active spelling dictionary` | `dictionary` | r/o | Returns a dictionary object that represents the active spelling dictionary for the specified language. |
| `active thesaurus dictionary` | `dictionary` | r/o | Returns a dictionary object that represents the active thesaurus dictionary for the specified language. |
| `default writing style` | `text` | r/w | Returns or sets the default writing style used by the grammar checker for the specified language. The name of the writing style is the localized name for the specified language. |
| `language ID` | `WdLanguageID` | r/o | Returns an enumeration that identifies the specified language. |
| `name` | `text` | r/o | Returns the name of the language |
| `name local` | `text` | r/o | Returns the name of a proofing tool language in the language of the user. |
| `spelling dictionary type` | `WdDictionaryType` | r/w | Returns or sets the proofing tool type |
| `writing style list` | `text` | r/o | Returns a list of strings that contains the names of all writing styles available for the specified language. |

### `other corrections exception`

Represents a single auto correct exception.

*Plural:* `other corrections exceptions`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | Returns the name of this other corrections exception object. |

### `readability statistic`

Represents one of the readability statistics for a document or range.

*Plural:* `readability statistics`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of this readability statistic object. |
| `readability value` | `real` | r/o | The value of this readability statistic object. |

### `synonym info`

Represents the information about synonyms, antonyms, related words, or related expressions for the specified range or a given string.

*Plural:* `synonym infos`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `antonyms` | `text` | r/o | Returns a list of antonyms for the word or phrase. |
| `context` | `text` | r/o | Returns the word or phrase that was looked up by the thesaurus. |
| `found` | `boolean` | r/o | Returns true if the thesaurus finds synonyms, antonyms, related words, or related expressions for the word or phrase. |
| `meaning count` | `integer` | r/o | Returns the number of entries in the list of meanings found in the thesaurus for the word or phrase. Returns zero if no meanings were found. |
| `meanings` | `text` | r/o | Returns the list of meanings for the word or phrase. |
| `part of speech` | `WdPartOfSpeech` | r/o | Returns a list of the parts of speech corresponding to the meanings found for the word or phrase looked up in the thesaurus. |
| `related expressions` | `text` | r/o | Returns a list of expressions related to the specified word or phrase. |
| `related words` | `text` | r/o | Returns a list of words related to the specified word or phrase. |

### `two initial caps exception`

Represents a single initial-capital autocorrect exception.

*Plural:* `two initial caps exceptions`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | The name of this two initial caps exception object. |

---

## Table Suite

Classes and types for scripting table manipulations

### Commands

### `auto fit`

Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4060` | no |  |

### `auto fit behavior`

Determines how Microsoft Word resizes a table when the autofit feature is used. Word can resize the table based on the content of the table cells or the width of the document window.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `table` | no |  |
| `behavior` | `WdAutoFitBehavior` | no | The auto fit behavior to set. |

### `auto format table`

Applies a predefined look to a table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `table` | no |  |
| `table format` | `WdTableFormat` | yes | Specifies one of the predefined formats |
| `apply borders` | `boolean` | yes | Set to true to apply the border properties of the specified format. |
| `apply shading` | `boolean` | yes | Set to true to apply the shading properties of the specified format. |
| `apply font` | `boolean` | yes | Set to true to apply the font properties of the specified format. |
| `apply color` | `boolean` | yes | Set to true to apply the color properties of the specified format. The default value is false. |
| `apply heading rows` | `boolean` | yes | Set to true to apply the heading-row properties of the specified format. The default value is true. |
| `apply last row` | `boolean` | yes | Set to true to apply the last-row properties of the specified format. |
| `apply first column` | `boolean` | yes | Set to true to apply the first-column properties of the specified format. |
| `apply last column` | `boolean` | yes | Set to true to apply the last-column properties of the specified format. |
| `auto fit` | `boolean` | yes | Set to true to decrease the width of the table columns as much as possible without changing the way text wraps in the cells. |

### `auto sum`

Inserts an = Formula field that calculates and displays the sum of the values in table cells above or to the left of the cell specified in the expression.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cell` | no |  |

### `convert row to text`

Converts a row to text and returns a text range object that represents the delimited text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4059` | no |  |
| `separator` | `WdTableFieldSeparator` | yes | The character that delimits the converted, columns paragraph marks delimit the converted rows. |
| `nested tables` | `boolean` | yes | Set to true if nested tables are converted to text. |

**Returns:** `text range` — A reference to the newly created text range object.

### `convert to text`

Converts a table to text and returns a text range object that represents the delimited text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `table` | no |  |
| `separator` | `WdTableFieldSeparator` | yes | The character that delimits the converted, columns paragraph marks delimit the converted rows. |
| `nested tables` | `boolean` | yes | Set to true if nested tables are converted to text. |

**Returns:** `text range` — A reference to the newly created text range object.

### `distribute row height`

Adjusts the height of the specified rows or cells so that they're equal.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `row options` | no |  |

### `distribute width`

Adjusts the width of the specified columns or cells so that they're equal.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `column options` | no |  |

### `formula`

Inserts an = Formula field that contains the specified formula into a table cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cell` | no |  |
| `formula string` | `text` | yes | The mathematical formula you want the = Formula field to evaluate. Spreadsheet-type references to table cells are valid. |
| `number format string` | `text` | yes | A format for the result of the = Formula field. |

### `get border`

Returns the specified border object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4052` | no |  |
| `which border` | `WdBorderType` | no |  |

**Returns:** `border`

### `get cell from table`

Returns a cell object that represents a cell in a table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `table` | no |  |
| `row` | `integer` | no | The number of the row in the table to return. Can be an integer between 1 and the number of rows in the table. |
| `column` | `integer` | no | The number of the cell in the table to return. Can be an integer between 1 and the number of columns in the table. |

**Returns:** `cell`

### `merge cell`

Merges the specified table cell with another cell. The result is a single table cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cell` | no |  |
| `with` | `cell` | no | A reference to the newly merged cell object. |

### `set left indent`

Sets the indentation for a row or rows in a table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4058` | no |  |
| `left indent` | `real` | no | The distance in points between the current left edge of the specified row or rows and the desired left edge |
| `ruler style` | `WdRulerStyle` | no | Controls the way Word adjusts the table when the left indent is changed. |

### `set table item height`

Sets the height of table rows.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4057` | no |  |
| `row height` | `integer` | no | The height of the row or rows, in points. |
| `height rule` | `WdRowHeightRule` | no | The rule for determining the height of the specified rows. |

### `set table item width`

Sets the width of columns or cells in a table

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4056` | no |  |
| `column width` | `real` | no | The width of the specified column or columns, in points. |
| `ruler style` | `WdRulerStyle` | no | Controls the way Word adjusts cell widths. |

### `sort ascending`

Sorts paragraphs or table rows in ascending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4054` | no |  |

### `sort descending`

Sorts paragraphs or table rows in descending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4055` | no |  |

### `split cell`

Splits a single table cell into multiple cells.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cell` | no |  |
| `number of rows` | `integer` | yes | The number of rows that the cell or group of cells is to be split into. |
| `number of columns` | `integer` | yes | The number of columns that the cell or group of cells is to be split into. |

### `split table`

Inserts an empty paragraph immediately above the specified row in the table, and returns a table object that contains both the specified row and the rows that follow it.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `table` | no |  |
| `row` | `integer` | no | The row that the table is to be split before. |

**Returns:** `table` — A reference to the newly split out table.

### `table sort`

Sort a table object.  For the column object only the first field is used

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4053` | no |  |
| `exclude header` | `boolean` | yes | Set to true to exclude row with the header format property set from the sort operation. The default value is false. |
| `field number` | `integer` | yes | The first field to sort by. |
| `sort field type` | `WdSortFieldType` | yes | The type of the first field |
| `sort order` | `WdSortOrder` | yes | The sorting order to use when sorting the first field. |
| `field number two` | `integer` | yes | The second field to sort by. |
| `sort field type two` | `WdSortFieldType` | yes | The type of the second field |
| `sort order two` | `WdSortOrder` | yes | The sorting order to use when sorting the second field. |
| `field number three` | `integer` | yes | The third field to sort by. |
| `sort field type three` | `WdSortFieldType` | yes | The type of the third field |
| `sort order three` | `WdSortOrder` | yes | The sorting order to use when sorting the third field. |
| `sort column` | `boolean` | yes | True to sort only the column specified by the table object. |
| `separator` | `WdSortSeparator` | yes | The type of field separator. |
| `case sensitive` | `boolean` | yes | Set to true to sort with case sensitivity. |
| `language ID` | `WdLanguageID` | yes | Specifies the sorting language. |

### `update auto format`

Updates the table with the characteristics of a predefined table format. For example, if you apply a table format with auto format and then insert rows and columns, the table may no longer match the predefined look.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `table` | no |  |

### Classes

### `cell`

Represents a single table cell.

*Plural:* `cells`

**Contains:** `table`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border options` | `border options` | r/o | Returns back border options object associated with this cell object. |
| `bottom padding` | `real` | r/w | Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table. |
| `column` | `column` | r/o | Returns the column object that contains this cell object. |
| `column index` | `integer` | r/o | Returns the number of the column that contains the specified cell. |
| `fit text` | `boolean` | r/w | Returns or sets if Microsoft Word visually reduces the size of text typed into a cell so that it fits within the column width. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `height rule` | `WdRowHeightRule` | r/w | Returns or sets the rule for determining the height of the specified cells. |
| `left padding` | `real` | r/w | Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table. |
| `nesting level` | `integer` | r/o | Returns the nesting level of the specified cell. |
| `next cell` | `cell` | r/o | Returns the next cell object |
| `preferred width` | `real` | r/w | Returns or sets the preferred width in points for the specified cell. |
| `preferred width type` | `WdPreferredWidthType` | r/w | Returns or sets the preferred unit of measurement to use for the width of the specified column. |
| `previous cell` | `cell` | r/o | Returns the previous cell object |
| `right padding` | `real` | r/w | Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table. |
| `row` | `row` | r/o | Returns the row object that contains this cell object. |
| `row index` | `integer` | r/o | Returns the number of the row that contains the specified cell. |
| `shading` | `shading` | r/o | Returns the shading object associated with the cell object. |
| `text object` | `text range` | r/o | Returns a text range object that represents the portion of a document that's contained in the cell object. |
| `top padding` | `real` | r/w | Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table. |
| `vertical alignment` | `WdCellVerticalAlignment` | r/w | Returns or sets the vertical alignment of text in one or more cells of a table. |
| `width` | `real` | r/w | Returns or sets the width of the object. |
| `word wrap` | `boolean` | r/w | Returns or set  if Microsoft Word wraps text to multiple lines and lengthens the cell so that the cell width remains the same. |

### `column options`

Represents options that can be set for columns.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border options` | `border options` | r/o | Returns back border options object associated with this cell object. |
| `default width` | `real` | r/w | Returns or sets the default width of columns. |
| `nesting level` | `integer` | r/o |  |
| `preferred width` | `real` | r/w | Returns or sets the preferred width in points for the specified columns. |
| `preferred width type` | `WdPreferredWidthType` | r/w | Returns or sets the preferred unit of measurement to use for the width of the specified columns. |
| `shading` | `shading` | r/o | Returns the shading object associated with columns. |

### `column`

Represents a single table column.

*Plural:* `columns`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border options` | `border options` | r/o | Returns back border options object associated with this column object. |
| `column index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `is first` | `boolean` | r/o | Returns if the specified column is the first one in the table. |
| `is last` | `boolean` | r/o | Returns if the specified column is the last one in the table. |
| `nesting level` | `integer` | r/o | Returns the nesting level of the specified column. |
| `next column` | `column` | r/o | Returns the next column object |
| `preferred width` | `real` | r/w | Returns or sets the preferred width in points for the specified column. |
| `preferred width type` | `WdPreferredWidthType` | r/w | Returns or sets the preferred unit of measurement to use for the width of the specified column. |
| `previous column` | `column` | r/o | Returns the previous column object |
| `shading` | `shading` | r/o | Returns the shading object associated with the column object. |
| `width` | `real` | r/w | Returns or sets the width of the object. |

### `condition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border options` | `border options` | r/o | Returns the border options object associated with this condition. |
| `bottom padding` | `real` | r/w |  |
| `font object` | `font` | r/o |  |
| `left padding` | `real` | r/w |  |
| `paragraph format` | `paragraph format` | r/w |  |
| `right padding` | `real` | r/w |  |
| `shading` | `shading` | r/o | Returns the shading object associated with this condition. |
| `top padding` | `real` | r/w |  |

### `row options`

Represents options that can be set for rows.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `WdRowAlignment` | r/w | Returns or sets a constant that represents the alignment for rows. |
| `allow break across pages` | `boolean` | r/w | Returns or sets if the text in a table row or rows are allowed to split across a page break. |
| `allow overlap` | `boolean` | r/w | Returns or sets a value that specifies whether the specified rows can overlap other rows. |
| `border options` | `border options` | r/o | Returns back border options object associated with this cell object. |
| `distance bottom` | `real` | r/w | Returns or sets the distance in points between the document text and the bottom edge of the specified table. This property doesn't have any effect if wrap around text is false. |
| `distance left` | `real` | r/w | Returns or sets the distance in points between the document text and the left edge of the specified table. This property doesn't have any effect if wrap around text is false. |
| `distance right` | `real` | r/w | Returns or sets the distance in points between the document text and the right edge of the specified table. This property doesn't have any effect if wrap around text is false. |
| `distance top` | `real` | r/w | Returns or sets the distance in points between the document text and the top edge of the specified table. This property doesn't have any effect if wrap around text is false. |
| `heading format` | `boolean` | r/w | Returns or sets if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `height rule` | `WdRowHeightRule` | r/w | Returns or sets the rule for determining the height of the specified rows. |
| `horizontal position` | `real` | r/w | Returns or sets the horizontal distance between the edge of the rows. |
| `nesting level` | `integer` | r/o |  |
| `relative horizontal position` | `WdRelativeHorizontalPosition` | r/w | Specifies to what the horizontal position of a group of rows is relative. |
| `relative vertical position` | `WdRelativeVerticalPosition` | r/w | Specifies to what the vertical position of a group of rows is relative. |
| `row left indent` | `real` | r/w | Returns or sets the left indent in points for the specified rows. |
| `shading` | `shading` | r/o | Returns the shading object associated with rows. |
| `space between columns` | `real` | r/w | Returns or sets the distance in points between text in adjacent columns of the specified row or rows. |
| `vertical position` | `real` | r/w | Returns or sets the vertical distance between the edge of the rows. |
| `wrap around text` | `boolean` | r/w | Returns or sets whether text should wrap around the specified rows. |

### `row`

Represents a row in a table.

*Plural:* `rows`

**Contains:** `cell`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `WdRowAlignment` | r/w | Returns or sets a constant that represents the alignment for the specified row. |
| `allow break across pages` | `boolean` | r/w | Returns or sets if the text in a row or rows are allowed to split across a page break. |
| `border options` | `border options` | r/o | Returns back border options object associated with this table object. |
| `heading format` | `boolean` | r/w | Returns or sets if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `height rule` | `WdRowHeightRule` | r/w | Returns or sets the rule for determining the height of the specified rows. |
| `is first` | `boolean` | r/o | Returns if the specified row is the first one in the table. |
| `is last` | `boolean` | r/o | Returns if the specified row is the last one in the table. |
| `nesting level` | `integer` | r/o | Returns the nesting level of the specified row. |
| `next row` | `row` | r/o | Returns the next row object |
| `previous row` | `row` | r/o | Returns the previous row object |
| `row index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `row left indent` | `real` | r/w | Returns or sets the left indent in points for the specified rows. |
| `shading` | `shading` | r/o | Returns the shading object associated with the row object. |
| `space between columns` | `real` | r/w | Returns or sets the distance in points between text in adjacent columns of the specified row or rows. |
| `text object` | `text range` | r/o | Returns a text range object that represents the portion of a document that's contained in the row object. |

### `table style`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `WdRowAlignment` | r/w | Returns or sets a constant that represents the alignment for the specified row. |
| `allow break accross page` | `boolean` | r/w |  |
| `border options` | `border options` | r/o | Returns back border options object associated with this table object. |
| `bottom left cell condition` | `condition` | r/o | Returns the conditional style for the bottom left cell. |
| `bottom padding` | `real` | r/w | Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table. |
| `bottom right cell condition` | `condition` | r/o | Returns the conditional style for the bottom right cell. |
| `column stripe` | `integer` | r/w |  |
| `even column condition` | `condition` | r/o | Returns the conditional style for even column bands. |
| `even row condition` | `condition` | r/o | Returns the conditional style for even row bands. |
| `first column condition` | `condition` | r/o | Returns the conditional style for the first column. |
| `first row condition` | `condition` | r/o | Returns the conditional style for the first row. |
| `last column condition` | `condition` | r/o | Returns the conditional style for the last column. |
| `last row condition` | `condition` | r/o | Returns the conditional style for the last row. |
| `left padding` | `real` | r/w | Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table. |
| `odd column condition` | `condition` | r/o | Returns the conditional style for odd column bands. |
| `odd row condition` | `condition` | r/o | Returns the conditional style for odd row bands. |
| `right padding` | `real` | r/w | Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table. |
| `row left indent` | `real` | r/w | Returns or sets the left indent in points for this table style. |
| `row stripe` | `integer` | r/w |  |
| `shading` | `shading` | r/o | Returns the shading object associated with this table style. |
| `spacing` | `real` | r/w | Returns or sets the spacing between the cells in a table. |
| `top left cell condition` | `condition` | r/o | Returns the conditional style for the top left cell. |
| `top padding` | `real` | r/w | Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table. |
| `top right cell condition` | `condition` | r/o | Returns the conditional style for the top right cell. |

### `table`

Represents a single table.

*Plural:* `tables`

**Contains:** `column`, `row`, `table`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow auto fit` | `boolean` | r/w | Returns or sets if Microsoft Word will automatically resize cells in a table to fit their contents. |
| `allow page breaks` | `boolean` | r/w | Returns or sets if Microsoft Word allows to break the specified table across pages. |
| `apply style first column` | `boolean` | r/w |  |
| `apply style heading rows` | `boolean` | r/w |  |
| `apply style last column` | `boolean` | r/w |  |
| `apply style last row` | `boolean` | r/w |  |
| `auto format type` | `WdTableFormat` | r/o | Returns the type of automatic formatting that's been applied to the specified table. |
| `border options` | `border options` | r/o | Returns back border options object associated with this table object. |
| `bottom padding` | `real` | r/w | Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table. |
| `column options` | `column options` | r/o | Returns the column options object associated with this table object. |
| `left padding` | `real` | r/w | Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table. |
| `nesting level` | `integer` | r/o | Returns the nesting level of the specified table. |
| `number of columns` | `integer` | r/o | Returns the number of columns in this table |
| `number of rows` | `integer` | r/o | Returns the number of rows in this table |
| `preferred width` | `real` | r/w | Returns or sets the preferred width in points for the specified table. |
| `preferred width type` | `WdPreferredWidthType` | r/w | Returns or sets the preferred unit of measurement to use for the width of the specified table. |
| `right padding` | `real` | r/w | Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table. |
| `row options` | `row options` | r/o | Returns the row options object associated with this table object. |
| `shading` | `shading` | r/o | Returns the shading object associated with the table object. |
| `spacing` | `real` | r/w | Returns or sets the spacing between the cells in a table. |
| `style` | `WdBuiltinStyle` | r/w | Returns or sets the Word style associated with the table object. |
| `text object` | `text range` | r/o | Returns a text range object that represents the portion of a document that's contained in the table object. |
| `top padding` | `real` | r/w | Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table. |
| `uniform` | `boolean` | r/o | Returns if all the rows in a table have the same number of columns. |

### Enumerations

### `set left indent`

| Value | Description |
|-------|-------------|
| `row` |  |
| `row options` |  |

### `convert row to text`

| Value | Description |
|-------|-------------|
| `row` |  |
| `row options` |  |

### `auto fit`

| Value | Description |
|-------|-------------|
| `column` |  |
| `column options` |  |

### `table sort`

| Value | Description |
|-------|-------------|
| `table` |  |
| `column` |  |

### `get border`

| Value | Description |
|-------|-------------|
| `table` |  |
| `row` |  |
| `column` |  |
| `cell` |  |
| `row options` |  |
| `column options` |  |
| `table style` |  |
| `condition` |  |

### `set table item height`

| Value | Description |
|-------|-------------|
| `row` |  |
| `cell` |  |
| `row options` |  |

### `set table item width`

| Value | Description |
|-------|-------------|
| `column` |  |
| `cell` |  |
| `column options` |  |

### `sort ascending`

| Value | Description |
|-------|-------------|
| `table` |  |
| `column` |  |

### `sort descending`

| Value | Description |
|-------|-------------|
| `table` |  |
| `column` |  |
