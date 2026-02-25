# Excel AppleScript Dictionary

> Auto-generated from `Excel.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Excel"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Microsoft Office Suite](#microsoft-office-suite)
- [Microsoft Excel Suite](#microsoft-excel-suite)
- [Drawing Suite](#drawing-suite)
- [Text Suite](#text-suite)
- [Table Suite](#table-suite)
- [Proofing Suite](#proofing-suite)
- [Chart Suite](#chart-suite)

---

## Standard Suite

Common classes and commands for all applications.

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

## Microsoft Excel Suite

Excel specific classes and commands

### Commands

### `Excel comment text`

Returns or sets the text of the comment

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `Excel comment` | no |  |
| `text` | `text` | yes | The next text for the comment. |
| `start` | `integer` | yes | The starting point within the existing text that the next test should be placed. |
| `over write` | `boolean` | yes | Set to true to over write the existing text of the comment with the new text. |

**Returns:** `text` — The text of the comment.

### `Excel repeat`

Repeats the last user-interface action.

### `Run InProc Tests`

Run InProc Tests.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `command` | `text` | no | The command specifies test cases to run. |

**Returns:** `boolean`

### `accept all changes`

Accepts all changes in the specified shared workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `when` | `text` | yes | Specifies when all the changes are accepted. |
| `who` | `text` | yes | Specifies by whom all the changes are accepted. |
| `where` | `text` | yes | Specifies where all the changes are accepted. |

### `activate next`

Activates the current window, sends it to the back of the window z-order, then activates the next window according to the z-order.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `window` | no |  |

### `activate object`

Make the object active

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4001` | no |  |

### `activate previous`

Activates the specified window and then activates the window at the back of the window z-order.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `window` | no |  |

### `add chart autoformat`

Adds a custom chart autoformat to the list of available chart autoformats.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `chart` | `chart` | no | A chart that contains the format that will be applied when the new chart autoformat is applied. |
| `name` | `text` | no | The name of the autoformat. |
| `description` | `text` | yes | A description of the custom autoformat. |

### `add colorstop`

Adds ColorStop object to the ColorStops object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `colorstops` | no |  |
| `position` | `real` | no | Represents the position in which to apply the ColorStop. |

**Returns:** `colorstop`

### `add custom list`

Adds a custom list for custom autofill and/or custom sort.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `list array` | `XLCustomListType` | no | Specifies the source data, as either an array of strings or a range object. |
| `by row` | `boolean` | yes | Only used if list array is a range object. true to create a custom list from each row in the range. false to create a custom list from each column in the range. |

### `add data field`

Adds a data field to a PivotTable report.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `field` | `null` | no | The unique field on the server. |
| `caption` | `text` | yes | The label used in the PivotTable report to identify this data field. |
| `function` | `text` | yes | The function performed in the added data field. |

**Returns:** `pivot field`

### `add data validation`

Adds data validation to the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `validation` | no |  |
| `type` | `XlDVType` | no | The validation type. |
| `alert style` | `XlDVAlertStyle` | yes | The validation alert style. |
| `operator` | `XlFormatConditionOperator` | yes | The data validation operator. |
| `formula1` | `text` | yes | The first part of the data validation equation. |
| `formula2` | `text` | yes | The second part of the data validation when operator is operator between or operator not between, otherwise, this argument is ignored. |

### `add fields to pivot table`

Adds row, column, and page fields to a pivot table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `row fields` | `text` | yes | Specifies a list of pivot field names to be added as rows. The list elements are strings. |
| `column fields` | `text` | yes | Specifies a list of pivot field names to be added as columns. The list elements are strings. |
| `page fields` | `text` | yes | Specifies a list of pivot field names to be added as pages. The list elements are strings. |
| `add to table` | `boolean` | yes | Set to true to add the fields to the pivot table, none of the existing fields are replaced. False to replace existing fields with the new fields. The default value is false. |

### `add item to list`

Adds a string to the list control

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4016` | no |  |
| `item text` | `text` | no | The string that will be added. |
| `entry_index` | `integer` | yes | The position in the list to add the string. |

### `add member property field`

Adds a member property field to the display for the cube field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cube field` | no |  |
| `property` | `text` | no | The unique name of the member property. |
| `property order` | `text` | yes | Sets the PropertyOrder property value for a CubeField object. |
| `property displayed in` | `XlPropertyDisplayedIn` | yes |  |

### `add page item`

Adds an additional item to a multiple item page field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot field` | no |  |
| `item` | `text` | no | Source name of a PivotItem object, corresponding to the specific Online Analytical Processing member unique name. |
| `clear list` | `text` | yes | If False, adds a page item to the existing list. If True, deletes all current items and adds Item. |

### `add sortfield`

Creates a new sortfield and returns a sortfieldset object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sortfieldset` | no |  |
| `key` | `range` | no | Specifies a key value for the sort. |
| `sorton` | `XlSortOn` | yes | The field to sort on. |
| `order` | `XlSortOrder` | yes | Specifies the sort order. |
| `customorder` | `any` | yes | Specifies if a custom sort order should be used. |
| `dataoption` | `XlSortDataOption` | yes | Specifies the data option. |

### `allocate change`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot cell` | no |  |

### `allocate changes`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

### `apply filter`

Apply the defined filter

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `apply sort`

Sorts the range based on the currently applied sort states.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `apply theme`

Applies a theme or design template to the specified workbook

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `file name` | `text` | no |  |

### `arrange_windows`

Arranges the windows on the screen

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `arrange style` | `XlArrangeStyle` | yes | Specifies the way the windows will be arranged. The default is tiled. |
| `active workbook` | `boolean` | yes | Set to true to arrange only the visible windows of the active workbook, false otherwise.  The default is false. |
| `sync horizontal` | `boolean` | yes | Ignored if active workbook is false. Set to true to synchronize the windows of the active workbook when scrolling horizontally. False to not synchronize the windows. The default value is false. |
| `sync vertical` | `boolean` | yes | Ignored if active workbook is false. Set to true to synchronize the windows of the active workbook when scrolling vertically. False to not synchronize the windows. The default value is false. |

### `auto show`

Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot field` | no |  |
| `type` | `XLAutoShowType` | no | Use type_automatic to cause the pivot table to show the items that match the specified criteria. Use type_manual to disable this feature. |
| `range` | `XLAutoShowPosition` | no | The location at which to start showing items. |
| `count` | `integer` | no | The number of items to be shown. |
| `field` | `text` | no | The name of the base data field. |

### `auto sort`

Establishes automatic pivot table field-sorting rules.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot field` | no |  |
| `sort order` | `XlSortOrder` | no | The sort order. |
| `sort field` | `text` | no | The name of the sort key field. |

### `break link`

Converts formulas linked to other Microsoft Excel sources to values.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `name` | `text` | no | The name of the link. |
| `type` | `XlLinkType` | no | The type of link. |

### `bring to front`

Bring the object to the front of the z-order of objects.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4011` | no |  |

### `calculate`

Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet..

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4004` | no |  |

### `calculate full`

Forces a full calculation of the data in all open workbooks.

### `calculate full rebuild`

For all open workbooks, forces a full calculation of the data and rebuilds the dependencies

### `can check in`

Returns True if Excel can check in a specified workbook to a server.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

**Returns:** `boolean`

### `can check out`

Returns True if Excel can check out a specified workbook from a server.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `file name` | `text` | no | The server path and name of the workbook. |

**Returns:** `boolean`

### `cancel refresh`

Cancels all background queries for the specified query table. Use the refreshing property to determine whether a background query is currently in progress.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `query table` | no |  |

### `centimeters to points`

Converts a measurement from centimeters to points, one point equals 0.035 centimeters.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `centimeters` | `real` | no | The centimeter value to be converted. |

**Returns:** `real` — The value in points.

### `change connection`

Changes the connection of the specified PivotTable.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `connection` | `workbook connection` | no |  |

### `change file access`

Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `mode` | `XlFileAccess` | no | Specifies the new access mode. |
| `write password` | `text` | yes | Specifies the write-reserved password if the file is write reserved and mode is read write. Ignored if there's no password for the file or if mode is read only. |
| `notify` | `boolean` | yes | Set to true, or omit, to notify the user if the file cannot be immediately accessed. |

### `change link`

Changes a link from one document to another.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `name` | `text` | no | The name of the Microsoft Excel link to be changed, as it was returned from the link sources method. |
| `new name` | `text` | no | The new name of the link. |
| `type` | `XlLinkType` | yes | The link type. |

### `change pivot cache`

Changes the PivotCache of the specified PivotTable.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `pivot cache` | `text` | no | A PivotTable or PivotCache object that represents the new PivotCache for the specfied PivotTable. |

### `change scenario`

Changes the scenario to have a new set of changing cells and, optionally, scenario values.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `scenario` | no |  |
| `changing cells` | `XLRangeReference` | no | A range object that specifies the new set of changing cells for the scenario. The changing cells must be on the same sheet as the scenario. |
| `values` | `any` | yes | A list that contains the new scenario values for the changing cells. If this argument is omitted, the scenario values are assumed to be the current values in the changing cells. |

### `check in`

Returns a workbook from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `save changes` | `boolean` | yes | True saves the workbook to the server location. The default value is False. |
| `comments` | `text` | yes | Comments for the revision of the workbook being checked in. Only applies if SaveChanges equals True. |
| `make public` | `boolean` | yes | True allows the user to perform a publish on the workbook after being checked in. |

### `check in with version`

Returns a workbook from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `save changes` | `boolean` | yes | True saves the workbook to the server location. The default value is False. |
| `comments` | `text` | yes | Comments for the revision of the workbook being checked in. Only applies if SaveChanges equals True |
| `make public` | `boolean` | yes | True allows the user to perform a publish on the workbook after being checked in. |
| `version type` | `XlCheckInVersionType` | yes | Version number of the workbook. |

### `check out`

Copies a specified workbook from a server to a local computer for editing. Returns a String that represents the local path and file name of the workbook checked out.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `file name` | `text` | no | The server path and name of the workbook. |

### `check spelling`

Checks the spelling of an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4010` | no |  |
| `custom dictionary` | `text` | yes | A string that indicates the file name of the custom dictionary to examine if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used. |
| `ignore uppercase` | `boolean` | yes | Set to true to have Microsoft Excel ignore words that are all uppercase. False to have Microsoft Excel check words that are all uppercase. |
| `always suggest` | `boolean` | yes | Set to true to have Microsoft Excel display a list of suggested alternate spellings when an incorrect spelling is found. False to have Microsoft Excel wait for you to input the correct spelling. |

### `check spelling for`

Checks the spelling of a single word. Returns True if the word is found in one of the dictionaries.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `text to check` | `text` | no | The word to be checked. |
| `custom dictionary` | `text` | yes | A string that indicates the file name of the custom dictionary to examine if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used. |
| `ignore uppercase` | `boolean` | yes | Set to true to have Microsoft Excel ignore words that are all uppercase. |

**Returns:** `boolean`

### `circle invalid`

Circles invalid entries on the worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |

### `clear all filters`

The ClearAllFilters method deletes all filters currently applied to the PivotTable.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4008` | no |  |

### `clear arrows`

Clears the tracer arrows from the worksheet. Tracer arrows are added by using the auditing feature.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |

### `clear circles`

Clears circles from invalid entries on the worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |

### `clear colorstops`

Clears the represented object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `clear contents`

Clears the formulas from the range. Clears the data from a chart but leaves the formatting. Clears all the data, formatting, and formulas from a list object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list object` | no |  |

### `clear label filters`

This method deletes all label filters or all date filters in the PivotFilters collection of the PivotField.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot field` | no |  |

### `clear manual filter`

The ClearManualFilter method provides an easy way to set the Visible property to True for all items of a PivotField in PivotTables, and to empty the HiddenItemsList or VisibleItemsList collections in OLAP PivotTables.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4007` | no |  |

### `clear sortfieldset`

Clears all the sortfield objects.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `clear table`

The ClearTable method is used for clearing a PivotTable.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

### `clear value filters`

Calling this method deletes all value filters in the PivotFilters collection of the PivotField.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot field` | no |  |

### `commit changes`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

### `convert formula`

Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `formula to convert` | `text` | no | A string that contains the formula you want to convert. This must be a valid formula, and it must begin with an equal sign. |
| `from reference style` | `XlReferenceStyle` | no | The reference style of the formula. |
| `to reference style` | `XlReferenceStyle` | yes | The reference style you want returned. |
| `to absolute` | `XlReferenceStyle` | yes | Specifies the converted reference type. |
| `relative to` | `XLRangeReference` | yes | A range object that contains one cell. Relative references relate to this cell. |

**Returns:** `text` — The converted formula

### `convert to formulas`

The ConvertToFormulas method is new in 1st_Excel12 and is used for converting a PivotTable to cube formulas.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `convert filters` | `boolean` | no |  |

### `convert to range`

Converts a list object to a normal Excel range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list object` | no |  |

**Returns:** `range`

### `copy object`

Copies the object to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4012` | no |  |

### `copy picture`

Copies the selected object to the clipboard as a picture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4013` | no |  |
| `appearance` | `XlPictureAppearance` | yes | Specifies how the picture should be copied. |
| `format` | `XlCopyPictureFormat` | yes | The format of the picture. |

### `copy worksheet`

Copies the sheet to another location in the workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |
| `before` | `sheet` | yes | The sheet before which the copied sheet will be placed. You cannot specify before if you specify after. |
| `after` | `sheet` | yes | The sheet after which the copied sheet will be placed. You cannot specify after if you specify before. |

### `create cube file`

Creates a cube file from a PivotTable report connected to an Online Analytical Processing data source.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `file` | `text` | no | The name of the cube file to be created. |
| `measures` | `text` | yes | An array of unique names of measures that are to be part of the slice. |
| `levels` | `text` | yes | An array of strings. Each array item is a unique level name. |
| `members` | `text` | yes | An array of string arrays. The elements correspond, in order, to the hierarchies represented in the Levels array. |
| `properties` | `text` | yes | False results in no member properties being included in the slice. The default value is True. |

**Returns:** `text`

### `create new document`

Creates a new document linked to the specified hyperlink.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `hyperlink` | no |  |
| `file name` | `text` | no | The file name of the specified document. |
| `edit now` | `boolean` | no | Set to true to have the specified document open immediately in its associated editing environment.. The default value is true. |
| `overwrite` | `boolean` | no | Set to true to overwrite any existing file of the same name in the same folder. False if any existing file of the same name is preserved and the file name argument specifies a new file name. The default value is false. |

### `create pivot fields`

The CreatePivotFields method enables users to create all PivotFields of a CubeField.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cube field` | no |  |

### `create pivot table`

Creates a PivotTable report based on a PivotCache object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot cache` | no |  |
| `table destination` | `text` | no |  |
| `table name` | `text` | yes |  |
| `read data` | `text` | yes |  |
| `default version` | `text` | yes |  |

### `create summary for scenarios`

Creates a new worksheet that contains a summary report for the scenarios on the specified worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `worksheet` | no |  |
| `report type` | `XlSummaryReportType` | yes | Specifies the report type. |
| `result cells` | `XLRangeReference` | yes | Normally, this range refers to one or more cells containing the formulas that depend on the changing cell values for your model, that is, the cells that show the results of a particular scenario. |

### `cut`

Cuts the object to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4014` | no |  |

### `delete`

Deletes the object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4006` | no |  |

### `delete chart autoformat`

Removes a custom chart autoformat from the list of available chart autoformats.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `name` | `text` | no | The name of the custom autoformat to be removed. |

### `delete colorstop`

Clears the represented object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `colorstop` | no |  |

### `delete custom list`

Deletes a custom list.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `list num` | `integer` | no | The custom list number. This number must be greater than or equal to 5. |

### `delete number format`

Deletes a custom number format from the workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `number format` | `text` | no | Names the number format to be deleted. |

### `delete sortfield`

Removes the specified sortfield object from the sortfieldset collection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `discard change`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot cell` | no |  |

### `discard changes`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

### `double click`

Equivalent to double-clicking the active cell.

### `drill to`

The DrillTo method supports drilling to a specified PivotField from another PivotField.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4009` | no |  |
| `field` | `text` | no |  |

### `evaluate`

Converts a Microsoft Excel name to an object or value.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `name` | `any` | no | The name of the object, using the naming convention of Microsoft Excel. |

**Returns:** `any` — The resulting object or value.

### `exclusive access`

Assigns the current user exclusive access to the workbook that's open as a shared list.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

**Returns:** `boolean` — Specifies success or failure.

### `follow`

Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `hyperlink` | no |  |
| `new window` | `boolean` | yes | Set to true to display the target application in a new window. The default value is false. |
| `extra info` | `text` | yes | A string that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use extra info to specify the coordinates of an image map, the contents of a form, or a file name. |
| `method` | `MsoExtraInfoMethod` | yes | Specifies the way extra info is attached. |
| `header info` | `text` | yes | A string that specifies header information for the HTTP request. The default value is an empty string. |

### `follow hyperlink`

Displays a cached document, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `address` | `text` | no | The address of the target document. |
| `sub address` | `text` | yes | The location within the target document. The default value is the empty string. |
| `new window` | `boolean` | yes | Set to true to display the target application in a new window. The default value is false. |
| `extra info` | `text` | yes | A string that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use extra info to specify the coordinates of an image map, the contents of a form, or a file name. |
| `method` | `MsoAnimationType` | yes | Specifies the way extra info is attached. |
| `header info` | `text` | yes | A string that specifies header information for the HTTP request. The default value is an empty string. |

### `get border`

Returns the specified border object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4025` | no |  |
| `which border` | `XlBordersIndex` | no | This specifies which border object should be retrieved. |

**Returns:** `border`

### `get clipboard formats`

Returns a list of the  formats that are currently on the clipboard.

**Returns:** `XlClipboardFormat`

### `get custom list contents`

Returns a custom list of strings.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `list num` | `integer` | no | The list number. |

**Returns:** `any`

### `get custom list num`

Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `list array` | `text` | no | A list of strings. |

**Returns:** `integer`

### `get dialog`

Returns a the specified dialog object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `XlBuiltInDialog` | no |  |

**Returns:** `dialog`

### `get file converters`

Returns a list of all of the file converter objects.

**Returns:** `text`

### `get international`

Returns information about a specific international setting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `data type` | `XlApplicationInternational` | no | The international data to be returned. |

**Returns:** `any`

### `get list item`

Returns a string from the list

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4017` | no |  |
| `entry_index` | `integer` | yes | The index of the string to be returned. |

**Returns:** `text`

### `get open filename`

Displays the standard Open dialog box and gets a file name from the user without actually opening any files.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `file filter` | `text` | yes | This string is a list of comma-separated file type codes, for example, TEXT,XLA5,XLS4. If omitted, this argument defaults to all file types. |
| `button text` | `text` | yes | Specifies the text used for the open button in the dialog box. If this argument is omitted, the button text is Open. |
| `multi select` | `boolean` | yes | Set to true to allow multiple file names to be selected. Set to false to allow only one file name to be selected. The default value is false. |

**Returns:** `text` — The name of the selected file to open

### `get pivot data`

Returns a Range object with information about a data item in a PivotTable report.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `data field` | `text` | yes | The name of the field containing the data for the PivotTable. |
| `field1` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item1` | `text` | yes | The name of an item in field1. |
| `field2` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item2` | `text` | yes | The name of an item in field2. |
| `field3` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item3` | `text` | yes | The name of an item in field3. |
| `field4` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item4` | `text` | yes | The name of an item in field4. |
| `field5` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item5` | `text` | yes | The name of an item in field5. |
| `field6` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item6` | `text` | yes | The name of an item in field6. |
| `field7` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item7` | `text` | yes | The name of an item in field7. |
| `field8` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item8` | `text` | yes | The name of an item in field8. |
| `field9` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item9` | `text` | yes | The name of an item in field9. |
| `field10` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item10` | `text` | yes | The name of an item in field10. |
| `field11` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item11` | `text` | yes | The name of an item in field11. |
| `field12` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item12` | `text` | yes | The name of an item in field12. |
| `field13` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item13` | `text` | yes | The name of an item in field13. |
| `field14` | `text` | yes | The name of a column or row field in the PivotTable report. |
| `item14` | `text` | yes | The name of an item in field14. |

**Returns:** `range`

### `get pivot table data`

Retrieve data from a pivot table

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `name` | `text` | no | Describes a single cell in the pivot table, using syntax similar to the pivot select method or the pivot table references in calculated item formulas. |

**Returns:** `real`

### `get previous selections`

Returns a list of the last four ranges or names selected. Each element in the list is a range object.

**Returns:** `specifier`

### `get registered functions`

Returns information about functions in code resources that were registered with the REGISTER or REGISTER.ID macro functions.

**Returns:** `text`

### `get save as filename`

Displays the standard save as dialog box and gets a file name from the user without actually saving any files.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `initial filename` | `text` | yes | Specifies the suggested file name. If this argument is omitted, Microsoft Excel uses the active workbook's name. |
| `file filter` | `text` | yes | This string is a list of comma-separated file type codes, for example, TEXT, XLA5, XLS4. If omitted, this argument defaults to all file types. |
| `filter index` | `integer` | yes | This string is a list of comma-separated file type codes, for example, TEXT, XLA5, XLS4. If omitted, this argument defaults to all file types. |
| `button text` | `text` | yes | Specifies the text used for the save button in the dialog box. If this argument is omitted, the button text is Save. |

**Returns:** `text` — The selected file name.

### `get subtotals`

Returns subtotals displayed with the specified field. Valid only for nondata fields.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot field` | no |  |
| `subtotal index` | `XLSubTotalsIndex` | yes | Specifies the subtotal to be returned as follows: 1 automatic, 2  sum, 3  count, 4 average, 5  max, 6  min, 7  product, 8  count nums, 9  StdDev, 10  StdDevp, 11  Var, 12  Varp. |

**Returns:** `boolean`

### `get values`

Returns or sets a list that contains the current values of the changing cells for the scenario.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `scenario` | no |  |

**Returns:** `any`

### `get visible fields`

Returns a list of all the visible pivot fields. Visible pivot fields are shown as row, column, page or data fields.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

**Returns:** `specifier`

### `goto`

Selects any range in any workbook, and activates that workbook if it's not already active.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `reference` | `XLRangeReference` | yes | The destination. Can be a range object, a string that contains a cell reference in R1C1-style notation. If this argument is omitted, the destination is the last range you used the goto method to select. |
| `scroll` | `boolean` | yes | Set to true to scroll through the window so that the upper-left corner of the range appears in the upper-left corner of the window. The default is false. |

### `help`

Displays the Help window.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `help file` | `text` | yes | The name of the online Help file to display. If this argument isn't specified, Microsoft Excel Help is displayed. |
| `help context id` | `integer` | yes | Specifies the context ID number for the Help topic to display. If this argument isn't specified, the default topic for the specified Help file is displayed. |

### `highlight changes options`

Controls how changes are shown in a shared workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `when` | `XlHighlightChangesTime` | yes | The changes that are shown. |
| `who` | `text` | yes | The user or users whose changes are shown. Can be Everyone, Everyone but Me, or the name of one of the users of the shared workbook. |
| `where` | `XLRangeReference` | yes | An A1-style range reference that specifies the area to check for changes. |

### `inches to points`

Converts a measurement from inches to points.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `inches` | `real` | no | Specifies the inch value to be converted to points. |

**Returns:** `real` — The value of the inches converted to points.

### `input box`

Displays a dialog box for user input. Returns the information entered in the dialog box.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `prompt` | `text` | no | The message to be displayed in the dialog box. |
| `title` | `text` | yes | The title for the input box. If this argument is omitted, the default title is Input. |
| `default` | `XLInputDefault` | yes | Specifies a value that will appear in the text box when the dialog box is initially displayed. If this argument is omitted, the text box is left empty. |
| `left position` | `integer` | yes | Specifies an x position for the dialog box in relation to the upper-left corner of the screen, in points. |
| `top` | `integer` | yes | Specifies a y position for the dialog box in relation to the upper-left corner of the screen, in points. |
| `type` | `XLInputType` | yes | Specifies what type of data the result should be. |

**Returns:** `specifier`

### `intersect`

Returns a range object that represents the rectangular intersection of two or more ranges.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `range1` | `range` | no | One of the intersecting ranges. |
| `range2` | `range` | no | One of the intersecting ranges. |
| `range3` | `range` | yes | One of the intersecting ranges. |
| `range4` | `range` | yes | One of the intersecting ranges. |
| `range5` | `range` | yes | One of the intersecting ranges. |
| `range6` | `range` | yes | One of the intersecting ranges. |
| `range7` | `range` | yes | One of the intersecting ranges. |
| `range8` | `range` | yes | One of the intersecting ranges. |
| `range9` | `range` | yes | One of the intersecting ranges. |
| `range10` | `range` | yes | One of the intersecting ranges. |
| `range11` | `range` | yes | One of the intersecting ranges. |
| `range12` | `range` | yes | One of the intersecting ranges. |
| `range13` | `range` | yes | One of the intersecting ranges. |
| `range14` | `range` | yes | One of the intersecting ranges. |
| `range15` | `range` | yes | One of the intersecting ranges. |
| `range16` | `range` | yes | One of the intersecting ranges. |
| `range17` | `range` | yes | One of the intersecting ranges. |
| `range18` | `range` | yes | One of the intersecting ranges. |
| `range19` | `range` | yes | One of the intersecting ranges. |
| `range20` | `range` | yes | One of the intersecting ranges. |
| `range21` | `range` | yes | One of the intersecting ranges. |
| `range22` | `range` | yes | One of the intersecting ranges. |
| `range23` | `range` | yes | One of the intersecting ranges. |
| `range24` | `range` | yes | One of the intersecting ranges. |
| `range25` | `range` | yes | One of the intersecting ranges. |
| `range26` | `range` | yes | One of the intersecting ranges. |
| `range27` | `range` | yes | One of the intersecting ranges. |
| `range28` | `range` | yes | One of the intersecting ranges. |
| `range29` | `range` | yes | One of the intersecting ranges. |
| `range30` | `range` | yes | One of the intersecting ranges. |

**Returns:** `range` — The intersection region of the specified ranges.

### `large scroll`

Scrolls the contents of the window by pages.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4021` | no |  |
| `down` | `integer` | yes | The number of pages to scroll the contents down. |
| `up` | `integer` | yes | The number of pages to scroll the contents up. |
| `to right` | `integer` | yes | The number of pages to scroll the contents to the right. |
| `to left` | `integer` | yes | The number of pages to scroll the contents to the left. |

### `link info`

Returns the link date and update status.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `name` | `text` | no | The name of the link, as returned from the link sources method. |
| `link info` | `XlLinkInfo` | no | The type of information to be returned. |
| `type` | `XlLinkInfoType` | yes | The type of link to return. |

**Returns:** `specifier`

### `link sources`

Returns a list of links in the workbook. The names in the array are the names of the linked documents. Returns empty if there are no links.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `type` | `XlLinkType` | yes | The type of link to return. |

**Returns:** `text`

### `list formulas`

Creates a list of calculated pivot table items and fields on a separate worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

### `make connection`

Establishes a connection for the specified PivotTable cache.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

### `merge scenarios`

Merges the scenarios from the merge source worksheet into this worksheet

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `worksheet` | no |  |
| `merge source` | `e298` | no | The worksheet to merge with. |

### `merge workbook`

Merges changes from one workbook into an open workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `file name` | `text` | no | The file name of the workbook that contains the changes to be merged into the open workbook. |

### `modify`

Modifies data validation for a range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `validation` | no |  |
| `type` | `XlDVType` | yes | The validation type. |
| `alert style` | `XlDVAlertStyle` | yes | The validation alert style. |
| `operator` | `XlFormatConditionOperator` | yes | The data validation operator. |
| `formula1` | `text` | yes | The first part of the data validation equation. |
| `formula2` | `text` | yes | The second part of the data validation when operator is operator between or operator not between, otherwise, this argument is ignored. |

### `modify applies to range`

Changes the range that this format condition applies to.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4028` | no |  |
| `range` | `range` | no |  |

### `modify condition`

Modifies an existing conditional format.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `format condition` | no |  |
| `type` | `XlFormatConditionType` | no | Specifies whether the conditional format is based on a cell value or an expression. |
| `operator` | `XlFormatConditionOperator` | yes | The conditional format operator. |
| `formula1` | `text` | yes | The value or expression associated with the conditional format. |
| `formula2` | `text` | yes | he value or expression associated with the second part of the conditional format when operator is operator between or operator not between. |
| `string` | `text` | yes |  |
| `operator2` | `XlFormatConditionOperator` | yes | The conditional format operator. |

### `modify condition value`

Modifies an existing condition value.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `condition value` | no |  |
| `type` | `XlConditionValueTypes` | no | Specifies whether the conditional format is based on a cell value or an expression. |
| `condition value` | `any` | yes | The value or expression associated with the conditional format. |

### `modify sort key`

Modify the key value by which values are sorted in the field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sortfield` | no |  |
| `rng` | `range` | no | Specifies the key to be modified. |

### `new window on workbook`

Creates a new window or a copy of the specified workbook window.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

**Returns:** `window`

### `next Excel comment`

Returns the next Excel comment object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `Excel comment` | no |  |

**Returns:** `Excel comment`

### `on key`

Runs a specified procedure when a particular key or key combination is pressed.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `key` | `text` | no | Specifies the key to bind the procedure to. |
| `command key pressed` | `boolean` | yes | Specifies if the command key needs to pressed as well and the key specified by the key argument. |
| `shift key pressed` | `boolean` | yes | Specifies if the shift key needs to pressed as well and the key specified by the key argument. |
| `option key pressed` | `boolean` | yes | Specifies if the option key needs to pressed as well and the key specified by the key argument. |
| `control key pressed` | `boolean` | yes | Specifies if the control key needs to pressed as well and the key specified by the key argument. |
| `procedure` | `XLOnDataType` | yes | A string containing AppleScript commands. If this argument is empty nothing happens when key is pressed.  If procedure is omitted, key reverts to its normal result in Microsoft Excel. |

### `open data base`

Open a data base

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `filename` | `text` | no |  |
| `command text` | `text` | yes |  |
| `rcommand type` | `integer` | yes |  |
| `back ground query` | `null` | yes |  |
| `import data as` | `integer` | yes |  |

### `open links`

Opens the supporting documents for a link or links.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `name` | `text` | no | The name of the Microsoft Excel link, as returned from the link sources method. |
| `read only` | `boolean` | yes | Set to true to open documents as read-only. The default value is false. |
| `type` | `XlLinkType` | yes | The link type. |

### `open text file`

Loads and parses a text file as a new workbook with a single sheet that contains the parsed text-file data.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `filename` | `text` | no | Specifies the file name of the text file to be opened and parsed. |
| `origin` | `XlPlatform` | yes | Specifies the origin of the text file. |
| `start row` | `integer` | yes | The row number at which to start parsing text.  The default is 1. |
| `data type` | `XlTextParsingType` | yes | Specifies the column format of the data in the file. |
| `text qualifier` | `XlTextQualifier` | yes | Specifies the text qualifier |
| `consecutive delimiter` | `boolean` | yes | Set to true to have consecutive delimiters considered one delimiter.  The default is false. |
| `tab` | `boolean` | yes | Set to true to have the tab character be the delimiter.  The default is false. |
| `semicolon` | `boolean` | yes | Set to true to have the semicolon character be the delimiter.  The default is false. |
| `comma` | `boolean` | yes | Set to true to have the comma character be the delimiter.  The default is false. |
| `space` | `boolean` | yes | Set to true to have the comma character be the delimiter.  The default is false. |
| `use other` | `boolean` | yes | Set to true to have the character specified by the other char argument be the delimiter.  The data type argument must be delimited. The default is false. |
| `other char` | `text` | yes | This is required if the use other argument is true.  Specifies the delimiter character when Other is true. If more than one character is specified, only the first character of the string is used, the remaining characters are ignored. |
| `field info` | `XlColumnDataType` | yes | A list contain parse information for the individual columns of data. Formats are general format, text format, MDY format, DMY format, YMD format, MYD format, DYM format, YDM format, skip column. |
| `decimal separator` | `text` | yes | The decimal separator that Microsoft Excel uses when recognizing numbers. The default setting is the system setting. |
| `thousands separator` | `text` | yes | The thousands separator that Excel uses when recognizing numbers. The default setting is the system setting. |

### `open workbook`

Opens a workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `workbook file name` | `text` | no | The name of the file containing the workbook to be opened. |
| `update links` | `MyUDateLinks` | yes | Specifies the way links in the file are updated. If this argument is omitted, the user is prompted to specify how links will be updated. |
| `read only` | `boolean` | yes | Set to true to open the workbook in read-only mode. |
| `format` | `MyODelimiter` | yes | If Microsoft Excel is opening a text file, this argument specifies the delimiter character.  If this argument is omitted, the current delimiter is used. |
| `password` | `text` | yes | A string that contains the password required to open a protected workbook. If this argument is omitted and the workbook requires a password, the user is prompted for the password. |
| `write reserved password` | `text` | yes | A string that contains the password required to write to a write-reserved workbook. If this argument is omitted and the workbook requires a password, the user will be prompted for the password. |
| `ignore read only recommended` | `boolean` | yes | Set to true to have Microsoft Excel not display the read-only recommended message if the workbook was saved with the read-only recommended option. |
| `origin` | `XlPlatform` | yes | If the file is a text file, this argument indicates where it originated so that code pages and Carriage return/line feed, CR/LF can be mapped correctly. |
| `delimiter` | `text` | yes | If the file is a text file and the format argument is custom character delimiter, this argument is a string that specifies the character to be used as the delimiter. |
| `editable` | `boolean` | yes | If the file is a Microsoft Excel 4.0 add-in, this argument is true to open the add-in so that it's a visible window. If the file isn't an add-in, true prevents the running of any Auto_Open macros. |
| `notify` | `boolean` | yes | If the file cannot be opened in read/write mode, this argument is true to add the file to the file notification list. Excel will open the file as read-only, poll the file notification list, and then notify the user when the file becomes available. |
| `converter` | `integer` | yes | The index of the first file converter to try when opening the file. The specified file converter is tried first, if this converter doesn't recognize the file, all other converters are tried. |
| `add to mru` | `boolean` | yes | Set to true to add this workbook to the list of recently used files. The default value is false. |

**Returns:** `workbook` — A workbook object for the newly opened workbook.

### `open xml`

Open an XML file

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `filename` | `text` | no |  |
| `style sheets` | `integer` | yes |  |
| `load option` | `integer` | yes |  |

### `paste special on worksheet`

Pastes the contents of the clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |
| `format` | `text` | yes | A string that specifies the clipboard format of the data. |
| `link` | `boolean` | yes | Set to true to establish a link to the source of the pasted data. If the source data isn't suitable for linking or the source application doesn't support linking, this parameter is ignored. The default value is false. |
| `display as icon` | `boolean` | yes | Set to true to display the pasted as an icon. The default value is false. |
| `icon file name` | `text` | yes | The name of the file that contains the icon to use if display as icon is true |
| `icon index` | `integer` | yes | The index number of the icon within the icon file. |
| `icon label` | `text` | yes | The text label of the icon. |
| `no HTML formatting` | `boolean` | yes | Set to true to remove all formatting, hyperlinks, and images from HTML. Set to false to paste HTML as is. The default value is false. |

### `paste worksheet`

Pastes the contents of the Clipboard onto the sheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |
| `destination` | `XLRangeReference` | yes | A range object that specifies where the clipboard contents should be pasted. If this argument is omitted, the current selection is used. |
| `link` | `boolean` | yes | Set to true to establish a link to the source of the pasted data. If this argument is specified, the destination argument cannot be used. The default value is false. |

### `pivot select`

Selects part of a pivot table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `name` | `text` | no | The selection, in standard pivot table selection format. |
| `mode` | `XlPTSelectionMode` | no | Specifies the structured selection mode. |

### `previous Excel comment`

Returns the previous Excel comment object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `Excel comment` | no |  |

**Returns:** `Excel comment`

### `print out`

Prints the object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4002` | no |  |
| `from` | `integer` | yes | The number of the page at which to start printing. If this argument is omitted, printing starts at the beginning. |
| `to` | `integer` | yes | The number of the last page to print. If this argument is omitted, printing ends with the last page. |
| `copies` | `integer` | yes | The number of copies to print. If this argument is omitted, one copy is printed. |
| `preview` | `boolean` | yes | Set to true to have Microsoft Excel invoke print preview before printing the object. |
| `active printer` | `text` | yes | Sets the name of the active printer. |
| `print to file` | `boolean` | yes | Set to true to print to a file. |
| `collate` | `boolean` | yes | Set to true to collate multiple copies. |

### `print preview`

Shows a preview of the object as it would look when printed. This function has been deprecated.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4003` | no |  |
| `enable changes` | `boolean` | yes | Controls access to the page setup dialog and the ability to change margins from the preview window by enabling or disabling the setup and margins buttons, respectively. |

### `protect sharing`

Saves the workbook and protects it for sharing.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `file name` | `text` | yes | A string indicating the name of the saved file. You can include a full path. If you don't, Microsoft Excel saves the file in the current folder. |
| `password` | `text` | yes | A case-sensitive string indicating the protection password to be given to the file. Should be no longer than 15 characters. |
| `write reservation password` | `text` | yes | A string indicating the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened read-only. |
| `read only recommended` | `boolean` | yes | Set to true to display a message when the file is opened, recommending that the file be opened read-only. |
| `create backup` | `boolean` | yes | Set to true to create a backup file. |
| `sharing password` | `text` | yes | A string indicating the password to be used to protect the file for sharing. |
| `file format` | `XlFileFormat` | yes | A enum indicating the format of the file for sharing. |

### `protect workbook`

Protect workbook structure from changes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `password` | `text` | yes | A string that specifies a case-sensitive password for the sheet or workbook. If this argument is omitted, you can unprotect the sheet or workbook without using a password. Otherwise, you must specify the password to unprotect the sheet or workbook. |
| `structure` | `boolean` | yes | Set to true to protect the structure of the workbook, the relative position of the sheets. The default value is true. |
| `windows` | `boolean` | yes | Set to true to protect the workbook windows. If this argument is omitted, the windows aren't protected. |

### `protect worksheet`

Protects a worksheet so that it cannot be modified.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |
| `password` | `text` | yes | A string that specifies a case-sensitive password for the sheet or workbook. If this argument is omitted, you can unprotect the sheet or workbook without using a password. Otherwise, you must specify the password to unprotect the sheet or workbook. |
| `drawing objects` | `boolean` | yes | Set to true to protect shapes. The default value is false. |
| `worksheet contents` | `boolean` | yes | Set to true to protect contents. The default value is true. |
| `scenarios` | `boolean` | yes | Set to true to protect scenarios. The default value is true. |
| `user interface only` | `boolean` | yes | Set to true to protect the user interface, but not macros. If this argument is omitted, protection applies both to macros and to the user interface. |
| `allow formatting cells` | `boolean` | yes | Set to true to allow user format cells. The default value is false. |
| `allow formatting columns` | `boolean` | yes | Set to true to allow user format columns. The default value is false. |
| `allow formatting rows` | `boolean` | yes | Set to true to allow user format rows. The default value is false. |
| `allow inserting columns` | `boolean` | yes | Set to true to allow user insert columns. The default value is false. |
| `allow inserting rows` | `boolean` | yes | Set to true to allow user insert rows. The default value is false. |
| `allow inserting hyperlinks` | `boolean` | yes | Set to true to allow user insert hyperlinks. The default value is false. |
| `allow deleting columns` | `boolean` | yes | Set to true to allow user delete columns. The default value is false. |
| `allow deleting rows` | `boolean` | yes | Set to true to allow user delete rows. The default value is false. |
| `allow sorting` | `boolean` | yes | Set to true to allow user sort. The default value is false. |
| `allow filtering` | `boolean` | yes | Set to true to allow user use autofilter. The default value is false. |
| `allow using pivot table` | `boolean` | yes | Set to true to allow user use pivotTable reports. The default value is false. |

### `purge change history now`

Removes entries from the change log for the specified workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `days` | `integer` | no | The number of days that changes in the change log are to be retained. |
| `sharing password` | `text` | yes | The password that unprotects the workbook for sharing. If the workbook is protected for sharing with a password and this argument is omitted, the user is prompted for the password. |

### `refresh`

Updates the pivot table cache or query table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot cache` | no |  |

### `refresh all`

Refreshes all external data ranges and PivotTables in the workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

### `refresh data source values`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

### `refresh query table`

Updates the PivotTable cache or query table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `query table` | no |  |
| `background query` | `boolean` | yes | Set to true to return control to the procedure as soon as a database connection is made and the query is submitted and the query is updated in the background. |

**Returns:** `boolean`

### `refresh table`

Refreshes the pivot table from the source data. Returns true if it's successful.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

**Returns:** `boolean`

### `register xll`

Loads an XLL code resource and automatically registers the functions and commands contained in the resource.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `filename` | `text` | no | The file name containing the XLL code resource. |

**Returns:** `boolean` — Specifies success or failure.

### `reject all changes`

Rejects all changes in the specified shared workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `when` | `text` | yes | Specifies when all the changes are rejected. |
| `who` | `text` | yes | Specifies by whom all the changes are rejected. |
| `where` | `text` | yes | Specifies where all the changes are rejected. |

### `remove all items`

Removes all of the strings from the list

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4019` | no |  |

### `remove item`

Removes a specified string from the list

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4020` | no |  |
| `entry_index` | `integer` | no | The index of the string to be removed. |
| `count` | `integer` | yes | The number of string starting a the index that should be removed.  The default is one. |

### `remove user`

Disconnects the specified user from the shared workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `entry_index` | `integer` | no | The user index. |

### `repeat all labels`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `repeat` | `XlPivotFieldRepeatLabels` | no |  |

### `reset all page breaks`

Resets all page breaks on the specified worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |

### `reset colors`

Resets the color palette to the default colors.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

### `reset timer`

Resets the refresh timer for the specified PivotTable report to the last interval you set using the RefreshPeriod property.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4024` | no |  |

### `resize`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `list object` | no |  |
| `range` | `range` | yes |  |

### `row axis layout`

This method is used for simultaneously setting layout options for all existing PivotFields.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `layout` | `XlLayoutRowType` | no | Can be xlCompactRow, xlTabularRow, or xlOutlineRow. |

### `run VB Macro`

Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `e297` | no |  |
| `arg1` | `any` | yes | An argument to be passed to the macro. |
| `arg2` | `any` | yes | An argument to be passed to the macro. |
| `arg3` | `any` | yes | An argument to be passed to the macro. |
| `arg4` | `any` | yes | An argument to be passed to the macro. |
| `arg5` | `any` | yes | An argument to be passed to the macro. |
| `arg6` | `any` | yes | An argument to be passed to the macro. |
| `arg7` | `any` | yes | An argument to be passed to the macro. |
| `arg8` | `any` | yes | An argument to be passed to the macro. |
| `arg9` | `any` | yes | An argument to be passed to the macro. |
| `arg10` | `any` | yes | An argument to be passed to the macro. |
| `arg11` | `any` | yes | An argument to be passed to the macro. |
| `arg12` | `any` | yes | An argument to be passed to the macro. |
| `arg13` | `any` | yes | An argument to be passed to the macro. |
| `arg14` | `any` | yes | An argument to be passed to the macro. |
| `arg15` | `any` | yes | An argument to be passed to the macro. |
| `arg16` | `any` | yes | An argument to be passed to the macro. |
| `arg17` | `any` | yes | An argument to be passed to the macro. |
| `arg18` | `any` | yes | An argument to be passed to the macro. |
| `arg19` | `any` | yes | An argument to be passed to the macro. |
| `arg20` | `any` | yes | An argument to be passed to the macro. |
| `arg21` | `any` | yes | An argument to be passed to the macro. |
| `arg22` | `any` | yes | An argument to be passed to the macro. |
| `arg23` | `any` | yes | An argument to be passed to the macro. |
| `arg24` | `any` | yes | An argument to be passed to the macro. |
| `arg25` | `any` | yes | An argument to be passed to the macro. |
| `arg26` | `any` | yes | An argument to be passed to the macro. |
| `arg27` | `any` | yes | An argument to be passed to the macro. |
| `arg28` | `any` | yes | An argument to be passed to the macro. |
| `arg29` | `any` | yes | An argument to be passed to the macro. |
| `arg30` | `any` | yes | An argument to be passed to the macro. |

**Returns:** `text` — The result of the macro call.

### `run XLM Macro`

Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `e297` | no |  |
| `arg1` | `any` | yes | An argument to be passed to the macro. |
| `arg2` | `any` | yes | An argument to be passed to the macro. |
| `arg3` | `any` | yes | An argument to be passed to the macro. |
| `arg4` | `any` | yes | An argument to be passed to the macro. |
| `arg5` | `any` | yes | An argument to be passed to the macro. |
| `arg6` | `any` | yes | An argument to be passed to the macro. |
| `arg7` | `any` | yes | An argument to be passed to the macro. |
| `arg8` | `any` | yes | An argument to be passed to the macro. |
| `arg9` | `any` | yes | An argument to be passed to the macro. |
| `arg10` | `any` | yes | An argument to be passed to the macro. |
| `arg11` | `any` | yes | An argument to be passed to the macro. |
| `arg12` | `any` | yes | An argument to be passed to the macro. |
| `arg13` | `any` | yes | An argument to be passed to the macro. |
| `arg14` | `any` | yes | An argument to be passed to the macro. |
| `arg15` | `any` | yes | An argument to be passed to the macro. |
| `arg16` | `any` | yes | An argument to be passed to the macro. |
| `arg17` | `any` | yes | An argument to be passed to the macro. |
| `arg18` | `any` | yes | An argument to be passed to the macro. |
| `arg19` | `any` | yes | An argument to be passed to the macro. |
| `arg20` | `any` | yes | An argument to be passed to the macro. |
| `arg21` | `any` | yes | An argument to be passed to the macro. |
| `arg22` | `any` | yes | An argument to be passed to the macro. |
| `arg23` | `any` | yes | An argument to be passed to the macro. |
| `arg24` | `any` | yes | An argument to be passed to the macro. |
| `arg25` | `any` | yes | An argument to be passed to the macro. |
| `arg26` | `any` | yes | An argument to be passed to the macro. |
| `arg27` | `any` | yes | An argument to be passed to the macro. |
| `arg28` | `any` | yes | An argument to be passed to the macro. |
| `arg29` | `any` | yes | An argument to be passed to the macro. |
| `arg30` | `any` | yes | An argument to be passed to the macro. |

**Returns:** `text` — The result of the macro call.

### `save as`

Saves changes into a different file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |
| `filename` | `text` | no | A string that indicates the name of the file to be saved. You can include a full path. If you don't, Microsoft Excel saves the file in the current folder. |
| `file format` | `XlFileFormat` | yes | Specifies the file format to use when you save the file. |
| `password` | `text` | yes | A case-sensitive string, no more than 15 characters, that indicates the protection password to be given to the file. |
| `write reservation password` | `text` | yes | A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only. |
| `read only recommended` | `boolean` | yes | Set to true to display a message when the file is opened, recommending that the file be opened as read-only. |
| `create backup` | `boolean` | yes | Set to true to create a backup file. |
| `add to most recently used list` | `boolean` | yes | Set to true to add this workbook to the list of recently used files. The default value is false. |
| `save as local language` | `boolean` | yes | True saves files against the language of Microsoft Excel. False is the default, which saves files against the language of Visual Basic for Applications |

### `save as ODC`

Saves the PivotTable cache source as a Microsoft Office Data Connection file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `ODC file name` | `text` | no | Location to save the file. |
| `description` | `text` | yes | Description that will be saved in the file. |
| `keywords` | `text` | yes | Space-separated keywords that can be used to search for this file. |

### `save workbook as`

Saves changes into a different file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `filename` | `text` | yes | A string that indicates the name of the file to be saved. You can include a full path. If you don't, Microsoft Excel saves the file in the current folder. |
| `file format` | `XlFileFormat` | yes | Specifies the file format to use when you save the file. |
| `password` | `text` | yes | A case-sensitive string, no more than 15 characters, that indicates the protection password to be given to the file. |
| `write reservation password` | `text` | yes | A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only. |
| `read only recommended` | `boolean` | yes | Set to true to display a message when the file is opened, recommending that the file be opened as read-only. |
| `create backup` | `boolean` | yes | Set to true to create a backup file. |
| `access mode` | `XlSaveAsAccessMode` | yes | Specifies the access mode for the new file. |
| `conflict resolution` | `XlSaveConflictResolution` | yes | Specifies who conflict resolutions will be handled. |
| `add to most recently used list` | `boolean` | yes | Set to true to add this workbook to the list of recently used files. The default value is false. |

### `save workspace`

Saves the current workspace.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `workspace file name` | `text` | yes | The saved file name. |

### `scroll workbook tabs`

Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `window` | no |  |
| `sheets` | `integer` | yes | The number of sheets to scroll by. Use a positive number to scroll forward, a negative number to scroll backward, or 0  to not scroll at all. You must specify sheets if you don't specify position. |
| `position` | `XLScrollTabPosition` | yes | Specifies the sheet to scroll to. You must specify position if you don't specify sheets. |

### `send mail`

Opens a message window in your registered mail program for sending the specified document as an attachment.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

### `send to back`

Sends the object to the back of the z-order.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4015` | no |  |

### `set background picture`

Sets the background graphic for a worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |
| `picture file name` | `text` | yes | The name of the graphic file. |

### `set first priority`

Sets this condition to the highest priority.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4026` | no |  |

### `set last priority`

Sets this condition to the lowest priority.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4027` | no |  |

### `set list item`

Set a string in the list

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4018` | no |  |
| `entry_index` | `integer` | yes | The index of the string to be set. |
| `item text` | `text` | no | The new text to be set. |

### `set sort range`

Sets the starting and ending character positions for Sort object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sort` | no |  |
| `rng` | `range` | no | Specifies the range for the sort object. |

### `set subtotals`

Sets subtotals displayed with the specified field. Valid only for nondata fields.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot field` | no |  |
| `subtotal index` | `XLSubTotalType` | yes | Specifies the subtotal to be set as follows: 1 automatic, 2  sum, 3  count, 4 average, 5  max, 6  min, 7  product, 8  count nums, 9  StdDev, 10  StdDevp, 11  Var, 12  Varp. |
| `value` | `boolean` | no | Specifies the subtotal to be set as follows: 1 automatic, 2  sum, 3  count, 4 average, 5  max, 6  min, 7  product, 8  count nums, 9  StdDev, 10  StdDevp, 11  Var, 12  Varp. |

### `show`

Displays the built-in dialog box and waits for the user to input data.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4023` | no |  |

### `show all`

Show all the hidden data, but do not exist the filter mode

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `show all data`

Makes all rows of the currently filtered list visible. If autofilter is in use, this method changes the arrows to All.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |

### `show custom view`

Displays the custom view

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `custom view` | no |  |

### `show data form`

Displays the data form associated with the worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sheet` | no |  |

### `show levels`

Displays the specified number of row and/or column levels of an outline.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `outline` | no |  |
| `row levels` | `integer` | yes | Specifies the number of row levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is zero or is omitted, no action is taken on rows. |
| `column levels` | `integer` | yes | Specifies the number of column levels of an outline to display. If the outline has fewer levels than the number specified, Microsoft Excel displays all the levels. If this argument is zero or is omitted, no action is taken on columns. |

### `show pages`

Creates a new pivot table for each item in the page field. Each new pivot table is created on a new worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `page field` | `text` | yes | A string that names a single page field in the pivot table. |

### `small scroll`

Scrolls the contents of the window by rows or columns.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4022` | no |  |
| `down` | `integer` | yes | The number of rows to scroll the contents down. |
| `up` | `integer` | yes | The number of rows to scroll the contents up. |
| `to right` | `integer` | yes | The number of rows to scroll the contents to the right. |
| `to left` | `integer` | yes | The number of rows to scroll the contents to the left. |

### `subtotal location`

This method changes the subtotal location for all existing PivotFields.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |
| `location` | `XlSubtototalLocationType` | no | xlSubtotalLocationType can be either xlAtTop or xlAtBottom. |

### `toggle forms design`

Toggles Microsoft Office Excel into and out of design mode.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

### `undo`

Cancels the last user-interface action.

### `union`

Returns the union of two or more ranges.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `range1` | `range` | no | One of the ranges that will be put into the union range. |
| `range2` | `range` | no | One of the ranges that will be put into the union range. |
| `range3` | `range` | yes | One of the ranges that will be put into the union range. |
| `range4` | `range` | yes | One of the ranges that will be put into the union range. |
| `range5` | `range` | yes | One of the ranges that will be put into the union range. |
| `range6` | `range` | yes | One of the ranges that will be put into the union range. |
| `range7` | `range` | yes | One of the ranges that will be put into the union range. |
| `range8` | `range` | yes | One of the ranges that will be put into the union range. |
| `range9` | `range` | yes | One of the ranges that will be put into the union range. |
| `range10` | `range` | yes | One of the ranges that will be put into the union range. |
| `range11` | `range` | yes | One of the ranges that will be put into the union range. |
| `range12` | `range` | yes | One of the ranges that will be put into the union range. |
| `range13` | `range` | yes | One of the ranges that will be put into the union range. |
| `range14` | `range` | yes | One of the ranges that will be put into the union range. |
| `range15` | `range` | yes | One of the ranges that will be put into the union range. |
| `range16` | `range` | yes | One of the ranges that will be put into the union range. |
| `range17` | `range` | yes | One of the ranges that will be put into the union range. |
| `range18` | `range` | yes | One of the ranges that will be put into the union range. |
| `range19` | `range` | yes | One of the ranges that will be put into the union range. |
| `range20` | `range` | yes | One of the ranges that will be put into the union range. |
| `range21` | `range` | yes | One of the ranges that will be put into the union range. |
| `range22` | `range` | yes | One of the ranges that will be put into the union range. |
| `range23` | `range` | yes | One of the ranges that will be put into the union range. |
| `range24` | `range` | yes | One of the ranges that will be put into the union range. |
| `range25` | `range` | yes | One of the ranges that will be put into the union range. |
| `range26` | `range` | yes | One of the ranges that will be put into the union range. |
| `range27` | `range` | yes | One of the ranges that will be put into the union range. |
| `range28` | `range` | yes | One of the ranges that will be put into the union range. |
| `range29` | `range` | yes | One of the ranges that will be put into the union range. |
| `range30` | `range` | yes | One of the ranges that will be put into the union range. |

**Returns:** `range` — A range object that is the union of the supplied range objects.

### `unprotect`

Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4005` | no |  |
| `password` | `text` | yes | A string that denotes the case-sensitive password to use to unprotect the sheet or workbook.  If you omit this argument for a sheet that's protected with a password, you'll be prompted for the password. |

### `unprotect sharing`

Turns off protection for sharing and saves the workbook.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `sharing password` | `text` | yes | The workbook password. |

### `update`

Updates the link or pivot table.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `pivot table` | no |  |

### `update from file`

Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

### `update link`

Updates a Microsoft Excel  link.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |
| `name` | `text` | yes | The name of the Microsoft Excel link to be updated, as returned from the link sources method. |
| `type` | `XlLinkType` | yes | The link type. |

### `use default folder suffix`

Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `web options` | no |  |

### `web page preview`

Displays a preview of the specified workbook as it would look if saved as a Web page.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `workbook` | no |  |

### Classes

### `Excel comment`

Represents a cell comment.

*Plural:* `Excel comments`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `author` | `text` | r/o | Returns the author of the comment. |
| `shape object` | `shape` | r/o | Returns a shape object that represents the shape attached to the specified comment. |
| `visible` | `boolean` | r/w | Returns or sets if the specified object is visible. |

### `ODBC error`

*Plural:* `ODBC errors`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `error string` | `text` | r/o | Returns the ODBC error string. |
| `sql state` | `text` | r/o | Returns the SQL state error. |

### `Protection`

Represents the various types of protection options available for a worksheet

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow deleting columns` | `boolean` | r/o | Returns True if the deleting of columns is allowed on a protected worksheet. Read-only Boolean. |
| `allow deleting rows` | `boolean` | r/o | Returns True if the deleting of rows is allowed on a protected worksheet. Read-only Boolean. |
| `allow filtering` | `boolean` | r/o | Returns True if the filtering is allowed on a protected worksheet. Read-only Boolean. |
| `allow formatting cells` | `boolean` | r/o | Returns True if the formatting of cells is allowed on a protected worksheet. Read-only Boolean. |
| `allow formatting columns` | `boolean` | r/o | Returns True if the formatting of columns is allowed on a protected worksheet. Read-only Boolean. |
| `allow formatting rows` | `boolean` | r/o | Returns True if the formatting of rows is allowed on a protected worksheet. Read-only Boolean. |
| `allow inserting columns` | `boolean` | r/o | Returns True if the inserting of columns is allowed on a protected worksheet. Read-only Boolean. |
| `allow inserting hyperlinks` | `boolean` | r/o | Returns True if the inserting of hyperlinks is allowed on a protected worksheet. Read-only Boolean. |
| `allow inserting rows` | `boolean` | r/o | Returns True if the inserting of rows is allowed on a protected worksheet. Read-only Boolean. |
| `allow sorting` | `boolean` | r/o | Returns True if the sorting is allowed on a protected worksheet. Read-only Boolean. |
| `allow using pivot table` | `boolean` | r/o | Returns True if the using of pivot table is allowed on a protected worksheet. Read-only Boolean. |

### `above average format condition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `above or below` | `XlAboveBelow` | r/w |  |
| `applies to` | `range` | r/o | Returns the range this format condition applies to. Read Only. |
| `calc for` | `XlCalcFor` | r/w |  |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `format condition priority` | `integer` | r/w | Specifies the priority for the format condition. Read/Write. |
| `format condition type` | `XlFormatConditionType` | r/o | Return the conditional format type. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `number format` | `text` | r/w | Returns or sets the format code for the object. |
| `number of standard deviations` | `integer` | r/w |  |
| `pivot condition scope type` | `XlPivotConditionScope` | r/w | Returns or sets the part of the pivot table that the pivot table format condition is scoped to |
| `pivot table condition` | `boolean` | r/o | Tells whether this format condition is a pivot table condition. |
| `stop if true` | `boolean` | r/o | Tells whether the processing of format conditions stops if this condition is true. Read Only. |

### `active filter`

*Plural:* `active filters`

### `add in`

Represents a single add-in, either installed or not installed.

*Plural:* `add ins`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `full name` | `text` | r/o | Returns the name of the add in, including its path on disk, as a string. |
| `installed` | `boolean` | r/w | Returns or sets if the add in is installed. |
| `name` | `text` | r/o | Returns the name of the object. |
| `path` | `text` | r/o | Returns the complete path of the object, excluding the final separator and name of the add in. |

### `application`

A representation of the Microsoft Excel application.

*Plural:* `applications`

**Contains:** `add in`, `chart sheet`, `command bar`, `named item`, `range`, `cell`, `row`, `column`, `window`, `workbook`, `sheet`, `worksheet`, `international macro sheet`, `macro sheet`, `recent file`, `ODBC error`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `AutoFormatAsYouTypeReplaceHyperlinks` | `boolean` | r/w | True if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type. |
| `Excel cursor` | `XlMousePointer` | r/w | Returns or sets the appearance of the mouse pointer in Microsoft Excel. |
| `ODBC timeout` | `integer` | r/w | Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. |
| `active cell` | `cell` | r/o | Returns a range object that represents the active cell in the active window, the window on top, or in the specified window. If the window isn't displaying a worksheet, this property fails. |
| `active chart` | `chart` | r/o | Returns a chart object that represents the active chart, either an embedded chart or a chart sheet. An embedded chart is considered active when it's either selected or activated. |
| `active printer` | `text` | r/w | Returns or sets the name of the active printer. |
| `active sheet` | `sheet` | r/o | Returns an object that represents the active sheet, the sheet on top, in the active workbook or in the specified window or workbook. |
| `active window` | `window` | r/o | Returns a window object that represents the active window, the window on top. |
| `active workbook` | `workbook` | r/o | Returns a workbook object that represents the workbook in the active window, the window on top. |
| `alert before overwriting` | `boolean` | r/w | Returns or sets if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation. |
| `alt startup path` | `text` | r/w | Returns or sets the name of the alternate startup folder. |
| `arbitrary XML support available` | `boolean` | r/o | Returns a Boolean value that indicates whether the XML features in Microsoft Excel are available |
| `ask to update links` | `boolean` | r/w | Returns or sets if Microsoft Excel asks the user to update links when opening files with links. |
| `autocorrect object` | `autocorrect` | r/o | Returns an autocorrect object that represents the Microsoft Excel AutoCorrect attributes. |
| `automation security` | `MsoAutomationSecurity` | r/w |  |
| `build` | `integer` | r/o | Returns the Microsoft Excel build number. |
| `calculate before save` | `boolean` | r/w | Returns or sets if workbooks are calculated before they're saved to disk. |
| `calculation` | `XlCalculation` | r/w | Returns or sets the calculation mode. |
| `calculation version` | `integer` | r/o | Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel. |
| `caption` | `text` | r/o | Returns the name of the application. |
| `cell drag and drop` | `boolean` | r/w | Returns or sets if dragging and dropping cells is enabled. |
| `command underlines` | `XlCommandUnderlines` | r/w | Returns or sets the state of the command underlines in Microsoft Excel. |
| `copy objects with cells` | `boolean` | r/w | Returns or sets if objects are cut, copied, extracted, and sorted with cells. |
| `custom list count` | `integer` | r/o | Returns the number of defined custom lists, including built-in lists. |
| `cut copy mode` | `XlCutCopyMode` | r/w | Returns or sets the status of cut or copy mode. |
| `data entry mode` | `XLDataEntryMode` | r/w | Returns or sets data entry mode, as shown in the following table. When in data entry mode, you can enter data only in the unlocked cells in the currently selected range. |
| `default file path` | `text` | r/w | Returns or sets the default path that Microsoft Excel uses when it opens files. |
| `default save format` | `XlFileFormat` | r/w | Returns or sets the default format for saving files. |
| `default web options object` | `default web options` | r/o | Returns the default web object. |
| `display alerts` | `boolean` | r/w | Returns or sets if Microsoft Excel displays certain alerts and messages while handling events from AppleScript. |
| `display comment indicator` | `XlCommentDisplayMode` | r/w | Returns or sets the way cells display comments and indicators. |
| `display excel 4 menus` | `boolean` | r/w | Returns or sets if Microsoft Excel displays version 4.0 menu bars. |
| `display formula bar` | `boolean` | r/w | Returns or sets  if the formula bar is displayed. |
| `display full screen` | `boolean` | r/w | Returns or sets if Microsoft Excel is in full-screen mode. |
| `display function tooltips` | `boolean` | r/w | Returns or sets if function tool tips can be displayed. |
| `display insert options` | `boolean` | r/w | Returns or sets if the insert options button should be displayed. |
| `display note indicator` | `boolean` | r/w | Returns or sets if cells containing notes display cell tips and contain note indicators, small dots in their upper-right corners. |
| `display recent files` | `boolean` | r/w | Returns or sets if the list of recently used files is displayed on the file menu. |
| `display scroll bars` | `boolean` | r/w | Returns or sets if scroll bars are visible for all workbooks. |
| `display status bar` | `boolean` | r/w | Returns or sets if the status bar is displayed. |
| `edit directly in cell` | `boolean` | r/w | Returns or sets if Microsoft Excel allows editing in cells. |
| `enable animations` | `boolean` | r/w | Returns or sets if animated insertion and deletion is enabled. |
| `enable autocomplete` | `boolean` | r/w | Returns or sets if the autocomplete feature is enabled. |
| `enable cancel key` | `XlEnableCancelKey` | r/w | Controls how Microsoft Excel handles CTRL-BREAK, ESC, or cmd-period user interruptions to the running procedure. |
| `enable events` | `boolean` | r/w | Returns or sets if events are enabled for the application object. |
| `enable sound` | `boolean` | r/w | Returns or sets if sound is enabled for Microsoft Office. |
| `extend list` | `boolean` | r/w | Returns or sets if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list. |
| `fixed decimal` | `boolean` | r/w | All data entered after this property is set to true will be formatted with the number of fixed decimal places set by the fixed decimal places property. |
| `fixed decimal places` | `integer` | r/w | Returns or sets the number of fixed decimal places used when the fixed decimal property is set to true. |
| `frontmost` | `null` | r/o | Returns if the application is the frontmost window. |
| `include empty cells in lists` | `boolean` | r/w | Returns or sets if empty cells are included in range lists. |
| `iteration` | `boolean` | r/w | Returns or sets  if Microsoft Excel will use iteration to resolve circular references. |
| `library path` | `text` | r/o | Returns the path to the Library folder. |
| `math coprocessor available` | `boolean` | r/o | Returns true if a math coprocessor is available. |
| `max change` | `real` | r/w | Returns or sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references. |
| `max iterations` | `integer` | r/w | Returns or sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference. |
| `measurement unit` | `integer` | r/w | Returns or set the current unit of measure. |
| `memory free` | `integer` | r/o | Returns the amount of memory that's still available for Microsoft Excel to use, in bytes. |
| `memory total` | `integer` | r/o | Returns the total amount of memory, in bytes, that's available to Microsoft Excel, including memory already in use. |
| `memory used` | `integer` | r/o | Returns the amount of memory that Microsoft Excel is currently using, in bytes. |
| `move after return` | `boolean` | r/w | Returns or sets if the active cell will be moved as soon as the ENTER/RETURN key is pressed. |
| `move after return direction` | `XlDirection` | r/w | Returns or sets the direction in which the active cell is moved when the user presses ENTER. |
| `name` | `text` | r/o | Returns the name of the application. |
| `network templates path` | `text` | r/o | Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. |
| `operating system` | `text` | r/o | Returns the name and version number of the current operating system. |
| `organization name` | `text` | r/o | Returns the registered organization name. |
| `path` | `text` | r/o | Returns the complete path of the application, excluding the final separator and name of the application. |
| `path separator` | `text` | r/o | Returns the path separator character. |
| `pivot table selection` | `boolean` | r/w | Returns or sets if pivot tables use structured selection. |
| `prompt for summary info` | `boolean` | r/w | Returns or sets if Microsoft Excel asks for summary information when files are first saved. |
| `reference style` | `XlReferenceStyle` | r/w | Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style. |
| `roll zoom` | `boolean` | r/w | Returns or sets if the IntelliMouse zooms instead of scrolling. |
| `screen updating` | `boolean` | r/w | Returns or sets if screen updating is turned on. Turn screen updating off to speed up your AppleScript code. You won't be able to see what the code is doing, but it will run faster. |
| `selection` | `range` | r/o | Returns the selected object in the active window. |
| `sheets in new workbook` | `integer` | r/w | Returns or sets the number of sheets that Microsoft Excel automatically inserts into new workbooks. |
| `show chart tip names` | `boolean` | r/w | Returns or sets if charts show chart tip names. |
| `show chart tip values` | `boolean` | r/w | Returns or sets if charts show chart tip values. |
| `show tool tips` | `boolean` | r/w | Returns or sets if tool tips are turned on. |
| `spelling options` | `xlspelling options` | r/o | Returns the default spelling options object. |
| `standard font` | `text` | r/w | Returns or sets the name of the standard font. |
| `standard font size` | `real` | r/w | Returns or sets the standard font size, in points. |
| `startup dialog` | `boolean` | r/w | Returns or sets if the startup dialog should be shown when starting up the application. |
| `startup path` | `text` | r/o | Returns the complete path of the startup folder, excluding the final separator. |
| `status bar` | `text` | r/w | Returns or sets the text in the status bar. |
| `templates path` | `text` | r/o | Returns the local path where templates are stored. |
| `this cell` | `null` | r/o | Returns the cell in which the user-defined function is being called from as a Range object. |
| `transition menu key` | `text` | r/w | Returns or sets the alternate menu or help key. |
| `transition menu key action` | `XLTransitionMenuKeyAction` | r/w | Returns or sets the action taken when the alternate menu key is pressed. |
| `usable height` | `real` | r/o | Returns the maximum height of the space that a window can occupy in points. |
| `usable width` | `real` | r/o | Returns the maximum width of the space that a window can occupy in points. |
| `user name` | `text` | r/w | Returns or sets the name of the current user. |

### `autofilter`

Represents autofiltering for the specified worksheet.

*Plural:* `autofilters`

**Contains:** `filter`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `autofiltermode` | `boolean` | r/o | Returns True if filters have been defined else False. |
| `range object` | `range` | r/o | Returns a range object that represents the range to which the specified autofilter applies. |
| `sort object` | `sort` | r/o | Returns the sort object in the auto filter object. |

### `border`

Represents the border of an object.

*Plural:* `borders`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `integer` | r/w | Returns or sets the primary color of the object. |
| `color index` | `XlColorIndex` | r/w | Returns or sets the color of the border. The color is specified as an index value into the current color palette. |
| `line style` | `XlLineStyle` | r/w | Returns or sets the line style for the border. |
| `theme color` | `XlThemeColor` | r/w | Returns or sets the theme color in the applied color scheme that is associated with the specified object. |
| `tint and shade` | `real` | r/w | Returns or sets a Single that lightens or darkens a color. |
| `weight` | `XlBorderWeight` | r/w | Returns or sets the weight of the border. |

### `button`

Represents a button control.

*Plural:* `buttons`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accelerator` | `text` | r/w | Returns or sets the accelerator character for this control. |
| `add indent` | `boolean` | r/w | Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. |
| `auto scale font` | `boolean` | r/w | Returns or sets  if the text in the object changes font size when the object size changes. |
| `auto size` | `boolean` | r/w | Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries. |
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `cancel button` | `boolean` | r/w | Returns or sets if this button is the cancel button. |
| `caption` | `text` | r/w | Returns or sets the caption for this object. |
| `control text` | `text` | r/w | Returns or sets the default text for the control. |
| `default button` | `boolean` | r/w | Returns or sets if this button is the default button. |
| `dismiss button` | `boolean` | r/w | Returns or sets if this button is the dismiss button. |
| `enabled` | `boolean` | r/w | Returns or sets if the object is enabled. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `formula` | `text` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `height` | `real` | r/w | Returns or set the height of the object. |
| `help button` | `boolean` | r/w | Returns or sets if this button is the help button. |
| `horizontal alignment` | `XlHAlign` | r/w | Returns or sets the horizontal alignment for the object. |
| `left position` | `real` | r/w | Returns or sets the position of the specified object, in points. |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `locked text` | `boolean` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `orientation` | `XlOrientation` | r/w | May also be a number value from -90 to 90 degrees. |
| `phonetic accelerator` | `text` | r/w | Returns or sets the link to a range. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `top` | `real` | r/w | Returns the top position of the specified object, in points. |
| `top left cell` | `range` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `vertical alignment` | `XlVerticalAlignmentTarget` | r/w | Returns or sets the vertical alignment of the object. |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `calculated field`

*Plural:* `calculated fields`

### `calculated item`

*Plural:* `calculated items`

### `calculated member`

Represents the calculated fields and calculated items for PivotTables with Online Analytical Processing data sources.

*Plural:* `calculated members`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `_default` | `text` | r/o |  |
| `display folder` | `text` | r/o | An ST_Xstring attribute that specifies the display folder of this PivotTable named set. |
| `dynamic` | `boolean` | r/o |  |
| `flatten hierarchies` | `boolean` | r/o |  |
| `formula` | `text` | r/o | Returns the member's formula in multidimensional expressions syntax. |
| `hierarchize distinct` | `boolean` | r/o |  |
| `is valid` | `boolean` | r/o | Returns a Boolean that indicates whether the specified calculated member has been successfully instantiated with the OLAP provider during the current session. |
| `name` | `text` | r/o | Calculated Member Name. |
| `solve order` | `integer` | r/o | Calculated Members Solve Order. |
| `source name` | `text` | r/o | Returns the specified object's name as it appears in the original source data for the specified PivotTable report. |
| `type` | `XlCalculatedMemberType` | r/o |  |

### `checkbox`

*Plural:* `checkboxes`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accelerator` | `text` | r/w | Returns or sets the accelerator character for this control. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `text` | r/w | Returns or sets the caption for this object. |
| `control text` | `text` | r/w | Returns or sets the default text for the control. |
| `display threeD shading` | `boolean` | r/w |  |
| `enabled` | `boolean` | r/w | Returns or sets if the control is enabled |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or sets the height of the control. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the left position of the control |
| `linked cell` | `text` | r/w |  |
| `locked` | `boolean` | r/w | Returns or sets whether the control is locked for editing. |
| `locked text` | `boolean` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `text` | r/w | Returns or sets the name of this control |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `phonetic accelerator` | `text` | r/w | Returns or sets the link to a range. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `top` | `real` | r/w | Returns or sets the position of the top of the control. |
| `top left cell` | `range` | r/o | Returns the top left cell of the range within which the control is positioned. |
| `value` | `XlCheckBoxState` | r/w |  |
| `visible` | `boolean` | r/w | Returns or sets the current value of the control |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text. |

### `child item`

*Plural:* `child items`

### `color scale criteria`

### `color scale criterion`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color scale criterion index` | `integer` | r/o | The index of the color scale criterion. Read only. |
| `color scale criterion type` | `XlConditionValueTypes` | r/w | Returns or sets the condition value type of the color scale criterion. Read/Write. |
| `color scale criterion value` | `any` | r/w |  |
| `format color` | `format color` | r/o | Returns the FormatColor for the ColorScaleCriterion object. |

### `color scale format condition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `applies to` | `range` | r/o | Returns the range this format condition applies to. Read Only. |
| `color scale criteria` | `color scale criteria` | r/o | Returns the ColorScaleCriteria for the ColorScale object. |
| `color scale type` | `integer` | r/w |  |
| `format condition priority` | `integer` | r/w | Specifies the priority for the format condition. Read/Write. |
| `format condition type` | `XlFormatConditionType` | r/o | Return the conditional format type. |
| `formula` | `text` | r/w | Returns or sets the value or expression associated with the conditional format or data validation. |
| `pivot condition scope type` | `XlPivotConditionScope` | r/w | Returns or sets the part of the pivot table that the pivot table format condition is scoped to |
| `pivot table condition` | `boolean` | r/o | Tells whether this format condition is a pivot table condition. |
| `stop if true` | `boolean` | r/o | Tells whether the processing of format conditions stops if this condition is true. Read Only. |

### `colorstop`

Represents the color stop point for a gradient fill in an range or selection.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `integer` | r/w | Returns or sets the primary color of the object. |
| `colorstop position` | `real` | r/w | Returns or sets the position of the ColorStop. |
| `theme color` | `XlThemeColor` | r/w | Returns or sets the theme color in the applied color scheme that is associated with the specified object. |
| `tint and shade` | `real` | r/w | Returns or sets a Single that lightens or darkens a color. |

### `colorstops`

A collection of all the ColorStop objects for the specified series.

### `column field`

*Plural:* `column fields`

### `column item`

*Plural:* `column items`

### `condition value`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `condition value type` | `XlConditionValueTypes` | r/o |  |
| `condition value value` | `any` | r/o |  |

### `cube field`

Represents a hierarchy or measure field from an OLAP cube

*Plural:* `cube fields`

**Contains:** `pivot field`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `all items visible` | `boolean` | r/o | The AllItemsVisible property checks whether manual filtering is applied to a PivotField or CubeField. |
| `caption` | `text` | r/w | The label text for the cube field. |
| `cube field sub type` | `XlCubeFieldSubType` | r/o | Specifies the type of a CubeField. |
| `cube field type` | `XlCubeFieldType` | r/o | Indicates whether the OLAP cube field is a hierarchy field or a measure field. |
| `current page name` | `text` | r/w | Returns or sets the page name for a CubeField. |
| `drag to column` | `boolean` | r/w | True if the specified field can be dragged to the column position. |
| `drag to data` | `boolean` | r/w | True if the specified field can be dragged to the data position. |
| `drag to hide` | `boolean` | r/w | True if the specified field can be dragged to the column position. |
| `drag to page` | `boolean` | r/w | True if the field can be dragged to the page position. |
| `drag to row` | `boolean` | r/w | True if the field can be dragged to the row position. |
| `enable multiple page items` | `boolean` | r/w | True to allow multiple items in the page field area for OLAP PivotTables to be selected. |
| `flatten hierarchies` | `boolean` | r/w |  |
| `has member properties` | `boolean` | r/o | Returns True when there are member properties specified to be displayed for the cube field. |
| `hierarchize distinct` | `boolean` | r/w |  |
| `include new items in filter` | `boolean` | r/w | The IncludeNewItemsInFilter property is used to track included or excluded items in OLAP PivotTables. |
| `is date` | `boolean` | r/o | Returns True if the CubeField is a date. |
| `layout form` | `XlLayoutFormType` | r/o | Returns or sets the way the specified PivotTable items appear -- in table format or in outline format. |
| `layout subtotal location` | `XlSubtototalLocationType` | r/w | Returns or sets the position of the PivotTable field subtotals in relation to, either above or below, the specified field. |
| `name` | `text` | r/o | Returns the name of the object. |
| `orientation` | `XlPivotFieldOrientation` | r/w | The location of the field in the specified PivotTable report. |
| `position` | `integer` | r/w | Position of the item in its field if the item is currently showing. |
| `show in field list` | `boolean` | r/w | When set to True, a CubeField object will be shown in the field list. |
| `treeview control` | `treeview control` | r/o |  |
| `value` | `text` | r/o | Returns the name of the specified field. |

### `custom view`

Represents a custom workbook view.

*Plural:* `custom views`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `custom view print settings` | `boolean` | r/o | True if print settings are included in the custom view. |
| `name` | `text` | r/o | Returns the name of this object. |
| `row col settings` | `boolean` | r/o | Returns true if the custom view includes settings for hidden rows and columns, including filter information. |

### `data field`

*Plural:* `data fields`

### `databar border`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `databar border color` | `format color` | r/o |  |
| `databar border type` | `XlDataBarBorderType` | r/w | Returns or sets the type of border of the databar |

### `databar format condition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `applies to` | `range` | r/o | Returns the range this format condition applies to. Read Only. |
| `databar axis color` | `format color` | r/o |  |
| `databar axis position` | `XlDataBarAxisPosition` | r/w | Returns or sets the position of the databar axis |
| `databar bar color` | `format color` | r/o |  |
| `databar border` | `databar border` | r/o | Returns the DataBarBorder for the Databar object. |
| `databar direction` | `integer` | r/w | Specifies the direction of the databar. Read/Write. |
| `databar fill type` | `XlDataBarFillType` | r/w | Returns or sets the type of fill of the databar |
| `format condition priority` | `integer` | r/w | Specifies the priority for the format condition. Read/Write. |
| `format condition show value` | `boolean` | r/w | Specifies whether to show the cell value along with the databar. Read/Write. |
| `format condition type` | `XlFormatConditionType` | r/o | Return the conditional format type. |
| `formula` | `text` | r/w | Returns or sets the value or expression associated with the conditional format or data validation. |
| `max point condition value` | `condition value` | r/o | Returns the ConditionValue for the maximum point of the DataBar object. |
| `max point percent` | `integer` | r/w | Specifies the percentage of the data bar to draw at the maximum point. Read/Write. |
| `min point condition value` | `condition value` | r/o | Returns the ConditionValue for the minimum point of the DataBar object. |
| `min point percent` | `integer` | r/w | Specifies the percentage of the data bar to draw at the minimum point. Read/Write. |
| `negative bar format` | `negative bar format` | r/o | Returns the NegativeBarFormat for the Databar object. |
| `pivot condition scope type` | `XlPivotConditionScope` | r/w | Returns or sets the part of the pivot table that the pivot table format condition is scoped to |
| `pivot table condition` | `boolean` | r/o | Tells whether this format condition is a pivot table condition. |
| `stop if true` | `boolean` | r/o | Tells whether the processing of format conditions stops if this condition is true. Read Only. |

### `default web options`

Contains global application-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow png` | `boolean` | r/w | Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save documents as a Web page. |
| `always save in default encoding` | `boolean` | r/w | Returns or sets if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened. |
| `encoding` | `MsoEncoding` | r/w | Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document. |
| `location of components` | `text` | r/w | Returns or sets the central URL, on the intranet or Web, or path, local or network, to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. |
| `pixels per inch` | `integer` | r/w | Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. |
| `screen size` | `MsoScreenSize` | r/w | Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser. |
| `update links on save` | `boolean` | r/w | Returns or sets if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved. If false, the links are not updated. |

### `dialog`

Represents a built-in Microsoft Excel dialog box.

*Plural:* `dialogs`

### `display format`

Represents the formatting shown to the user.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `add indent` | `boolean` | r/o | Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `formula hidden` | `boolean` | r/o | Returns or sets if the formula will be hidden when the workbook or worksheet is protected. |
| `horizontal alignment` | `XlHAlign` | r/o | Returns or sets the horizontal alignment for the range. |
| `indent level` | `integer` | r/o | Returns or sets the indent level for the range or style. Can be an integer from 0 to 15. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `locked` | `boolean` | r/o | Returns or sets  if the range is locked. |
| `merge cells` | `boolean` | r/o | Returns or sets if the range contains merged cells. |
| `number format` | `text` | r/o | Returns or sets the format code for the object. |
| `reading order` | `XLDefaultSheetDir` | r/o | Returns or sets the reading order for the specified object. |
| `shrink to fit` | `boolean` | r/o | Returns or sets if text automatically shrinks to fit in the available column width. |
| `style object` | `style` | r/o | Returns or sets a style object that represents the style of the specified range. |
| `text orientation` | `XlOrientation` | r/o | The text orientation. Can be a number value from -90 to 90 degrees. |
| `vertical alignment` | `XlVerticalAlignmentTarget` | r/o | Returns or sets the vertical alignment of the range. |
| `wrap text` | `boolean` | r/o | Returns or sets if Microsoft Excel wraps the text in the object. |

### `document`

*Plural:* `documents`

### `dropdown`

*Plural:* `dropdowns`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `text` | r/w | Returns or sets the caption for the control. |
| `control text` | `text` | r/w | Returns or sets the default text for the control. |
| `display threeD shading` | `boolean` | r/w | Returns or sets whether a 3-d effect will be employed when displaying the control |
| `drop down lines` | `integer` | r/w | Returns or sets the number of dropdown items. |
| `enabled` | `boolean` | r/w | Returns or sets if the control is enabled |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or sets the height of the control. |
| `left position` | `real` | r/w | Returns or sets the left position of the control |
| `linked cell` | `text` | r/w | Returns or sets reference to a call which contains the value of the control. |
| `list fill range` | `text` | r/w | Returns or sets the items which are contained in the drop down. |
| `list index` | `integer` | r/w | Returns or sets the selected item. |
| `locked` | `boolean` | r/w | Returns or sets whether the control is locked for editing. |
| `name` | `text` | r/w | Returns or sets the name of this control. |
| `number of items in list` | `integer` | r/o | Returns the number of total items in the list. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `top` | `real` | r/w | Returns or sets the position of the top of the control. |
| `top left cell` | `range` | r/o | Returns the top left cell of the range within which the control is positioned. |
| `value` | `integer` | r/w | Returns or sets the current value of the control |
| `visible` | `boolean` | r/w | Returns or sets if the control is visible. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text. |

### `filter`

Represents a filter for a single column.

*Plural:* `filters`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `criteria1` | `any` | r/o | Returns the first filtered value for the specified column in a filtered range. |
| `criteria2` | `any` | r/o | Returns the second filtered value for the specified column in a filtered range. |
| `filter on` | `boolean` | r/o | Returns true if the specified filter is on. |
| `operator` | `XlAutoFilterOperator` | r/w | set or return the operator that associates the two criteria applied by the specified filter. |

### `format color`

Represents a color that a format condition can be colored with.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `integer` | r/w | Returns or sets the primary color of the object. |
| `color index` | `XlColorIndex` | r/w |  |
| `theme color` | `XlThemeColor` | r/w | Returns or sets the theme color in the applied color scheme that is associated with the specified object. |
| `tint and shade` | `real` | r/w | Returns or sets a Single that lightens or darkens a color. |

### `format condition icon object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `format condition icon index` | `integer` | r/o | The index of the icon. Read only. |

### `format condition icon set`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `icon set id` | `XlIconSet` | r/o |  |

### `format condition icon sets`

### `format condition`

Represents a conditional format.

*Plural:* `format conditions`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `applies to` | `range` | r/o | Returns the range this format condition applies to. Read Only. |
| `condition operator` | `XlFormatConditionOperator` | r/o | Returns the operator for the conditional format. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `format condition date operator` | `XlTimePeriods` | r/w |  |
| `format condition priority` | `integer` | r/w | Specifies the priority for the format condition. Read/Write. |
| `format condition text` | `text` | r/w | Returns or sets the text for the specified object. |
| `format condition text operator` | `XlContainsOperator` | r/w |  |
| `format condition type` | `XlFormatConditionType` | r/o | Return the conditional format type. |
| `formula 1` | `text` | r/o | Returns the value or expression associated with the conditional format or data validation. |
| `formula 2` | `text` | r/o | Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is operator between or operator not between. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `number format` | `text` | r/w | Returns or sets the format code for the object. |
| `pivot condition scope type` | `XlPivotConditionScope` | r/w | Returns or sets the part of the pivot table that the pivot table format condition is scoped to |
| `pivot table condition` | `boolean` | r/o | Tells whether this format condition is a pivot table condition. |
| `stop if true` | `boolean` | r/w | Tells whether the processing of format conditions stops if this condition is true. Read/Write. |

### `graphic`

Contains properties that apply to header and footer picture objects.

*Plural:* `graphics`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `brightness` | `real` | r/w | Returns or sets the brightness of the specified picture. The value for this property must be a number from 0.0, dimmest, to 1.0, brightest. |
| `color type` | `MsoPictureColorType` | r/w | Returns or sets the type of color transformation applied to the specified picture. |
| `contrast` | `real` | r/w | Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast, to 1.0, the greatest contrast. |
| `crop bottom` | `real` | r/w | Returns or sets the number of points that are cropped off the bottom of the specified picture. |
| `crop left` | `real` | r/w | Returns or sets the number of points that are cropped off the left side of the specified picture. |
| `crop right` | `real` | r/w | Returns or sets the number of points that are cropped off the right side of the specified picture. |
| `crop top` | `real` | r/w | Returns or sets the number of points that are cropped off the top of the specified picture. |
| `file name` | `text` | r/w | Returns or sets the URL, on the intranet or the Web, or path, local or network, to the location where the specified source object was saved. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `lock aspect ratio` | `boolean` | r/w | Returns or sets if the specified shape retains its original proportions when you resize it. False if you can change the height and width of the shape independently of one another when you resize it. |
| `width` | `real` | r/w | Returns or sets the width of the object. |

### `groupbox`

*Plural:* `groupboxes`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accelerator` | `text` | r/w | Returns or sets the accelerator character for this control. |
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `text` | r/w | Returns or sets the caption for this object. |
| `control text` | `text` | r/w | Returns or sets the default text for the control. |
| `display threeD shading` | `boolean` | r/w |  |
| `enabled` | `boolean` | r/w | Returns or sets if the object is enabled. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or set the height of the object. |
| `left position` | `real` | r/w | Returns or sets the position of the specified object, in points. |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `locked text` | `boolean` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `phonetic accelerator` | `text` | r/w | Returns or sets the link to a range. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `top` | `real` | r/w | Returns or sets the position of the top of the control. |
| `top left cell` | `range` | r/o | Returns the top left cell of the range within which the control is positioned. |
| `visible` | `boolean` | r/w | Returns or sets if the control is visible |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text. |

### `hidden field`

*Plural:* `hidden fields`

### `hidden item`

*Plural:* `hidden items`

### `horizontal page break`

Represents a horizontal page break.

*Plural:* `horizontal page breaks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `extent` | `XlPageBreakExtent` | r/o | Returns the type of the specified page break: full-screen or only within a print area. |
| `horizontal page break type` | `XlPageBreak` | r/w | Returns or sets the type of horizontal page break. |
| `location` | `range` | r/w | Returns or sets where this horizontal page break is located. |

### `hyperlink`

Represents a hyperlink.

*Plural:* `hyperlinks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `address` | `text` | r/w | Returns or sets the address of the target document. |
| `email subject` | `text` | r/w | Returns or sets the text string of the specified hyperlink's e-mail subject line. The subject line is appended to the hyperlink's address. |
| `hyperlink type` | `MsoHyperlinkType` | r/o | Returns the hyperlink type, what the hyperlink is associated with. |
| `name` | `text` | r/o | Returns the name of the object. |
| `range object` | `range` | r/o | Returns a range object that represents the range the specified hyperlink is attached to. |
| `screen tip` | `text` | r/w | Returns or sets the screen tip text for the specified hyperlink. |
| `shape object` | `shape` | r/o | Returns a shape object that represents the shape attached to the specified hyperlink. |
| `sub address` | `text` | r/w | Returns or sets the location within the document associated with the hyperlink. |
| `text to display` | `text` | r/w | Returns or sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink. |

### `icon criteria`

### `icon criterion`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `condition operator` | `XlFormatConditionOperator` | r/w | Returns the operator for the conditional format. |
| `icon criterion icon` | `XlIcon` | r/w |  |
| `icon criterion index` | `integer` | r/o | The index of the icon criterion. Read only. |
| `icon criterion type` | `XlConditionValueTypes` | r/w | Returns or sets the condition value type of the icon criterion. Read/Write. |
| `icon criterion value` | `any` | r/w |  |

### `icon set format condition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `applies to` | `range` | r/o | Returns the range this format condition applies to. Read Only. |
| `format condition icon set` | `format condition icon set` | r/w |  |
| `format condition priority` | `integer` | r/w | Specifies the priority for the format condition. Read/Write. |
| `format condition type` | `XlFormatConditionType` | r/o | Return the conditional format type. |
| `formula` | `text` | r/w | Returns or sets the value or expression associated with the conditional format or data validation. |
| `icon criteria` | `icon criteria` | r/o | Returns the IconCriteria for the IconSetCondition object. |
| `percentile values` | `boolean` | r/w |  |
| `pivot condition scope type` | `XlPivotConditionScope` | r/w | Returns or sets the part of the pivot table that the pivot table format condition is scoped to |
| `pivot table condition` | `boolean` | r/o | Tells whether this format condition is a pivot table condition. |
| `reverse icon set order` | `boolean` | r/w |  |
| `show icon only` | `boolean` | r/w |  |
| `stop if true` | `boolean` | r/o | Tells whether the processing of format conditions stops if this condition is true. Read Only. |

### `international macro sheet`

*Plural:* `international macro sheets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `enable selection` | `XlEnableSelection` | r/w | Returns or sets what can be selected on the sheet. |

### `label`

*Plural:* `labels`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accelerator` | `text` | r/w | Returns or sets the accelerator character for this control. |
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `text` | r/w | Returns or sets the caption for the control. |
| `control text` | `text` | r/w | Returns or sets the default text for the control. |
| `enabled` | `boolean` | r/w | Returns or sets if the control is enabled |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or sets the height of the control. |
| `left position` | `real` | r/w | Returns or sets the left position of the control |
| `locked` | `boolean` | r/w | Returns or sets whether the control is locked for editing. |
| `locked text` | `boolean` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `text` | r/w | Returns or sets the name of this control |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `phonetic accelerator` | `text` | r/w | Returns or sets the link to a range. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `top` | `real` | r/w | Returns or sets the position of the top of the control. |
| `top left cell` | `range` | r/o | Returns the top left cell of the range within which the control is positioned. |
| `visible` | `boolean` | r/w | Returns or sets if the control is visible |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text. |

### `linear gradient`

The LinearGradient object transitions through a series of colors in a linear manner along a specific angle.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `colorstops` | `colorstops` | r/o | Returns the ColorStops for the LinearGradient object. |
| `linear gradient degree` | `real` | r/w | The angle of the linear gradient fill within a selection. |

### `list column`

Represents a column in a list object.

*Plural:* `list columns`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `cell table` | `cell` | r/o | Returns the cell table from a specified list column. |
| `index` | `integer` | r/o |  |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `range object` | `range` | r/o | Returns a range object that represents the range to which the specified list column. |
| `total row` | `range` | r/o | Returns the totals row, if any, from a specified list column. |
| `totals calculation` | `XlTotalsCalculation` | r/w |  |

### `list object`

Represents a list object on a worksheet.

*Plural:* `list objects`

**Contains:** `list column`, `list row`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active` | `boolean` | r/o |  |
| `autofilter object` | `autofilter` | r/o | Returns the autofilter object associated with this list object. |
| `cell table` | `cell` | r/o | Returns the cell table from a specified list object. |
| `comment` | `text` | r/w | Returns or sets the comment of the object. |
| `default` | `text` | r/o |  |
| `display name` | `text` | r/w | Returns or sets the display name of the object. |
| `display right to left` | `boolean` | r/o |  |
| `header row` | `range` | r/o | Returns a range object that represents the used range on the specified worksheet. |
| `insert row` | `range` | r/o |  |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `query table` | `query table` | r/o | Returns a query table object that represents the query table that intersects the specified range. |
| `range object` | `range` | r/o | Returns or sets a range object that represents the range to which the specified list column, object, or row applies. |
| `show autofilter` | `boolean` | r/w | Returns or sets if the autofilter is implemented in a list object. |
| `show headers` | `boolean` | r/w | Returns or sets if headers is implemented in a list object. |
| `show table style column stripes` | `boolean` | r/w | The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns. |
| `show table style first column` | `boolean` | r/w | Returns or sets if the first column is displayed for the specified ListObject object. |
| `show table style last column` | `boolean` | r/w | Returns or sets if the last column is displayed for the specified ListObject object. |
| `show table style row stripes` | `boolean` | r/w | The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows. |
| `sort object` | `sort` | r/o | Returns the sort object associated with this list object. |
| `source type` | `XlListObjectSourceType` | r/o |  |
| `table style` | `null` | r/w | Returns or sets the style used in the table body. |
| `total` | `boolean` | r/w | Returns or sets if a totals row be implemented in a list object. |
| `total row` | `range` | r/o | Returns the totals row, if any, from a specified list object. |

### `list row`

Represents a row in a list object.

*Plural:* `list rows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `index` | `integer` | r/o |  |
| `range object` | `range` | r/o | Returns a range object that represents the range to which the specified list row applies. |

### `listbox`

*Plural:* `listboxes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `display threeD shading` | `boolean` | r/w | Returns or sets whether a 3-d effect will be employed when displaying the control |
| `enabled` | `boolean` | r/w | Returns or sets if the control is enabled |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or sets the height of the control. |
| `left position` | `real` | r/w | Returns or sets the left position of the control |
| `linked cell` | `text` | r/w | Returns or sets reference to a call which contains the value of the control. |
| `list fill range` | `text` | r/w | Returns or sets the items which are contained in the control. |
| `list index` | `integer` | r/w | Returns or sets the selected item. |
| `locked` | `boolean` | r/w | Returns or sets whether the control is locked for editing. |
| `multi select` | `XlMultiSelect` | r/w | Returns or sets the multiple selection setting. |
| `name` | `text` | r/w | Returns or sets the name of this control |
| `number of items in list` | `integer` | r/o |  |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `top` | `real` | r/w | Returns or sets the position of the top of the control. |
| `top left cell` | `range` | r/o | Returns the top left cell of the range within which the control is positioned. |
| `value` | `integer` | r/w | Returns or sets the current value of the control. |
| `visible` | `boolean` | r/w | Returns or sets if the control is visible |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text. |

### `macro sheet`

*Plural:* `macro sheets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `enable selection` | `XlEnableSelection` | r/w | Returns or sets what can be selected on the sheet. |

### `named item`

Represents a defined name for a range of cells. Named items can be either built-in names, such as Database, Print_Area, and Auto_Open, or custom names.

*Plural:* `named items`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `category` | `text` | r/w | Returns or sets the category for the specified name in the language of the macro. |
| `category local` | `text` | r/w | Returns or sets the category for the specified name, in the language of the user, if the name refers to a custom function or command. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `macro type` | `XlXLMMacroType` | r/w | Returns or sets what the name refers to. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `name local` | `text` | r/w | Returns or sets the name of the object, in the language of the user. |
| `reference local` | `text` | r/w | Returns or sets the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign. |
| `reference local r1c1` | `text` | r/w | Returns or sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign. |
| `reference r1c1` | `text` | r/w | Returns or sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign. |
| `reference range` | `range` | r/o | Returns the range object referred to by this object. |
| `references` | `text` | r/w | Returns or sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign. |
| `shortcut key` | `text` | r/w | Returns or sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command. |
| `value` | `text` | r/w | Returns or sets a string containing the formula that the name is defined to refer to. The string is in A1-style notation in the language of the macro, and it begins with an equal sign. |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible. |

### `negative bar format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `databar bar color` | `format color` | r/o |  |
| `databar border color` | `format color` | r/o | Returns the DataBarBorder for the Databar object. |
| `negative bar border color type` | `XlDataBarNegativeColorType` | r/w | Returns or sets the type of border of the databar |
| `negative bar color type` | `XlDataBarNegativeColorType` | r/w | Returns or sets the type of fill of the databar |

### `option button`

*Plural:* `option buttons`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accelerator` | `text` | r/w | Returns or sets the accelerator character for this control. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `text` | r/w | Returns or sets the caption for this object. |
| `control text` | `text` | r/w | Returns or sets the default text for the control. |
| `display threeD shading` | `boolean` | r/w |  |
| `enabled` | `boolean` | r/w | Returns or sets if the object is enabled. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `group box` | `groupbox` | r/o |  |
| `height` | `real` | r/w | Returns or set the height of the object. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the position of the specified object, in points. |
| `linked cell` | `text` | r/w |  |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `locked text` | `boolean` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `phonetic accelerator` | `text` | r/w | Returns or sets the link to a range. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `top` | `real` | r/w | Returns the top position of the specified object, in points. |
| `top left cell` | `range` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `value` | `XlCheckBoxState` | r/w |  |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `outline`

Represents an outline on a worksheet.

*Plural:* `outlines`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `automatic styles` | `boolean` | r/w | Returns or sets if the outline uses automatic styles. |
| `summary column` | `XlSummaryColumn` | r/w | Returns or sets the location of the summary columns in the outline, as shown in the following table. |
| `summary row` | `XlSummaryRow` | r/w | Returns or sets the location of the summary rows in the outline, as shown in the following table. |

### `page field`

*Plural:* `page fields`

### `page setup`

Represents the page setup description. The page setup object contains all page setup attributes, left margin, bottom margin, paper size, and so on, as properties.

*Plural:* `page setups`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `black and white` | `boolean` | r/w | Returns or sets if elements of the document will be printed in black and white. |
| `bottom margin` | `real` | r/w | Returns or sets the size of the bottom margin, in points. |
| `center footer` | `text` | r/w | Returns or sets the center part of the footer. |
| `center footer picture` | `graphic` | r/o | Returns or sets a graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture. |
| `center header` | `text` | r/w | Returns or sets the center part of the header. |
| `center header picture` | `graphic` | r/o | Returns or sets a graphic object that represents the picture for the center section of the footer. Used to set attributes about the picture. |
| `center horizontally` | `boolean` | r/w | Returns or sets  if the sheet is centered horizontally on the page when it's printed. |
| `center vertically` | `boolean` | r/w | Returns or sets if the sheet is centered vertically on the page when it's printed. |
| `chart size` | `XlObjectSize` | r/w | Returns or sets the way a chart is scaled to fit on a page. |
| `draft` | `boolean` | r/w | Returns or sets if the sheet will be printed without graphics. |
| `first page number` | `integer` | r/w | Returns or sets the first page number that will be used when this sheet is printed. |
| `fit to pages tall` | `integer` | r/w | Returns or sets the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets. |
| `fit to pages wide` | `integer` | r/w | Returns or sets the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets. |
| `footer margin` | `real` | r/w | Returns or sets the distance from the bottom of the page to the footer, in points. |
| `header margin` | `real` | r/w | Returns or sets the distance from the top of the page to the header, in points. |
| `left footer` | `text` | r/w | Returns or sets the left part of the footer. |
| `left footer picture` | `graphic` | r/o | Returns or sets a graphic object that represents the picture for the left section of the footer. Used to set attributes about the picture. |
| `left header` | `text` | r/w | Returns or sets the left part of the header. |
| `left header picture` | `graphic` | r/o | Returns or sets a graphic object that represents the picture for the left section of the header. Used to set attributes about the picture. |
| `left margin` | `real` | r/w | Returns or sets the size of the left margin, in points. |
| `order` | `XlOrder` | r/w | Returns or sets the order that Microsoft Excel uses to number pages when printing a large worksheet. |
| `page orientation` | `XlPageOrientation` | r/w | Returns or set if the printing mode will be portrait or landscape. |
| `print Excel comments` | `XlPrintLocation` | r/w | Returns or sets the way comments are printed with the sheet. |
| `print area` | `text` | r/w | Returns or sets the range to be printed, as a string using A1-style references in the language of the macro. |
| `print gridlines` | `boolean` | r/w | Returns or sets if cell gridlines are printed on the page. Applies only to worksheets. |
| `print headings` | `boolean` | r/w | Returns or sets if row and column headings are printed with this page. Applies only to worksheets. |
| `print notes` | `boolean` | r/w | Returns or sets if cell notes are printed as end notes with the sheet. Applies only to worksheets. |
| `print quality` | `any` | r/w | Set/Gets a two element list where 1 is the horizontal print quality and 2 is the vertical print quality |
| `print title columns` | `text` | r/w | Returns or sets the columns that contain the cells to be repeated on the left side of each page, as a string in A1-style notation in the language of the macro. |
| `print title rows` | `text` | r/w | Returns or sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro. |
| `right footer` | `text` | r/w | Returns or sets the right part of the footer. |
| `right footer picture` | `graphic` | r/o | Returns or sets a graphic object that represents the picture for the right section of the footer. Used to set attributes about the picture. |
| `right header` | `text` | r/w | Returns or sets the right part of the header. |
| `right header picture` | `graphic` | r/o | Returns or sets a graphic object that represents the picture for the right section of the header. Used to set attributes about the picture. |
| `right margin` | `real` | r/w | Returns or sets the size of the right margin, in points. |
| `top margin` | `real` | r/w | Returns or sets the size of the top margin, in points. |
| `zoom` | `XLSourceData` | r/w | Returns or sets a percentage, between 10 and 400 percent, by which Microsoft Excel will scale the worksheet for printing. Applies only to worksheets. |

### `pane`

Represents a pane of a window. Pane objects exist only for worksheets and Microsoft Excel 4.0 macro sheets.

*Plural:* `panes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `scroll column` | `integer` | r/w | Returns or sets the number of the leftmost column in the pane |
| `scroll row` | `integer` | r/w | Returns or sets the number of the row that appears at the top of the pane. |
| `visible range` | `range` | r/o | Returns a Range object that represents the range of cells that are visible in the pane. If a column or row is partially visible, it's included in the range. |

### `parent item`

*Plural:* `parent items`

### `phonetic`

Contains information about a specific phonetic text string in a cell.

*Plural:* `phonetics`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `character type` | `XlPhoneticCharacterType` | r/w | Returns or sets the type of phonetic text in the specified cell. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `phonetic alignment` | `XlPhoneticAlignment` | r/w | Returns or sets the alignment for the specified phonetic text. |
| `phonetic text` | `text` | r/w | Returns or sets the text for the specified object. |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible |

### `pivot axis`

Used as the base class for the PivotDataAxis, PivotFilterAxis, and PivotGroupAxis objects. There are no properties or methods which with you can use the PivotAxis object directly.

**Contains:** `pivot line`

### `pivot cache`

Represents the memory cache for a PivotTable report.

*Plural:* `pivot caches`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `ADO connection` | `null` | r/o | Returns an ADO connection object if the PivotTable cache is connected to an OLE DB data source. |
| `OLAP` | `boolean` | r/o | Returns True if the PivotTable cache is connected to an Online Analytical Processing server. |
| `SQL query` | `text` | r/w | Returns or sets the SQL query string used with the specified ODBC data source. |
| `background query` | `boolean` | r/w | Returns or sets if queries for the pivot table report or query table are performed asynchronously, in the background. |
| `command text` | `text` | r/w | Returns or sets the command string for the specified data source. |
| `command type` | `XlCmdType` | r/w | Returns or sets a constant describing the value type of the CommandText property. |
| `connection` | `text` | r/w | Returns or sets a string that contains one of the following. ODBC settings that enable Microsoft Excel to connect to an ODBC data source, a URL that enables Microsoft Excel to connect to a Web data source, or a file that specifies a database or Web query. |
| `enable refresh` | `boolean` | r/w | Returns or sets if the pivot table cache can be refreshed by the user. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `is connected` | `boolean` | r/o | Returns True if the MaintainConnection property is True and the PivotTable cache is currently connected to its source. |
| `local connection` | `text` | r/w | Returns or sets the connection string to an offline cube file. |
| `maintain connection` | `boolean` | r/w | True if the connection to the specified data source is maintained after the refresh and until the workbook is closed. |
| `memory used` | `integer` | r/o | Returns the amount of memory currently being used by the object, in bytes. |
| `missing items limit` | `XlPivotTableMissingItems` | r/w | Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records. |
| `optimize cache` | `boolean` | r/w | Returns or set if the pivot table cache is optimized when it's constructed. |
| `query type` | `XlQueryType` | r/o | Indicates the type of query used by Microsoft Excel to populate the PivotTable cache. |
| `record count` | `integer` | r/o | Returns the number of records in the pivot table cache or the number of cache records that contain the specified item. |
| `refresh date` | `date` | r/o | Returns the date on which the pivot cache was last refreshed. |
| `refresh name` | `text` | r/o | Returns the name of the person who last refreshed the pivot cache. |
| `refresh on file open` | `boolean` | r/w | Returns or sets if the pivot table cache or query table is automatically updated each time the workbook is opened. |
| `refresh period` | `integer` | r/w | Returns or sets the number of minutes between refreshes. |
| `robust connect` | `XlRobustConnect` | r/w | Returns or sets how the PivotTable cache connects to its data source. |
| `save password` | `boolean` | r/w | Returns or set if password information in an ODBC connection string is saved with the specified query. if false, the password is removed. |
| `source connection file` | `text` | r/w | Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable. |
| `source data` | `XLSourceDataLocation` | r/w | Returns or sets the data source for the pivot table. |
| `source type` | `XlPivotTableSourceType` | r/o | Returns a value that identifies the type of item being published. |
| `upgrade on refresh` | `boolean` | r/w | Contains information on whether to upgrade the PivotCache and all connected PivotTables on the next refresh. |
| `use local connection` | `boolean` | r/w | True if the LocalConnection property is used to specify the string that enables Microsoft Excel to connect to a data source. |
| `version` | `XlPivotTableVersionList` | r/o | Returns the version of Microsoft Excel in which the PivotCache was created. |
| `workbook connection` | `null` | r/o | Establishes a connection between the current workbook and the PivotCache object. |

### `pivot cell`

Represents a cell in a PivotTable report.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `MDX` | `text` | r/o |  |
| `cell changed` | `XlCellChangedState` | r/o |  |
| `custom subtotal function` | `XlConsolidationFunction` | r/o | Returns the custom subtotal function field setting of a PivotCell object. |
| `data field` | `pivot field` | r/o | Returns a PivotField object that corresponds to the selected data field. |
| `data source value` | `null` | r/o |  |
| `pivot cell type` | `XlPivotCellType` | r/o | Returns a constant that identifies the PivotTable entity the cell corresponds to. |
| `pivot column line` | `pivot line` | r/o | Returns the PivotLine on a column for a specific PivotCell object. |
| `pivot field` | `pivot field` | r/o | Returns a PivotField object that represents the PivotTable field containing the upper-left corner of the specified range. |
| `pivot item` | `pivot item` | r/o | Returns a PivotItem object that represents the PivotTable item containing the upper-left corner of the specified range. |
| `pivot row line` | `pivot line` | r/o | Returns the PivotLine on a row for a specific PivotCell object. |
| `pivot table` | `pivot table` | r/o | Returns a PivotTable object that represents the PivotTable report associated with the PivotCell. |
| `range` | `range` | r/o | Returns a Range object that represents the range the specified PivotCell applies to. |
| `row items` | `null` | r/w | Returns a PivotItemList collection that corresponds to the items on the category axis that represent the selected cell. |

### `pivot field`

Represents a field in a pivot table.

*Plural:* `pivot fields`

**Contains:** `child item`, `hidden item`, `parent item`, `pivot item`, `calculated item`, `pivot filter`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `all items visible` | `boolean` | r/o | Used to retrieve a Boolean value that indicates whether or not any manual filtering is applied to the PivotField. |
| `auto show count` | `integer` | r/o | Returns the number of top or bottom items that are automatically shown in the pivot field. |
| `auto show field` | `text` | r/o | Returns the name of the data field used to determine the top or bottom items that are automatically shown in the pivot field. |
| `auto show range` | `XLAutoShowPosition` | r/o | Returns position top if the top items are shown automatically in the pivot field, returns position bottom if the bottom items are shown. |
| `auto show type` | `XLAutoShowType` | r/o | Returns type_automatic if auto show is enabled for the pivot field, or  returns type_manual if auto show is disabled. |
| `auto sort custom subtotal` | `integer` | r/o | Returns the name of the custom subtotal used to sort the specified PivotTable field automatically. |
| `auto sort field` | `text` | r/o | Returns the name of the data field used to sort the pivot field automatically. |
| `auto sort order` | `XlSortOrder` | r/o | Returns the order used to sort the pivot field automatically. |
| `auto sort pivot line` | `pivot line` | r/o | Returns the name of the PivotLine used to sort the specified PivotTable field automatically. |
| `base field` | `text` | r/w | Returns or sets the base field for a custom calculation. |
| `base item` | `text` | r/w | Returns or sets the item in the base field for a custom calculation. |
| `calculation` | `XlPivotFieldCalculation` | r/w | Returns or sets the type of calculation done by the specified pivot field. |
| `caption` | `text` | r/w | The label text for the pivot field. |
| `child field` | `pivot field` | r/o | Returns a pivot field object that represents the child pivot field for the specified field, if the field is grouped and has a child field. |
| `cube field` | `cube field` | r/o | Returns the CubeField object from which the specified PivotTable field is descended. |
| `current page` | `null` | r/w | Returns or sets the current page showing for the page field, only valid for page fields. |
| `current page list` | `null` | r/w | Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report. |
| `current page name` | `text` | r/w | Returns or sets the currently displayed page of the specified PivotTable report. |
| `data range` | `range` | r/o | Returns a range object.  For a data field the range is the data contained in the field, for a row, column or page field it is the items in the field. |
| `database sort` | `boolean` | r/w | When set to True, manual repositioning of items in a PivotTable field is allowed. |
| `display as caption` | `boolean` | r/o | This property is used to display member properties of PivotFields as captions. |
| `display as tooltip` | `boolean` | r/w | This property is used to specify whether or not a specific member property PivotField is displayed in tooltips. |
| `display in report` | `boolean` | r/w | This property is used to specify whether the specified member property PivotField is displayed in the PivotTable or not. |
| `drag to column` | `boolean` | r/w | Returns or sets if the pivot field can be dragged to the column position. |
| `drag to data` | `boolean` | r/w | Returns or sets if the field can be dragged to the data position. |
| `drag to hide` | `boolean` | r/w | Returns or sets if the field can be hidden by being dragged off the pivot table. |
| `drag to page` | `boolean` | r/w | Returns or sets if the field can be dragged to the page position. |
| `drag to row` | `boolean` | r/w | Returns or sets if the field can be dragged to the row position. |
| `drilled down` | `boolean` | r/w | True if the flag for the specified PivotTable field or PivotTable item is set to drilled. |
| `enable item selection` | `boolean` | r/w | When set to False, disables the ability to use the field dropdown in the user interface. |
| `enable multiple page items` | `boolean` | r/w | Used for specifying whether or not check boxes are present in the filter drop-down list for fields in the page area. |
| `formula` | `text` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `function` | `XlConsolidationFunction` | r/w | Returns or sets the function used to summarize the pivot field. |
| `group level` | `integer` | r/o | Returns the placement of the specified field within a group of fields, if the field is a member of a grouped set of fields. |
| `hidden` | `boolean` | r/w | This property is used to hide the individual levels of an OLAP hierarchy. |
| `hidden items list` | `null` | r/w | Returns or sets an Object specifying an array of strings that are hidden items for a PivotTable field. |
| `include new items in filter` | `boolean` | r/w | Allows developers to specify whether excluded or included items should be tracked when manual filtering is applied to the PivotField. |
| `is calculated` | `boolean` | r/o | Returns true if the pivot field or item is a calculated field or item. |
| `is member property` | `boolean` | r/o | Returns True when the PivotField contains member properties. |
| `label range` | `range` | r/o | Returns a range object that represents the cell, or cells, that contain the field label. |
| `layout blank line` | `boolean` | r/w | True if a blank row is inserted after the specified row field in a PivotTable report. |
| `layout compact row` | `boolean` | r/w | Specifies whether or not a PivotField is compacted , items of multiple PivotFields are displayed in a single column, when rows are selected. |
| `layout form` | `XlLayoutFormType` | r/w | Returns or sets the way the specified PivotTable items appear. |
| `layout pagebreak` | `boolean` | r/w | True if a page break is inserted after each field. |
| `layout subtotal location` | `XlSubtototalLocationType` | r/w | Returns or sets the position of the PivotTable field subtotals in relation to either above or below the specified field. |
| `member property caption` | `text` | r/w | Setting the MemberPropertyCaption property controls which member property is used as caption for a given level. |
| `memory used` | `integer` | r/o | Returns the amount of memory currently being used by the object, in bytes. |
| `name` | `text` | r/w | Returns or sets the name of the pivot field. |
| `number format` | `text` | r/w | Returns or sets the format code for the pivot field. |
| `parent field` | `pivot field` | r/o | Returns a pivot field object that represents the pivot field that's the group parent of the object. The field must be grouped and have a parent field. |
| `pivot field data type` | `XlPivotFieldDataType` | r/o | Returns an enumeration describing the type of data in the pivot field. |
| `pivot field orientation` | `XlPivotFieldOrientation` | r/w | The location of the field in the pivot table. |
| `position` | `integer` | r/w | Position of the field, first, second, third, and so on, among all the fields in its orientation, rows, columns, pages, data. |
| `property order` | `integer` | r/w | Valid only for PivotTable fields that are member property fields. |
| `property parent field` | `pivot field` | r/o | Returns a PivotField object representing the field to which the properties in this field pertain. |
| `repeat labels` | `boolean` | r/w |  |
| `server based` | `boolean` | r/w | Returns or sets if the pivot table's data source is external and only the items matching the page field selection are retrieved. |
| `show all items` | `boolean` | r/w | Returns or sets if all items in the pivot table are displayed, even if they don't contain summary data. |
| `show detail` | `boolean` | r/w | Gets or sets whether the specified PivotField is showing detail. |
| `showing in axis` | `boolean` | r/o | Indicates if the PivotField is currently visible in the PivotTable or not. |
| `source caption` | `text` | r/o | The SourceCaption property is applicable only for OLAP PivotTables, and returns the original caption from the OLAP server for a PivotField. |
| `source name` | `text` | r/o | Return the specified object's name, as it appears in the original source data for the pivot table. This might be different from the current item name if the user renamed the item after creating the pivot table. |
| `standard formula` | `text` | r/w | Returns or sets a String specifying formulas with standard English formatting. |
| `subtotal name` | `text` | r/w | Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report. |
| `total levels` | `integer` | r/o | Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based, TotalLevels returns the value 1. |
| `use member property as caption` | `boolean` | r/w | This property is used to control whether member property captions are used for PivotItem captions of the PivotField. |
| `value` | `text` | r/w | Returns or sets the name of the specified field in the pivot table. |
| `visible items` | `specifier` | r/o | Returns a list of all the visible pivot items in the specified field. |
| `visible items list` | `null` | r/w | Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report. |

### `pivot filter`

PivotTable Advanced Filter.

*Plural:* `pivot filters`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active` | `boolean` | r/o | Returns whether the specified PivotFilter is active. |
| `data cube field` | `cube field` | r/o | This property is applicable only to non-OLAP PivotTables and provides the Value field ,PivotField in the Values area, being filtered by for a value filter. |
| `data field` | `pivot field` | r/o | This property is applicable only to non-OLAP PivotTables and provides the Value field ,PivotField in the Values area, being filtered by for a value filter. |
| `description` | `text` | r/o | Provides an optional description for the PivotFilter object. |
| `filter type` | `XlPivotFilterType` | r/o | Specifies the type of filter to be applied. |
| `is member property filter` | `boolean` | r/o | Specifies whether the label filter is based on the PivotItem captions of a member property of the field or on the PivotItem captions of the PivotField itself. |
| `member property field` | `pivot field` | r/o | This property specifies the member property PivotField on which the label filter is based. |
| `name` | `text` | r/o | This property provides the option of naming filters for reference. |
| `order` | `integer` | r/w | Specifies the evaluation order of the filter among all Value filters applied to the entire PivotTable. |
| `pivot field` | `pivot field` | r/o | Specifies the PivotField to which the filter is applied. |
| `value1` | `null` | r/o | This property is a user-supplied parameter to define a filter for a PivotField. |
| `value2` | `null` | r/o | This property is a user-supplied parameter to define a filter for a PivotField. |

### `pivot formula`

Represents a formula used to calculate results in a pivot table report.

*Plural:* `pivot formulas`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/w | Returns the index number of the object within the elements of the parent object. |
| `formula` | `text` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `standard formula` | `text` | r/w | Returns or sets a String specifying formulas with standard English formatting. |
| `value` | `text` | r/w | Returns or sets the name of the specified formula in the pivot table formula. |

### `pivot item`

Represents an item in a pivot field. The items are the individual data entries in a field category.

*Plural:* `pivot items`

**Contains:** `child item`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `caption` | `text` | r/w | The label text for the pivot item. |
| `data range` | `range` | r/o | Returns a range object specifying the data qualified by the item. |
| `drilled down` | `boolean` | r/w | True if the flag for the specified PivotTable field or PivotTable item is set to expanded, or visible. |
| `formula` | `text` | r/w | Returns or sets the pivot item's formula, in A1-style notation and in the language of the macro. |
| `is calculated` | `boolean` | r/o | Returns true if the pivot  item is a calculated item. |
| `label range` | `range` | r/o | Returns a range object that represents all the pivot table cells that contain the item. |
| `name` | `text` | r/w | Returns or sets the name of the pivot item. |
| `parent item` | `parent item` | r/o | Returns a pivot item object that represents the parent pivot item in the parent pivot field object, the field must be grouped so that it has a parent. |
| `parent show detail` | `boolean` | r/o | Return true if the specified item is showing because one of its parents is showing detail. False if the specified item isn't showing because one of its parents is hiding detail. This property is available only if the item is grouped. |
| `position` | `integer` | r/w | Returns or sets the position of the item in its field, if the item is currently showing. |
| `record count` | `integer` | r/o | Returns the number of records in the pivot table cache or the number of cache records that contain the specified item. |
| `show detail` | `boolean` | r/w | Returns  or sets if the pivot item is showing detail. |
| `source name` | `text` | r/o | Return the specified object's name, as it appears in the original source data for the pivot table. This might be different from the current item name if the user renamed the item after creating the pivot table. |
| `source name standard` | `text` | r/o | Returns a String that represents the PivotTable item's source name in standard English format settings. |
| `standard formula` | `text` | r/w | Returns or sets a String specifying formulas with standard English formatting. |
| `value` | `text` | r/w | Returns or sets set the name of the specified item in the pivot table field. |
| `visible` | `boolean` | r/w | Returns or sets if the pivot item is visible. |

### `pivot line`

The PivotLines object is a collection of lines in a PivotTable, containing all lines on rows or columns of the pivot.

*Plural:* `pivot lines`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `line type` | `XlPivotLineType` | r/o | Returns an XlPivotLineType constant that indicates the type of PivotLine. |
| `pivot line cells` | `null` | r/o | Returns a collection of PivotCell objects in a PivotLine. |
| `position` | `integer` | r/o | Returns or sets the position of the PivotLine object. |

### `pivot table`

Represents a pivot table on a worksheet.

*Plural:* `pivot tables`

**Contains:** `column field`, `data field`, `hidden field`, `page field`, `pivot field`, `row field`, `calculated field`, `calculated member`, `active filter`, `cube field`, `slicer`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `CompactRowIndent` | `integer` | r/w | Returns or sets the indent increment for PivotItems when compact row layout form is turned on. |
| `allocation` | `XlAllocation` | r/w |  |
| `allocation method` | `XlAllocationMethod` | r/w |  |
| `allocation value` | `XlAllocationValue` | r/w |  |
| `allocation weight expression` | `text` | r/w |  |
| `allow multiple filters` | `boolean` | r/w | Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time. |
| `alternative text` | `text` | r/w | Returns or sets the descriptive alternative text string for a ShapeRange object when the object is saved to a Web page. |
| `cache index` | `integer` | r/w | Returns or sets the index number of the pivot table cache. |
| `calculated members in filters` | `boolean` | r/w |  |
| `change list` | `null` | r/o |  |
| `column grand` | `boolean` | r/w | Returns or sets if the pivot table shows grand totals for columns. |
| `column range` | `range` | r/o | Returns a range object that represents the range that contains the pivot table column area. |
| `compact layout column header` | `text` | r/w | Specifies the caption that is displayed in the column header of a PivotTable when in compact row layout form. |
| `compact layout row header` | `text` | r/w | Specifies the caption that is displayed in the row header of a PivotTable when in compact row layout form. |
| `data body range` | `range` | r/o | Returns a range object that represents the range that contains the pivot table data area. |
| `data label range` | `range` | r/o | Returns a range object that represents the range that contains the labels for the pivot table data fields. |
| `data pivot field` | `pivot field` | r/o | Returns a PivotField object that represents all the data fields in a PivotTable. |
| `display context tooltips` | `boolean` | r/w | Controls whether or not tooltips are displayed for PivotTable cells. |
| `display empty column` | `boolean` | r/w | Returns True when the nonempty MDX keyword is included in the query to the OLAP provider for the value axis. |
| `display empty row` | `boolean` | r/w | Returns True when the nonempty MDX keyword is included in the query to the OLAP provider for the category axis. |
| `display error string` | `boolean` | r/w | Returns or sets if the pivot table displays a custom error string in cells that contain errors. |
| `display field captions` | `boolean` | r/w | Controls whether or not filter buttons and PivotField captions for rows and columns are displayed in the grid. |
| `display immediate items` | `boolean` | r/w | Returns or sets a Boolean that indicates whether items in the row and column areas are visible when the data area of the PivotTable is empty. |
| `display member property tooltips` | `boolean` | r/w | Controls whether or not to display member properties in tooltips. |
| `display null string` | `boolean` | r/w | Returns or sets if the pivot table displays a custom string in cells that contain null values. |
| `enable data value editing` | `boolean` | r/w | True to disable the alert for when the user overwrites values in the data area of the PivotTable. |
| `enable drilldown` | `boolean` | r/w | Returns or sets if drilldown is enabled. |
| `enable field dialog` | `boolean` | r/w | Returns or sets if the pivot table field dialog box is available when the user double-clicks the pivot table field. |
| `enable field list` | `boolean` | r/w | False to disable the ability to display the field well for the PivotTable. |
| `enable wizard` | `boolean` | r/w | Returns or sets if the pivot table wizard is available. |
| `enable writeback` | `boolean` | r/w |  |
| `error string` | `text` | r/w | Returns or sets the string displayed in cells that contain errors when the display error string property is true. The default value is an empty string. |
| `file list sort ascending` | `boolean` | r/w | Controls the sort order of fields in the PivotTable Field List. |
| `grand total name` | `text` | r/w | Returns or sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report |
| `has autoformat` | `boolean` | r/w | Returns or sets if the pivot table is automatically formatted when it's refreshed or when fields are moved. |
| `in grid drop zones` | `boolean` | r/w | This property is used to toggle in-grid drop zones for a PivotTable object. |
| `inner detail` | `text` | r/w | Returns or sets the name of the field that will be shown as detail when the show detail property is true for the innermost row or column field. |
| `layout row default` | `XlLayoutRowType` | r/w | This property specifies the layout settings for PivotFields when they are added to the PivotTable for the first time. |
| `location` | `text` | r/w | Gets or sets a String that represents the top-left cell in the body of the specified PivotTable. |
| `manual update` | `boolean` | r/w | Returns or sets if the pivot table is recalculated only at the user's request. |
| `mdx` | `text` | r/o | Returns a String indicating the MDX, Multidimensional-Expression, that would be sent to the provider to populate the current PivotTable view. |
| `merge labels` | `boolean` | r/w | Returns or sets if pivot table outer-row item, column item, subtotal, and grand total labels use merged cells. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `null string` | `text` | r/w | Returns or sets the string displayed in cells that contain null values when the display null string property is true. The default value is an empty string. |
| `page field order` | `XlOrder` | r/w | Returns or sets the order in which page fields are added to the pivot table layout. |
| `page field style` | `text` | r/w | Returns or sets the style used in the bound page field area. |
| `page field wrap count` | `integer` | r/w | Returns or sets the number of pivot table page fields in each column or row. |
| `page range` | `range` | r/o | Returns a range object that represents the range that contains the pivot table page area. |
| `page range cells` | `range` | r/o | Returns a range object that represents the cells in the pivot table containing only the page fields and item drop-down lists. |
| `pivot cache` | `pivot cache` | r/o | Returns a pivot cache object that represents the cache for the specified pivot table |
| `pivot column axis` | `pivot axis` | r/o | Returns a PivotAxis object representing the entire column axis |
| `pivot row axis` | `pivot axis` | r/o | Returns a PivotAxis object representing the entire row axis |
| `pivot selection` | `text` | r/w | Returns or sets the pivot table selection, in standard pivot table selection format. |
| `pivot selection standard` | `text` | r/w | Returns or sets a String indicating the PivotTable selection in standard PivotTable report format using English settings. |
| `preserve formatting` | `boolean` | r/w | Returns or sets if pivot table formatting is preserved when the pivot table is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items. |
| `print drill indicators` | `boolean` | r/w | Specifies whether or not drill indicators are printed with the PivotTable. |
| `print titles` | `boolean` | r/w | set/get whether the print titles for the worksheet are set based on the PivotTable report. |
| `refresh date` | `date` | r/o | Returns the date on which the pivot table was last refreshed. |
| `refresh name` | `text` | r/o | Returns the name of the person who last refreshed the pivot table data. |
| `repeat items on each printed page` | `boolean` | r/w | True if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. False if labels are printed only on the first page. |
| `row grand` | `boolean` | r/w | Returns or sets if the pivot table shows grand totals for rows. |
| `row range` | `range` | r/o | Returns a range object that represents the range including the pivot table row area. |
| `save data` | `boolean` | r/w | Returns or sets if data for the pivot table is saved with the workbook. If false only the pivot table definition is saved. |
| `selection mode` | `XlPTSelectionMode` | r/w | Returns or sets the pivot table structured selection mode. |
| `show drill indicators` | `boolean` | r/w | When the ShowDrillIndicators property is set to False, the PrintDrillIndicators property has no effect. |
| `show page multiple label` | `boolean` | r/w | When set to True, Multiple Items will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of nonhidden items is shown in the PivotTable view. |
| `show table style column headers` | `boolean` | r/w | The ShowTableStyleColumnHeaders property is set to True if the coulmn headers should be displayed in the PivotTable. |
| `show table style column stripes` | `boolean` | r/w | The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns. |
| `show table style last column` | `boolean` | r/w | Returns or sets if the last column is displayed for the specified ListObject object. |
| `show table style row headers` | `boolean` | r/w | The ShowTableStyleRowHeaders property is set to True if the row headers should be displayed in the PivotTable. |
| `show table style row stripes` | `boolean` | r/w | The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows. |
| `show values row` | `boolean` | r/w |  |
| `small grid` | `boolean` | r/w | Returns or sets if Microsoft Excel uses a grid that's two cells wide and two cells deep for a newly created pivot table report. False if Excel uses a blank stencil outline. |
| `sort using custom lists` | `boolean` | r/w | The SortUsingCustomLists property controls whether custom lists are used for sorting items of fields. |
| `source data` | `XLSourceDataLocation` | r/w | Either a range reference as a string or a list of SQL commands |
| `subtotal hidden page items` | `boolean` | r/w | Returns or sets if hidden page field items in the pivot table are included in row and column subtotals, block totals, and grand totals. |
| `summary` | `text` | r/w |  |
| `table range1` | `range` | r/o | Returns a range object that represents the range containing the entire pivot table, but doesn't include page fields. |
| `table range2` | `range` | r/o | Returns a range object that represents the range containing the entire pivot table, including page fields. |
| `table style` | `text` | r/w | Returns or sets the style used in the pivot table body. |
| `table style2` | `null` | r/w | The TableStyle2 property specifies the PivotTable style currently applied to the PivotTable. |
| `tag` | `text` | r/w | Returns or sets a string saved with the pivot table. |
| `totals annotation` | `boolean` | r/w | True if an asterisk is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source. |
| `vacated style` | `text` | r/w | Returns or sets the style applied to cells vacated when the pivot table is refreshed. |
| `value` | `text` | r/w | Returns or set the name of the pivot table. |
| `version` | `XlPivotTableVersionList` | r/o |  |
| `view calculated members` | `boolean` | r/w | When set to True, calculated members for Online Analytical Processing, OLAP, PivotTables can be viewed. |
| `visual totals` | `boolean` | r/w | True to enable Online Analytical Processing, OLAP, PivotTables to retotal after an item has been hidden from view. |
| `visual totals for sets` | `boolean` | r/w |  |

### `query table`

Represents a worksheet table built from data returned from an external data source, such as an SQL server.

*Plural:* `query tables`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `adjust column width` | `boolean` | r/w | Returns or sets if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. |
| `background query` | `boolean` | r/w | Returns or sets if queries for the query table are performed asynchronously. |
| `command type` | `XlCmdType` | r/w | Returns or sets one of the XlCmdType constants listed in the following table in the remarks section. |
| `connection` | `text` | r/w | Returns or sets a string that contains one of the following: ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; or a file that specifies a database or Web query. |
| `destination` | `range` | r/o | Returns the cell in the upper-left corner of the query table destination range, the range where the resulting query table will be placed. The destination range must be on the worksheet that contains the query table object. |
| `enable editing` | `boolean` | r/w | Returns or sets if the user can edit the specified query table. False if the user can only refresh the query table. |
| `enable refresh` | `boolean` | r/w | Returns or sets if the query table can be refreshed by the user. |
| `fetched row overflow` | `boolean` | r/o | Returns true if the number of rows returned by the last use of the refresh method is greater than the number of rows available on the worksheet. |
| `field names` | `boolean` | r/w | Returns or sets if field names from the data source appear as column headings for the returned data. |
| `fill adjacent formulas` | `boolean` | r/w | Returns or sets if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed. |
| `has autoformat` | `boolean` | r/w | Returns or sets if the query table is automatically formatted when it's refreshed or when fields are moved. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `post text` | `text` | r/w | Returns or sets the string used with the post method of inputting data into a Web server to return data from a Web query. |
| `query type` | `XlQueryType` | r/o | Returns the type of query used by Microsoft Excel to populate the query table. |
| `refresh on file open` | `boolean` | r/w | Returns or sets if the query table is automatically updated each time the workbook is opened. |
| `refresh style` | `XlCellInsertionMode` | r/w | Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a recordset returned by a query. |
| `refreshing` | `boolean` | r/o | Returns true if there's a background query in progress for the specified query table. |
| `result range` | `range` | r/o | Returns a range object that represents the area of the worksheet occupied by the specified query table. |
| `row numbers` | `boolean` | r/w | Returns or sets if row numbers are added as the first column of the specified query table. |
| `save data` | `boolean` | r/w | Returns or sets if data for the query table is saved with the workbook. |
| `save password` | `boolean` | r/w | Returns or sets if password information in an ODBC connection string is saved with the specified query. |
| `sql` | `text` | r/w | Returns or sets the SQL query string used with the specified ODBC data source. |
| `tables only from html` | `boolean` | r/w | Returns or sets if only the HTML tables in the document are read when a query table is refreshed. False if the entire HTML document is read when a query table is refreshed. This property has an effect only when the query table is using a URL connection. |
| `text file column data types` | `XlColumnDataType` | r/w | A list of enums: general format, text format, MDY format, DMY format, YMD format, MYD format, DYM format, YDM format, skip column |
| `text file comma delimiter` | `boolean` | r/w | Returns or sets if the comma character is the delimiter when you import a text file into a query table. |
| `text file consecutive delimiter` | `boolean` | r/w | Returns or sets if consecutive delimiters are treated as a single delimiter when you import a text file into a query table. |
| `text file decimal separator` | `text` | r/w | Returns or sets the decimal separator character that Microsoft Excel uses when you import a text file into a query table. The default is the system decimal separator character. |
| `text file fixed column widths` | `integer` | r/w | Returns or sets a list of numbers that correspond to the widths of the columns in the text file that you are importing into a query table. Valid widths are from 1 through 32767 characters. |
| `text file other delimiter` | `text` | r/w | Returns or sets the character used as the delimiter when you import a text file into a query table. |
| `text file parse type` | `XlTextParsingType` | r/w | Returns or sets the column format for the data in the text file that you're importing into a query table. |
| `text file platform` | `integer` | r/w | Returns or sets the origin of the text file you're importing into the query table. This property determines which code page is used during the data import. |
| `text file prompt on refresh` | `boolean` | r/w | Returns or sets if you want to specify the name of the imported text file each time the query table is refreshed. |
| `text file semicolon delimiter` | `boolean` | r/w | Returns or sets if the semicolon character is the delimiter when you import a text file into a query table. |
| `text file space delimiter` | `boolean` | r/w | Returns or sets if the space character is the delimiter when you import a text file into a query table. |
| `text file start row` | `integer` | r/w | Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1. |
| `text file tab delimiter` | `boolean` | r/w | Returns or sets if the tab character is the delimiter when you import a text file into a query table. |
| `text file text qualifier` | `XlTextQualifier` | r/w | Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format. |
| `text file thousands separator` | `text` | r/w | Returns or sets the thousands separator character thatMicrosoft Excel uses when you import a text file into a query table. The default is the system thousands separator character. |

### `recent file`

Represents a file in the list of recently used files.

*Plural:* `recent files`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `name` | `text` | r/o | Returns the name of the object. |
| `path` | `text` | r/o | Returns the complete path to the file, excluding the final separator and name of the file. |

### `rectangular gradient`

The RectangularGradient object transitions through a series of colors from the center to one or more sides.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `colorstops` | `colorstops` | r/o | Returns the ColorStops for the LinearGradient object. |
| `rectangular gradient bottom` | `real` | r/w | Represents the point or vector that the gradient fill converges to. |
| `rectangular gradient left` | `real` | r/w | Represents the point or vector that the gradient fill converges to. |
| `rectangular gradient right` | `real` | r/w | Represents the point or vector that the gradient fill converges to. |
| `rectangular gradient top` | `real` | r/w | Represents the point or vector that the gradient fill converges to. |

### `row field`

*Plural:* `row fields`

### `row item`

*Plural:* `row items`

### `scenario`

Represents a scenario on a worksheet. Represents a scenario on a worksheet. A scenario is a group of input values, called changing cells, that's named and saved.

*Plural:* `scenarios`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `Excel comment` | `text` | r/w | Returns or sets the comment associated with the scenario. The comment text cannot exceed 255 characters. |
| `changing cells` | `range` | r/o | Returns a range object that represents the changing cells for a scenario. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `hidden` | `boolean` | r/w | Returns or sets if the scenario is hidden. |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked. |
| `name` | `text` | r/w | Returns or sets the name of the object. |

### `scrollbar`

*Plural:* `scrollbars`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `display threeD shading` | `boolean` | r/w |  |
| `enabled` | `boolean` | r/w | Returns or sets if the object is enabled. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or set the height of the object. |
| `large change` | `integer` | r/w |  |
| `left position` | `real` | r/w | Returns or sets the position of the specified object, in points. |
| `linked cell` | `text` | r/w |  |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `maximum value` | `integer` | r/w |  |
| `minimum value` | `integer` | r/w |  |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `small change` | `integer` | r/w |  |
| `top` | `real` | r/w | Returns the top position of the specified object, in points. |
| `top left cell` | `range` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `value` | `integer` | r/w |  |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `sheet`

Represents a worksheet.

*Plural:* `sheets`

**Contains:** `shape`, `arc`, `button`, `chart object`, `checkbox`, `dropdown`, `groupbox`, `label`, `line`, `listbox`, `named item`, `option button`, `oval`, `pivot table`, `range`, `cell`, `row`, `column`, `rectangle`, `scenario`, `scrollbar`, `spinner`, `textbox`, `horizontal page break`, `vertical page break`, `query table`, `Excel comment`, `hyperlink`, `list object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `autofilter mode` | `boolean` | r/w | Returns or sets if the autofilter drop-down arrows are currently displayed on the sheet. This property is independent of the filter mode property. |
| `autofilter object` | `autofilter` | r/o | Returns the autofilter object associated with this sheet. |
| `circular reference` | `range` | r/o | Returns a range object that represents the range containing the first circular reference on the sheet, or returns missing value if there's no circular reference on the sheet. |
| `consolidation function` | `XlConsolidationFunction` | r/o | Returns the function code used for the current consolidation. |
| `consolidation options` | `boolean` | r/o | Returns a three-element list of boolean values. If the element is true, that option is set. Element 1 is use labels in top row, element 2 is use labels in left column, and element 3 is create links to source data. |
| `consolidation sources` | `text` | r/o | Returns an list of string values that name the source sheets for the worksheet's current consolidation. |
| `display page breaks` | `boolean` | r/w | Returns or sets if page breaks, both automatic and manual, on the specified worksheet are displayed. |
| `enable autofilter` | `boolean` | r/w | Returns or sets if autofilter arrows are enabled when user-interface-only protection is turned on. |
| `enable calculation` | `boolean` | r/w | Returns or sets if Microsoft Excel automatically recalculates the worksheet when necessary. If false, the user cannot request a recalculation, Microsoft Excel never recalculates the sheet automatically. |
| `enable outlining` | `boolean` | r/w | Returns or sets if outlining symbols are enabled when user-interface-only protection is turned on. |
| `enable pivot table` | `boolean` | r/w | Returns or sets if pivot table controls and actions are enabled when user-interface-only protection is turned on. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `filter mode` | `boolean` | r/o | Returns true if the worksheet is in filter mode. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `next` | `sheet` | r/o | Returns a worksheet object that represents the next sheet. |
| `outline object` | `outline` | r/o | Returns an outline object that represents the outline for the specified worksheet. |
| `page setup object` | `page setup` | r/o | Returns the page setup object associated with this chart. |
| `previous` | `sheet` | r/o | Returns a worksheet object that represents the previous sheet. |
| `protect contents` | `boolean` | r/o | Returns true if the contents of the sheet are protected. |
| `protect drawing objects` | `boolean` | r/o | Returns true if shapes are protected. |
| `protection mode` | `boolean` | r/o | Returns true if user-interface-only protection is turned on. To turn on user interface protection, use the protect method with the user interface only argument set to true. |
| `protection object` | `Protection` | r/o | Returns a Protection object that represents the protection options of the worksheet. |
| `scroll area` | `text` | r/w | Returns or sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected. |
| `sheet tab` | `tab` | r/o | Returns the sheet tab of the work sheet |
| `sort object` | `sort` | r/o | Returns the sort object in the sheet. |
| `standard height` | `real` | r/o | Returns the standard default height of all the rows in the worksheet, in points. |
| `standard width` | `real` | r/w | Returns or sets the standard default width of all the columns in the worksheet. |
| `transition expression evaluation` | `boolean` | r/w | Returns or sets if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet. |
| `used range` | `range` | r/o | Returns a range object that represents the used range on the specified worksheet. |
| `visible` | `XlSheetVisibility` | r/w | Returns or sets if the worksheet is visible. |
| `worksheet type` | `XlSheetType` | r/o | Returns the type of this worksheet. |

### `slicer`

A slicer is a mechanism for filtering data in PivotTable views and cube functions.

*Plural:* `slicers`

### `sort`

Represents a sort of a range of data.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `match case` | `boolean` | r/w | Set to true to perform a case-sensitive sort or set to false to perform non-case sensitive sort. Read/Write. |
| `sort header` | `XlYesNoGuess` | r/w | Specifies whether the first row contains header information. Read/Write. |
| `sort method` | `XlSortMethod` | r/w | Specifies the sort method for Chinese/Japanese languages. Read/Write. |
| `sort orientation` | `XlSortOrientation` | r/w | Specifies the orientation for the sort. Read/Write. |
| `sortfieldset` | `null` | r/o | Stores the sort state for workbooks, lists, and autofilters. Read-Only. |
| `sortrange` | `range` | r/o | Returns a range object that represents the range to which the specified sort applies. Read Only. |

### `sortfield`

The sortfield object contains all the sort information for the worksheet, lists, and autofilter object.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `custom order` | `null` | r/w | Specifies a custom order to sort the fields. Read/Write. |
| `sort data option` | `XlSortDataOption` | r/w | Specifies how to sort text in the range specified in sortfield object. Read/Write. |
| `sort key` | `range` | r/o | Specifies the range that is currently being sorted on. Read only. |
| `sort on` | `XlSortOn` | r/w | Returns or sets what attribute of the cell to sort on. Read/Write. |
| `sort on values` | `null` | r/o | Return the value on which the sort is performed for the specified sortfield object. Read Only. |
| `sort order` | `XlSortOrder` | r/w | Determines the sort order for the values specified in the key. Read/Write. |
| `sort priority` | `integer` | r/w | Specifies the priority for the sortfield. Read/Write. |

### `sortfieldset`

The sortfieldset collection is a collection of sortfield objects. It allows developers to store a sort state on worksheets, lists, and autofilters.

### `spinner`

*Plural:* `spinners`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `display threeD shading` | `boolean` | r/w | Returns or sets whether a 3-d effect will be employed when displaying the control |
| `enabled` | `boolean` | r/w | Returns or sets if the control is enabled |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or sets the height of the control. |
| `left position` | `real` | r/w | Returns or sets the left position of the control |
| `linked cell` | `text` | r/w | Returns or sets reference to a call which contains the value of the control. |
| `locked` | `boolean` | r/w | Returns or sets whether the control is locked for editing. |
| `maximum value` | `integer` | r/w | Returns or sets the maximum value allowed |
| `minimum value` | `integer` | r/w | Returns or sets the minimum value allowed |
| `name` | `text` | r/w | Returns or sets the name of this control |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `small change` | `integer` | r/w | Returns or sets the change in value of one click on the spinner control. |
| `top` | `real` | r/w | Returns or sets the position of the top of the control. |
| `top left cell` | `range` | r/o | Returns the top left cell of the range within which the control is positioned. |
| `value` | `integer` | r/w | Returns or sets the current value of the control |
| `visible` | `boolean` | r/w | Returns or sets if the control is visible |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the most recently set z order position, bring shape to front/send shape to back/bring shape forward/send shape backward/bring shape in front of text/send shape behind text. |

### `tab`

Represents the sheet tab of a work sheet or chart sheet.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `integer` | r/w |  |
| `color index` | `XlColorIndex` | r/w |  |
| `theme color` | `XlThemeColor` | r/w |  |
| `tint and shade` | `real` | r/w |  |

### `table style element`

Represents a table style element

*Plural:* `table style elements`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `has format` | `boolean` | r/o | Returns or sets if table style element has format. |
| `interior object` | `interior` | r/o | Returns a interior object that represents the interior for the specified table style element. |

### `table style`

Represents a table style

*Plural:* `table styles`

**Contains:** `table style element`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `built in` | `boolean` | r/o | if this is a built in table style or not |
| `default` | `text` | r/o |  |
| `name` | `text` | r/o | Returns or sets the name of the object. |
| `namelocal` | `text` | r/o | Returns or sets the local name of the object. |
| `show as available pivot table style` | `boolean` | r/w | whether to show as avaliable pivot table style |
| `show as available table style` | `boolean` | r/w | whether to show as avaliable table style |

### `textbox`

*Plural:* `textboxes`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `add indent` | `boolean` | r/w | Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. |
| `auto scale font` | `boolean` | r/w | Returns or sets  if the text in the object changes font size when the object size changes. |
| `auto size` | `boolean` | r/w | Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `text` | r/w | Returns or sets the caption for this object. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `formula` | `text` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `height` | `real` | r/w | Returns or set the height of the object. |
| `horizontal alignment` | `XlHAlign` | r/w | Returns or sets the horizontal alignment for the object. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the position of the specified object, in points. |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `locked text` | `boolean` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `orientation` | `XlOrientation` | r/w | May also be a number value from -90 to 90 degrees. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `rounded corners` | `boolean` | r/w |  |
| `shadow` | `boolean` | r/w |  |
| `string value` | `text` | r/w | Returns or sets the text of the specified object. |
| `top` | `real` | r/w | Returns the top position of the specified object, in points. |
| `top left cell` | `range` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `vertical alignment` | `XlVerticalAlignmentTarget` | r/w | Returns or sets the vertical alignment of the object. |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `top 10 format condition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `applies to` | `range` | r/o | Returns the range this format condition applies to. Read Only. |
| `calc for` | `XlCalcFor` | r/w |  |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `format condition priority` | `integer` | r/w | Specifies the priority for the format condition. Read/Write. |
| `format condition type` | `XlFormatConditionType` | r/o | Return the conditional format type. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `number format` | `text` | r/w | Returns or sets the format code for the object. |
| `pivot condition scope type` | `XlPivotConditionScope` | r/w | Returns or sets the part of the pivot table that the pivot table format condition is scoped to |
| `pivot table condition` | `boolean` | r/o | Tells whether this format condition is a pivot table condition. |
| `stop if true` | `boolean` | r/o | Tells whether the processing of format conditions stops if this condition is true. Read Only. |
| `top 10 percentage` | `boolean` | r/w |  |
| `top 10 rank` | `integer` | r/w |  |
| `top or bottom` | `XlTopBottom` | r/w |  |

### `treeview control`

Represents the hierarchical member-selection control of a cube field.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `drilled` | `null` | r/w |  |
| `hidden` | `null` | r/w |  |

### `unique values format condition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `applies to` | `range` | r/o | Returns the range this format condition applies to. Read Only. |
| `duplicate or unique` | `XlDupeUnique` | r/w |  |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `format condition priority` | `integer` | r/w | Specifies the priority for the format condition. Read/Write. |
| `format condition type` | `XlFormatConditionType` | r/o | Return the conditional format type. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `number format` | `text` | r/w | Returns or sets the format code for the object. |
| `pivot condition scope type` | `XlPivotConditionScope` | r/w | Returns or sets the part of the pivot table that the pivot table format condition is scoped to |
| `pivot table condition` | `boolean` | r/o | Tells whether this format condition is a pivot table condition. |
| `stop if true` | `boolean` | r/o | Tells whether the processing of format conditions stops if this condition is true. Read Only. |

### `validation`

Represents data validation for a worksheet range.

*Plural:* `validations`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `IME mode` | `XlIMEMode` | r/w | Returns or sets the description of the Japanese input rules. |
| `alert style` | `XlDVAlertStyle` | r/o | Returns the validation alert style. |
| `error message` | `text` | r/w | Returns or sets the data validation error message. |
| `error title` | `text` | r/w | Returns or sets the title of the data-validation error dialog box. |
| `formula1` | `text` | r/o | Returns the value or expression associated with the conditional format or data validation. |
| `formula2` | `text` | r/o | Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format Operator property is operator between or operator not between. |
| `ignore blank` | `boolean` | r/w | Returns or sets if blank values are permitted by the range data validation. |
| `in cell dropdown` | `boolean` | r/w | Returns or sets if data validation displays a drop-down list that contains acceptable values. |
| `input message` | `text` | r/w | Returns or sets the data validation input message. |
| `input title` | `text` | r/w | Returns or sets the title of the data-validation input dialog box. |
| `show error` | `boolean` | r/w | Returns or sets if the data validation error message will be displayed whenever the user enters invalid data. |
| `show input` | `boolean` | r/w | Returns or sets if the data validation input message will be displayed whenever the user selects a cell in the data validation range. |
| `validation operator` | `XlFormatConditionOperator` | r/o | Returns the operator for the conditional format or data validation. |
| `validation type` | `XlDVType` | r/o | Returns the data validation type. |
| `value` | `boolean` | r/o | Returns true if all the validation criteria are met, that is, if the range contains valid data. |

### `value change`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allocation method` | `XlAllocationMethod` | r/o |  |
| `allocation value` | `XlAllocationValue` | r/o |  |
| `allocation weight expression` | `text` | r/o |  |
| `order` | `integer` | r/o |  |
| `pivot cell` | `pivot cell` | r/o |  |
| `tuple` | `text` | r/o |  |
| `value` | `real` | r/o |  |
| `visible in pivot table` | `boolean` | r/o |  |

### `vertical page break`

Represents a vertical page break.

*Plural:* `vertical page breaks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `extent` | `XlPageBreakExtent` | r/o | Returns the type of the specified page break: full-screen or only within a print area. |
| `location` | `range` | r/w | Returns or sets where this vertical page break is located. |
| `vertical page break type` | `XlPageBreak` | r/w | Returns or sets the type of vertical page break. |

### `web options`

Contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow png` | `boolean` | r/w | Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save documents as a Web page. |
| `encoding` | `MsoEncoding` | r/w | Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document. |
| `location of components` | `text` | r/w | Returns or sets the central URL, on the intranet or Web, or path, local or network, to the location from which authorized users can download Microsoft Office Web components when viewing your saved document. |
| `pixels per inch` | `integer` | r/w | Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120. |
| `screen size` | `MsoScreenSize` | r/o | Returns the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser. |

### `window`

Represents a window. Many worksheet characteristics, such as scroll bars and gridlines, are actually properties of the window.

*Plural:* `windows`

**Contains:** `pane`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active cell` | `cell` | r/o | Returns a range object that represents the active cell in the active window, the window on top,  or in the specified window. If the window isn't displaying a worksheet, this property fails. |
| `active pane` | `pane` | r/o | Returns a pane object that represents the active pane in the active window, the window on top,  or in the specified window. If the window isn't displaying a worksheet, this property fails. |
| `active sheet` | `sheet` | r/o | Returns an object that represents the active sheet, the sheet on top, in the active workbook or in the specified window or workbook. |
| `caption` | `text` | r/w | Returns or sets the name that appears in the title bar of the document window. |
| `date grouping` | `boolean` | r/w | Returns or sets the date grouping flag in the window. |
| `display formulas` | `boolean` | r/w | Returns or set if the window is displaying formulas.  If false, the window is displaying values. |
| `display gridlines` | `boolean` | r/w | Returns or sets if gridlines are displayed. |
| `display headings` | `boolean` | r/w | Returns or sets if both row and column headings are displayed. |
| `display horizontal scroll bar` | `boolean` | r/w | Returns or sets if the horizontal scroll bar is displayed. |
| `display outline` | `boolean` | r/w | Returns or sets if outline symbols are displayed. |
| `display vertical scroll bar` | `boolean` | r/w | Returns or sets if the vertical scroll bar is displayed. |
| `display workbook tabs` | `boolean` | r/w | Returns or sets if the workbook tabs are displayed. |
| `display zeros` | `boolean` | r/w | Returns or sets if zero values are displayed. |
| `enable resize` | `boolean` | r/w | Returns or sets if the window can be resized. |
| `entry_index` | `integer` | r/o | Returns the index of this item in the element list of windows. |
| `freeze panes` | `boolean` | r/w | Returns or sets if split panes are frozen. It's possible for freeze panes to be true and split to be false, or vice versa. This property applies only to worksheets and macro sheets. |
| `gridline color` | `integer` | r/w | Returns or sets the gridline color as an RGB value. |
| `gridline color index` | `XlColorIndex` | r/w | Returns or sets the gridline color as an index into the current color palette. |
| `height` | `real` | r/w | Returns or sets the height of the window. Use the usable height property to determine the maximum size for the window. You cannot set this property if the window is maximized or minimized. Use the window state property to determine the window state. |
| `left position` | `real` | r/w | Returns or sets the distance from the left edge of the client area to the left edge of the window. |
| `range selection` | `range` | r/o | Returns a range object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet. |
| `scroll column` | `integer` | r/w | Returns or sets the number of the leftmost column in the window. |
| `scroll row` | `integer` | r/w | Returns or sets the number of the row that appears at the top of the window. |
| `selected sheets` | `text` | r/o | Returns a list of sheet objects that represents all the selected sheets in the specified window. |
| `selection` | `range` | r/o | Returns the selected range object in the specified window. |
| `split` | `boolean` | r/w | Returns or sets if the window is split. |
| `split column` | `integer` | r/w | Returns or sets the column number where the window is split into panes, the number of columns to the left of the split line. |
| `split horizontal` | `real` | r/w | Returns or sets the location of the horizontal window split, in points. |
| `split row` | `integer` | r/w | Returns or sets the row number where the window is split into panes, the number of rows above the split. |
| `split vertical` | `real` | r/w | Returns or sets the location of the vertical window split, in points. |
| `tab ratio` | `real` | r/w | Returns or sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar, as a number between 0 and 1, the default value is 0.75. |
| `top` | `real` | r/w | The distance from the top edge of the window to the top edge of the usable area, below the menus, any toolbars docked at the top, and the formula bar. You cannot set this property for a maximized window. |
| `usable height` | `real` | r/o | Returns the maximum height of the space that a window can occupy in points. |
| `usable width` | `real` | r/o | Returns the maximum width of the space that a window can occupy in the application window area, in points. |
| `view` | `XlWindowView` | r/w | Returns or sets the view showing in the window. |
| `visible` | `boolean` | r/w | Returns or sets if the window is visible. |
| `visible range` | `range` | r/o | Returns a range object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. |
| `width` | `real` | r/w | Returns or sets the width of the window. |
| `window number` | `integer` | r/o | Returns the window number. For example, a window named Book1.xls:2 has 2 as its window number. Most windows have the window number 1. |
| `window state` | `XlWindowState` | r/w | Returns or sets the state of the window. |
| `window type` | `XlWindowType` | r/o | Returns the window type. |
| `zoom` | `XLZoomType` | r/w | Returns or sets the display size of the window, as a percentage,100 equals normal size, 200 equals double size, and so on. |

### `workbook connection`

A connection is a set of information needed to obtain data from an external data source other than an 1st_Excel12 workbook.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `_default` | `text` | r/o |  |
| `description` | `text` | r/o | Returns or sets a brief description for a WorkbookConnection object. |
| `name` | `text` | r/o | Returns or sets the name of the workbook connection object. |
| `ranges` | `null` | r/o | Returns the range of object for the specified WorkbookConnection object. |
| `type` | `XlConnectionType` | r/o |  |

### `workbook`

Represents a Microsoft Excel workbook.

*Plural:* `workbooks`

**Contains:** `document property`, `chart sheet`, `command bar`, `custom document property`, `named item`, `pivot cache`, `sheet`, `style`, `custom view`, `window`, `worksheet`, `international macro sheet`, `macro sheet`, `table style`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accept labels in formulas` | `boolean` | r/w | Returns or sets if labels can be used in worksheet formulas. |
| `accuracy version` | `integer` | r/w | Returns or sets the accuracy version for this workbook. |
| `active sheet` | `sheet` | r/o | Returns an object that represents the active sheet, the sheet on top, in the specified window. |
| `auto update frequency` | `integer` | r/w | Returns or sets the number of minutes between automatic updates to the shared workbook. If this property is set to zero  updates occur only when the workbook is saved. |
| `auto update save changes` | `boolean` | r/w | Returns or sets if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. if false changes aren't posted, this workbook is still synchronized with changes made by other users. |
| `calculation version` | `integer` | r/o | Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits, on the left, are the major version of Microsoft Excel. |
| `change history duration` | `integer` | r/w | Returns or sets the number of days shown in the shared workbook's change history. |
| `conflict resolution` | `XlSaveConflictResolution` | r/w | Returns or sets the way conflicts are to be resolved whenever a shared workbook is updated. |
| `create backup` | `boolean` | r/o | Returns true if a backup file is created when this file is saved. |
| `date 1904` | `boolean` | r/w | Returns or sets if the workbook uses the 1904 date system. |
| `default pivottable style` | `null` | r/w | Set or get the default pivot table style for the current workbook |
| `default table style` | `null` | r/w | Set or get the default table style for the current workbook |
| `display drawing objects` | `XlDisplayDrawingObjects` | r/w | Returns or sets how shapes are displayed. |
| `enable auto recover` | `boolean` | r/w | Gets or sets a value that indicates whether Microsoft Office Excel saves changed files, of all formats, on a timed interval. |
| `excel 8 compatibility mode` | `boolean` | r/o | Gets a value that indicates whether the workbook is in compatibility mode. |
| `file format` | `XlFileFormat` | r/o | Returns the file format of the workbook. |
| `full name` | `text` | r/o | Returns the name of the workbook, including its path on disk, as a string. |
| `full name urlencoded` | `text` | r/o | Returns a String indicating the name of the object, including its path on disk, as a string. |
| `has password` | `boolean` | r/o | Returns true if the workbook has a protection password. |
| `has vb project` | `boolean` | r/o | Gets a value that indicates whether a workbook has an attached Microsoft Visual Basic for Applications <VBA> project. |
| `highlight changes on screen` | `boolean` | r/w | Returns or sets if changes to the shared workbook are highlighted on-screen. |
| `inactive list border visible` | `boolean` | r/w | Gets or sets a value that indicates whether list borders are visible when a list is not active. |
| `is add in` | `boolean` | r/w | Returns or sets if the workbook is running as an add in. |
| `keep change history` | `boolean` | r/w | Returns or sets if change tracking is enabled for the shared workbook. |
| `list changes on new sheet` | `boolean` | r/w | Returns or sets if changes to the shared workbook are shown on a separate worksheet. |
| `multi user editing` | `boolean` | r/o | Returns true if the workbook is open as a shared list. |
| `name` | `text` | r/o | Returns or sets the name of the workbook. |
| `password` | `text` | r/w | Returns or sets the password that must be supplied to open the specified workbook. |
| `path` | `text` | r/o | Returns the complete path of the object, excluding the final separator and name of the object. |
| `personal view list settings` | `boolean` | r/w | Returns or sets if filter and sort settings for lists are included in the user's personal view of the shared workbook. |
| `personal view print settings` | `boolean` | r/w | Returns or sets if print settings are included in the user's personal view of the shared workbook. |
| `precision as displayed` | `boolean` | r/w | Returns or sets if calculations in this workbook will be done using only the precision of the numbers as they're displayed. |
| `protect structure` | `boolean` | r/o | Returns true if the order of the sheets in the workbook is protected. |
| `protect windows` | `boolean` | r/o | Returns true if the windows of the workbook are protected. |
| `read only` | `boolean` | r/o | Returns true if the workbook has been opened as read-only. |
| `read only recommended` | `boolean` | r/w | Returns or sets if the workbook was saved as read-only recommended. |
| `remove personal information` | `boolean` | r/w | Returns or sets if personal information can be removed from the specified workbook. |
| `revision number` | `integer` | r/o | Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns zero. |
| `save link values` | `boolean` | r/w | Returns or sets if Microsoft Excel saves external link values with the workbook. |
| `saved` | `boolean` | r/w | Returns or sets if no changes have been made to the specified workbook since it was last saved. |
| `show conflict history` | `boolean` | r/w | Returns or sets if the conflict history worksheet is visible in the workbook that's open as a shared list. |
| `template remove external data` | `boolean` | r/w | Returns or sets if external data references are removed when the workbook is saved as a template. |
| `theme` | `office theme` | r/o | Represents a Microsoft Office theme. |
| `update remote references` | `boolean` | r/w | Returns or sets if Microsoft Excel updates remote references in for the workbook. |
| `user status` | `any` | r/o | Returns a list of lists that provides information about each user who has the workbook open as a shared list. The list is name of user, date and time user opened the workbook, and a number 1 for exclusive, 2 for shared access. |
| `web options` | `web options` | r/o | Returns the web options object, which contains workbook-level attributes used by Microsoft Excel when you save a document as a Web page or open a Web page. |
| `workbook comments` | `text` | r/w | Returns or sets the comment string for this workbook. |
| `write password` | `text` | r/w | Returns or sets a string for the write password of a workbook. |
| `write reserved` | `boolean` | r/o | Return true if the workbook is write-reserved |
| `write reserved by` | `text` | r/o | Returns the name of the user who currently has write permission for the workbook. |

### `worksheet`

*Plural:* `worksheets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `enable selection` | `XlEnableSelection` | r/w | Returns or sets what can be selected on the sheet. |
| `protect scenarios` | `boolean` | r/o | Returns true if the worksheet scenarios are protected. |

### `xlspelling options`

Contains global application-level spelling options

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `dictionary lang` | `XlLanguage` | r/w | Returns or sets the LCID used by the proofing tools |

### Enumerations

### `XlChartGallery`

| Value | Description |
|-------|-------------|
| `built in chart` |  |
| `user defined` |  |
| `any gallery` |  |

### `XlColorIndex`

| Value | Description |
|-------|-------------|
| `color index automatic` |  |
| `color index none` |  |
| `a color index integer` |  |

### `XlEndStyleCap`

| Value | Description |
|-------|-------------|
| `cap` |  |
| `no cap` |  |

### `XlRowCol`

| Value | Description |
|-------|-------------|
| `by columns` |  |
| `by rows` |  |

### `XlScaleType`

| Value | Description |
|-------|-------------|
| `scale linear` |  |
| `scale logarithmic` |  |

### `XlDataSeriesType`

| Value | Description |
|-------|-------------|
| `autofill series` |  |
| `chronological series` |  |
| `growth series` |  |
| `data series linear` |  |

### `XlAxisCrosses`

| Value | Description |
|-------|-------------|
| `axis crosses automatic` |  |
| `axis crosses custom` |  |
| `axis crosses maximum` |  |
| `axis crosses minimum` |  |

### `XlAxisGroup`

| Value | Description |
|-------|-------------|
| `primary axis` |  |
| `secondary axis` |  |

### `XlBackground`

| Value | Description |
|-------|-------------|
| `background automatic` |  |
| `background opaque` |  |
| `background transparent` |  |

### `XlWindowState`

| Value | Description |
|-------|-------------|
| `window state maximized` |  |
| `window state minimized` |  |
| `window state normal` |  |

### `XlAxisType`

| Value | Description |
|-------|-------------|
| `category axis` |  |
| `series axis` |  |
| `value axis` |  |

### `XlArrowHeadLength`

| Value | Description |
|-------|-------------|
| `arrowhead length long` |  |
| `arrowhead length medium` |  |
| `arrowhead length short` |  |

### `XlVAlign`

| Value | Description |
|-------|-------------|
| `valign bottom` |  |
| `valign center` |  |
| `valign distributed` |  |
| `valign justify` |  |
| `valign top` |  |

### `XlTickMark`

| Value | Description |
|-------|-------------|
| `tick mark cross` |  |
| `tick mark inside` |  |
| `tick mark none` |  |
| `tick mark outside` |  |

### `XlErrorBarDirection`

| Value | Description |
|-------|-------------|
| `error bar direction x` |  |
| `error bar direction y` |  |

### `XlErrorBarInclude`

| Value | Description |
|-------|-------------|
| `error bar include both` |  |
| `error bar include minus values` |  |
| `error bar include none` |  |
| `error bar include plus values` |  |

### `XlDisplayBlanksAs`

| Value | Description |
|-------|-------------|
| `interpolated` |  |
| `not plotted` |  |
| `zero` |  |

### `XlArrowHeadStyle`

| Value | Description |
|-------|-------------|
| `arrowhead style closed` |  |
| `arrowhead style double closed` |  |
| `arrowhead style double open` |  |
| `arrowhead style none` |  |
| `arrowhead style open` |  |

### `XlArrowHeadWidth`

| Value | Description |
|-------|-------------|
| `arrowhead width medium` |  |
| `arrowhead width narrow` |  |
| `arrowhead width wide` |  |

### `XlHAlign`

| Value | Description |
|-------|-------------|
| `horizontal align center` |  |
| `horizontal align center across selection` |  |
| `horizontal align distributed` |  |
| `horizontal align fill` |  |
| `horizontal align general` |  |
| `horizontal align justify` |  |
| `horizontal align left` |  |
| `horizontal align right` |  |

### `XlTickLabelPosition`

| Value | Description |
|-------|-------------|
| `tick label position high` |  |
| `tick label position low` |  |
| `tick label position next to axis` |  |
| `tick label position none` |  |

### `XlLegendPosition`

| Value | Description |
|-------|-------------|
| `legend position bottom` |  |
| `legend position corner` |  |
| `legend position left` |  |
| `legend position right` |  |
| `legend position top` |  |

### `XlChartPictureType`

| Value | Description |
|-------|-------------|
| `chart picture type stack scale` |  |
| `chart picture type stack` |  |
| `chart picture type stretch` |  |

### `XlChartPicturePlacement`

| Value | Description |
|-------|-------------|
| `sides` |  |
| `end` |  |
| `end sides` |  |
| `front` |  |
| `front sides` |  |
| `front end` |  |
| `all faces` |  |

### `XlOrientation`

| Value | Description |
|-------|-------------|
| `orientation downward` |  |
| `orientation horizontal` |  |
| `orientation upward` |  |
| `orientation vertical` |  |

### `XlTickLabelOrientation`

| Value | Description |
|-------|-------------|
| `tick label orientation automatic` |  |
| `tick label orientation downward` |  |
| `tick label orientation horizontal` |  |
| `tick label orientation upward` |  |
| `tick label orientation vertical` |  |

### `XlBorderWeight`

| Value | Description |
|-------|-------------|
| `border weight hairline` |  |
| `border weight medium` |  |
| `border weight thick` |  |
| `border weight thin` |  |

### `XlDataSeriesDate`

| Value | Description |
|-------|-------------|
| `series date day` |  |
| `series date month` |  |
| `series date weekday` |  |
| `series date year` |  |

### `XlUnderlineStyle`

| Value | Description |
|-------|-------------|
| `underline style double` |  |
| `underline style double accounting` |  |
| `underline style none` |  |
| `underline style single` |  |
| `underline style single accounting` |  |

### `XlErrorBarType`

| Value | Description |
|-------|-------------|
| `error bar type custom` |  |
| `error bar type fixed value` |  |
| `error bar type percent` |  |
| `error bar type standard deviation` |  |
| `error bar type standard error` |  |

### `XlTrendlineType`

| Value | Description |
|-------|-------------|
| `exponential` |  |
| `linear` |  |
| `logarithmic` |  |
| `moving average` |  |
| `polynomial` |  |
| `power` |  |

### `XlLineStyle`

| Value | Description |
|-------|-------------|
| `continuous` |  |
| `dash` |  |
| `dash dot` |  |
| `dash dot dot` |  |
| `dot` |  |
| `double` |  |
| `slant dash dot` |  |
| `line style none` |  |

### `XlDataLabelsType`

| Value | Description |
|-------|-------------|
| `data labels show none` |  |
| `data labels show value` |  |
| `data labels show percent` |  |
| `data labels show label` |  |
| `data labels show label and percent` |  |
| `data labels show bubble sizes` |  |

### `XlMarkerStyle`

| Value | Description |
|-------|-------------|
| `marker style automatic` |  |
| `marker style circle` |  |
| `marker style dash` |  |
| `marker style diamond` |  |
| `marker style dot` |  |
| `marker style none` |  |
| `marker style picture` |  |
| `marker style plus` |  |
| `marker style square` |  |
| `marker style star` |  |
| `marker style triangle` |  |
| `marker style x` |  |

### `XlPattern`

| Value | Description |
|-------|-------------|
| `pattern automatic` |  |
| `pattern checker` |  |
| `pattern criss cross` |  |
| `pattern down` |  |
| `pattern gray 16` |  |
| `pattern gray 25` |  |
| `pattern gray 50` |  |
| `pattern gray 75` |  |
| `pattern gray 8` |  |
| `pattern grid` |  |
| `pattern horizontal` |  |
| `pattern light down` |  |
| `pattern light horizontal` |  |
| `pattern light up` |  |
| `pattern light vertical` |  |
| `pattern none` |  |
| `pattern semi gray 75` |  |
| `pattern solid` |  |
| `pattern up` |  |
| `pattern vertical` |  |
| `pattern linear gradient` |  |
| `pattern rectangular gradient` |  |

### `XlChartSplitType`

| Value | Description |
|-------|-------------|
| `split by position` |  |
| `split by percent value` |  |
| `split by custom split` |  |
| `split by value` |  |

### `XlDisplayUnit`

| Value | Description |
|-------|-------------|
| `hundreds` |  |
| `thousands` |  |
| `ten thousands` |  |
| `hundred thousands` |  |
| `millions` |  |
| `ten millions` |  |
| `hundred millions` |  |
| `thousand millions` |  |
| `million millions` |  |
| `custom display unit` |  |

### `XlDataLabelPosition`

| Value | Description |
|-------|-------------|
| `label position center` |  |
| `label position above` |  |
| `label position below` |  |
| `label position left` |  |
| `label position right` |  |
| `label position outside end` |  |
| `label position inside end` |  |
| `label position inside base` |  |
| `label position best fit` |  |
| `label position mixed` |  |
| `label position custom` |  |

### `XlTimeUnit`

| Value | Description |
|-------|-------------|
| `days` |  |
| `months` |  |
| `years` |  |

### `XlCategoryType`

| Value | Description |
|-------|-------------|
| `category scale` |  |
| `time scale` |  |
| `automatic scale` |  |

### `XlBarShape`

| Value | Description |
|-------|-------------|
| `box` |  |
| `pyramid to point` |  |
| `pyramid to max` |  |
| `cylinder` |  |
| `cone to point` |  |
| `cone to max` |  |

### `XlChartType`

| Value | Description |
|-------|-------------|
| `column clustered` |  |
| `column stacked` |  |
| `column stacked 100` |  |
| `ThreeD column clustered` |  |
| `ThreeD column stacked` |  |
| `ThreeD column stacked 100` |  |
| `bar clustered` |  |
| `bar stacked` |  |
| `bar stacked 100` |  |
| `ThreeD bar clustered` |  |
| `ThreeD bar stacked` |  |
| `ThreeD bar stacked 100` |  |
| `line stacked` |  |
| `line stacked 100` |  |
| `line markers` |  |
| `line markers stacked` |  |
| `line markers stacked 100` |  |
| `pie of pie` |  |
| `pie exploded` |  |
| `ThreeD pie exploded` |  |
| `bar of pie` |  |
| `xy scatter smooth` |  |
| `xy scatter smooth no markers` |  |
| `xy scatter lines` |  |
| `xy scatter lines no markers` |  |
| `area stacked` |  |
| `area stacked 100` |  |
| `ThreeD area stacked` |  |
| `ThreeD area stacked 100` |  |
| `doughnut exploded` |  |
| `radar markers` |  |
| `radar filled` |  |
| `surface` |  |
| `surface wireframe` |  |
| `surface top view` |  |
| `surface top view wireframe` |  |
| `bubble` |  |
| `bubble ThreeD effect` |  |
| `stock HLC` |  |
| `stock OHLC` |  |
| `stock VHLC` |  |
| `stock VOHLC` |  |
| `cylinder column clustered` |  |
| `cylinder column stacked` |  |
| `cylinder column stacked 100` |  |
| `cylinder bar clustered` |  |
| `cylinder bar stacked` |  |
| `cylinder bar stacked 100` |  |
| `cylinder column` |  |
| `cone column clustered` |  |
| `cone column stacked` |  |
| `cone column stacked 100` |  |
| `cone bar clustered` |  |
| `cone bar stacked` |  |
| `cone bar stacked 100` |  |
| `cone col` |  |
| `pyramid column clustered` |  |
| `pyramid column stacked` |  |
| `pyramid column stacked 100` |  |
| `pyramid bar clustered` |  |
| `pyramid bar stacked` |  |
| `pyramid bar stacked 100` |  |
| `pyramid column` |  |
| `ThreeD column` |  |
| `line chart` |  |
| `ThreeD line` |  |
| `ThreeD pie` |  |
| `pie chart` |  |
| `xyscatter` |  |
| `ThreeD area` |  |
| `area chart` |  |
| `doughnut` |  |
| `radar` |  |
| `combination chart` |  |
| `treemap` |  |
| `histogram` |  |
| `waterfall` |  |
| `sunburst` |  |
| `boxwhisker` |  |
| `pareto` |  |
| `funnel` |  |
| `map chart` |  |

### `XlChartItem`

| Value | Description |
|-------|-------------|
| `data label` |  |
| `a chart area` |  |
| `a series` |  |
| `a chart title` |  |
| `walls` |  |
| `a corners object` |  |
| `data table` |  |
| `trendline` |  |
| `error bars object` |  |
| `xerror bars` |  |
| `yerror bars` |  |
| `legend entry` |  |
| `legend key` |  |
| `shape` |  |
| `major gridlines` |  |
| `minor gridlines` |  |
| `axis title` |  |
| `up bars` |  |
| `plot area` |  |
| `down bars` |  |
| `axis` |  |
| `series lines` |  |
| `floor` |  |
| `legend` |  |
| `hi lo lines` |  |
| `drop lines` |  |
| `radar axis labels` |  |
| `nothing` |  |
| `leader lines` |  |
| `display unit label` |  |

### `XlSizeRepresents`

| Value | Description |
|-------|-------------|
| `size is width` |  |
| `size is area` |  |

### `XlInsertShiftDirection`

| Value | Description |
|-------|-------------|
| `shift down` |  |
| `shift to right` |  |

### `XlDeleteShiftDirection`

| Value | Description |
|-------|-------------|
| `shift to left` |  |
| `shift up` |  |

### `XlDirection`

| Value | Description |
|-------|-------------|
| `toward the bottom` |  |
| `toward the left` |  |
| `toward the right` |  |
| `toward the top` |  |

### `XlConsolidationFunction`

| Value | Description |
|-------|-------------|
| `do average` |  |
| `do count` |  |
| `do count numbers` |  |
| `do maximum` |  |
| `do minimum` |  |
| `do product` |  |
| `do standard deviation` |  |
| `do standard deviation p` |  |
| `do sum` |  |
| `do var` |  |
| `do var p` |  |

### `XlSheetType`

| Value | Description |
|-------|-------------|
| `sheet type chart` |  |
| `sheet type dialog sheet` |  |
| `sheet type excel 4 intl macro sheet` |  |
| `sheet type excel 4 macro sheet` |  |
| `sheet type worksheet` |  |

### `XlLocationInTable`

| Value | Description |
|-------|-------------|
| `column header` |  |
| `column item` |  |
| `data header` |  |
| `data item` |  |
| `page header` |  |
| `page item` |  |
| `row header` |  |
| `row item` |  |
| `table body` |  |

### `XlFindLookIn`

| Value | Description |
|-------|-------------|
| `formulas` |  |
| `comments` |  |
| `values` |  |

### `XlWindowType`

| Value | Description |
|-------|-------------|
| `window type chart as window` |  |
| `window type chart in place` |  |
| `window type clipboard` |  |
| `window type info` |  |
| `window type workbook` |  |

### `XlPivotFieldDataType`

| Value | Description |
|-------|-------------|
| `pivot field type date` |  |
| `pivot field type number` |  |
| `pivot field type text` |  |

### `XlCopyPictureFormat`

| Value | Description |
|-------|-------------|
| `bitmap` |  |
| `picture` |  |

### `XlPivotTableSourceType`

| Value | Description |
|-------|-------------|
| `consolidation` |  |
| `database` |  |
| `external` |  |
| `pivot table` |  |

### `XlReferenceStyle`

| Value | Description |
|-------|-------------|
| `A1` |  |
| `R1C1` |  |

### `XlMSApplication`

| Value | Description |
|-------|-------------|
| `Microsoft Access` |  |
| `Microsoft Fox Pro` |  |
| `Microsoft Mail` |  |
| `Microsoft PowerPoint` |  |
| `Microsoft Project` |  |
| `Microsoft Schedule Plus` |  |
| `Microsoft Word` |  |

### `XlMouseButton`

| Value | Description |
|-------|-------------|
| `no button` |  |
| `primary button` |  |
| `secondary button` |  |

### `XlCutCopyMode`

| Value | Description |
|-------|-------------|
| `copy mode` |  |
| `cut mode` |  |

### `XlFilterAction`

| Value | Description |
|-------|-------------|
| `filter copy` |  |
| `filter in place` |  |

### `XlFormulaVersion`

| Value | Description |
|-------|-------------|
| `using formula` |  |
| `using formula2` |  |

### `XlOrder`

| Value | Description |
|-------|-------------|
| `down then over` |  |
| `over then down` |  |

### `XlLinkType`

| Value | Description |
|-------|-------------|
| `link type Excel links` |  |
| `link type OLE links` |  |

### `XlApplyNamesOrder`

| Value | Description |
|-------|-------------|
| `column then row` |  |
| `row then column` |  |

### `XlEnableCancelKey`

| Value | Description |
|-------|-------------|
| `cancel key disabled` |  |
| `error handler` |  |
| `interrupt` |  |

### `XlPageBreak`

| Value | Description |
|-------|-------------|
| `page break automatic` |  |
| `page break manual` |  |
| `page break none` |  |

### `XlPageOrientation`

| Value | Description |
|-------|-------------|
| `landscape` |  |
| `portrait` |  |

### `XlLinkInfo`

| Value | Description |
|-------|-------------|
| `edition date` |  |
| `update state` |  |

### `XlCommandUnderlines`

| Value | Description |
|-------|-------------|
| `command underlines automatic` |  |
| `command underlines off` |  |
| `command underlines on` |  |

### `XlOLEVerb`

| Value | Description |
|-------|-------------|
| `verb open` |  |
| `verb primary` |  |

### `XlCalculation`

| Value | Description |
|-------|-------------|
| `calculation automatic` |  |
| `calculation manual` |  |
| `calculation semiautomatic` |  |

### `XlFileAccess`

| Value | Description |
|-------|-------------|
| `workbook read only` |  |
| `workbook read write` |  |

### `XlObjectSize`

| Value | Description |
|-------|-------------|
| `fit to page` |  |
| `full page` |  |
| `full screen` |  |

### `XlLookAt`

| Value | Description |
|-------|-------------|
| `part` |  |
| `whole` |  |

### `XlMailSystem`

| Value | Description |
|-------|-------------|
| `MAPI` |  |
| `no mail system` |  |
| `power talk` |  |

### `XlLinkInfoType`

| Value | Description |
|-------|-------------|
| `link info olelinks` |  |
| `link info publishers` |  |
| `link info subscribers` |  |

### `XlCellType`

| Value | Description |
|-------|-------------|
| `cell type blanks` |  |
| `cell type constants` |  |
| `cell type formulas` |  |
| `cell type last cell` |  |
| `cell type comments` |  |
| `cell type visible` |  |
| `cell type all format conditions` |  |
| `cell type same format conditions` |  |
| `cell type all validation` |  |
| `cell type same validation` |  |

### `XlArrangeStyle`

| Value | Description |
|-------|-------------|
| `arrange style cascade` |  |
| `arrange style horizontal` |  |
| `arrange style tiled` |  |
| `arrange style vertical` |  |

### `XlMousePointer`

| Value | Description |
|-------|-------------|
| `I beam cursor` |  |
| `default cursor` |  |
| `northwest arrow cursor` |  |
| `wait cursor` |  |

### `XlAutoFillType`

| Value | Description |
|-------|-------------|
| `fill default` |  |
| `fill copy` |  |
| `fill series` |  |
| `fill formats` |  |
| `fill values` |  |
| `fill days` |  |
| `fill weekdays` |  |
| `fill months` |  |
| `fill years` |  |
| `linear trend` |  |
| `growth trend` |  |
| `flashfill` |  |

### `XlAutoFilterOperator`

| Value | Description |
|-------|-------------|
| `autofilter and` |  |
| `bottom 10 items` |  |
| `bottom 10 percent` |  |
| `autofilter or` |  |
| `top 10 items` |  |
| `top 10 percent` |  |
| `filter by value` |  |
| `filter by cell color` |  |
| `filter by font color` |  |
| `filter by icon` |  |
| `filter dynamic` |  |
| `filter no fill` |  |
| `filter by automatic font color` |  |
| `filter by no icon` |  |

### `XlClipboardFormat`

| Value | Description |
|-------|-------------|
| `clipboard format biff` |  |
| `clipboard format biff 2` |  |
| `clipboard format biff 3` |  |
| `clipboard format biff 4` |  |
| `clipboard format binary` |  |
| `clipboard format bitmap` |  |
| `clipboard format cgm` |  |
| `clipboard format csv` |  |
| `clipboard format dif` |  |
| `clipboard format dsp text` |  |
| `clipboard format embedded object` |  |
| `clipboard format embed source` |  |
| `clipboard format link` |  |
| `clipboard format link source` |  |
| `clipboard format link source desc` |  |
| `clipboard format movie` |  |
| `clipboard format native` |  |
| `clipboard format object desc` |  |
| `clipboard format object link` |  |
| `clipboard format owner link` |  |
| `clipboard format pict` |  |
| `clipboard format print pict` |  |
| `clipboard format rtf` |  |
| `clipboard format screen pict` |  |
| `clipboard format standard font` |  |
| `clipboard format standard scale` |  |
| `clipboard format sylk` |  |
| `clipboard format table` |  |
| `clipboard format text` |  |
| `clipboard format tool face` |  |
| `clipboard format tool face pict` |  |
| `clipboard format valu` |  |
| `clipboard format wk 1` |  |
| `clipboard format unicode text` |  |
| `clipboard format style text` |  |
| `clipboard format unicode style text` |  |
| `clipboard format biff 5` |  |
| `clipboard format picture build` |  |
| `clipboard format odbc conn` |  |
| `clipboard format odbc sql` |  |
| `clipboard format 3d picture` |  |
| `clipboard format unexpected 38` |  |
| `clipboard format drawing drag drop` |  |
| `clipboard format drawing` |  |
| `clipboard format unexpected 41` |  |
| `clipboard format unexpected 42` |  |
| `clipboard format unexpected 43` |  |
| `clipboard format hyperlink` |  |
| `clipboard format unexpected 45` |  |
| `clipboard format windows bitmap` |  |
| `clipboard format uniform resource locator` |  |
| `clipboard format file name` |  |
| `clipboard format unexpected 50` |  |
| `clipboard format unexpected 51` |  |
| `clipboard format hypertext markup language` |  |
| `clipboard format office scrapbook info` |  |
| `clipboard format portable document format` |  |
| `clipboard format excel internal shape` |  |
| `clipboard format office art shape` |  |

### `XlFileFormat`

| Value | Description |
|-------|-------------|
| `CSV file format` |  |
| `CSV Mac file format` |  |
| `CSV MSDos file format` |  |
| `CSV Windows file format` |  |
| `DBF3 file format` |  |
| `DBF4 file format` |  |
| `DIF file format` |  |
| `Excel2 file format` |  |
| `Excel 2 east asian file format` |  |
| `Excel3 file format` |  |
| `Excel4 file format` |  |
| `Excel5 file format` |  |
| `Excel7 file format` |  |
| `Excel 4 workbook file format` |  |
| `international add in file format` |  |
| `international macro file format` |  |
| `workbook normal file format` |  |
| `SYLK file format` |  |
| `current platform text file format` |  |
| `text Mac file format` |  |
| `text MSDos file format` |  |
| `text printer file format` |  |
| `text windows file format` |  |
| `HTML file format` |  |
| `XML spreadsheet file format` |  |
| `PDF file format` |  |
| `Excel binary file format` |  |
| `Excel XML file format` |  |
| `macro enabled XML file format` |  |
| `macro enabled template file format` |  |
| `template file format` |  |
| `add in file format` |  |
| `Excel98to2004 file format` |  |
| `Excel98to2004 template file format` |  |
| `Excel98to2004 add in file format` |  |

### `XlApplicationInternational`

| Value | Description |
|-------|-------------|
| `twenty_four_hour clock` |  |
| `four digit years` |  |
| `alternate array separator` |  |
| `column separator` |  |
| `country_code` |  |
| `country_setting` |  |
| `currency_before` |  |
| `currency_code` |  |
| `currency_digits` |  |
| `currency_leading_zeros` |  |
| `currency_minus_sign` |  |
| `currency_negative` |  |
| `currency_space_before` |  |
| `currency_trailing_zeros` |  |
| `date_order` |  |
| `date_separator` |  |
| `day code` |  |
| `day leading zero` |  |
| `decimal separator` |  |
| `general format name` |  |
| `hour code` |  |
| `left brace` |  |
| `left bracket` |  |
| `list separator` |  |
| `lower case column letter` |  |
| `lower case row letter` |  |
| `mdy` |  |
| `metric` |  |
| `minute_code` |  |
| `month_code` |  |
| `month_leading_zero` |  |
| `month_name_chars` |  |
| `noncurrency_digits` |  |
| `non english functions` |  |
| `right brace` |  |
| `right bracket` |  |
| `row separator` |  |
| `second code` |  |
| `thousands separator` |  |
| `time leading zero` |  |
| `time separator` |  |
| `upper case column letter` |  |
| `upper case row letter` |  |
| `weekday_name_chars` |  |
| `year code` |  |

### `XlPageBreakExtent`

| Value | Description |
|-------|-------------|
| `page break full` |  |
| `page break partial` |  |

### `XlCellInsertionMode`

| Value | Description |
|-------|-------------|
| `overwrite cells` |  |
| `insert delete cells` |  |
| `insert entire rows` |  |

### `XlFormulaLabel`

| Value | Description |
|-------|-------------|
| `no labels` |  |
| `row labels` |  |
| `column labels` |  |
| `mixed labels` |  |

### `XlHighlightChangesTime`

| Value | Description |
|-------|-------------|
| `since my last save` |  |
| `all changes` |  |
| `not yet reviewed` |  |

### `XlCommentDisplayMode`

| Value | Description |
|-------|-------------|
| `no indicator` |  |
| `comment indicator only` |  |
| `comment and indicator` |  |

### `XlFormatConditionType`

| Value | Description |
|-------|-------------|
| `cell value` |  |
| `expression` |  |
| `color scale` |  |
| `databar` |  |
| `top 10` |  |
| `icon sets` |  |
| `unique values` |  |
| `text string` |  |
| `blanks condition` |  |
| `time period` |  |
| `above average condition` |  |
| `no blanks condition` |  |
| `errors condition` |  |
| `no errors condition` |  |

### `XlFormatConditionOperator`

| Value | Description |
|-------|-------------|
| `operator between` |  |
| `operator not between` |  |
| `operator equal` |  |
| `operator not equal` |  |
| `operator greater` |  |
| `operator less` |  |
| `operator greater equal` |  |
| `operator less equal` |  |

### `XlEnableSelection`

| Value | Description |
|-------|-------------|
| `no restrictions` |  |
| `unlocked cells` |  |
| `no selection` |  |

### `XlDVType`

| Value | Description |
|-------|-------------|
| `validate input only` |  |
| `validate whole number` |  |
| `validate decimal` |  |
| `validate list` |  |
| `validated date` |  |
| `validate time` |  |
| `validate text length` |  |
| `validate custom` |  |

### `XlIMEMode`

| Value | Description |
|-------|-------------|
| `IME mode no control` |  |
| `IME mode on` |  |
| `IME mode off` |  |
| `IME mode disable` |  |
| `IME mode hiragana` |  |
| `IME mode katakana` |  |
| `IME mode katakana half` |  |
| `IME mode alpha full` |  |
| `IME mode alpha` |  |
| `IME mode hangul full` |  |
| `IME mode hangul` |  |

### `XlDVAlertStyle`

| Value | Description |
|-------|-------------|
| `valid alert none` |  |
| `valid alert stop` |  |
| `valid alert warning` |  |
| `valid alert information` |  |

### `XlChartLocation`

| Value | Description |
|-------|-------------|
| `location as new sheet` |  |
| `location as object` |  |
| `location automatic` |  |

### `XlChartElementPosition`

| Value | Description |
|-------|-------------|
| `automatic` |  |
| `custom` |  |

### `XlPivotTableVersionList`

| Value | Description |
|-------|-------------|
| `pivot table version 2000` |  |
| `pivot table version 10` |  |
| `pivot table version 11` |  |
| `pivot table version 12` |  |
| `pivot table version 14` |  |
| `pivot table version current` |  |

### `XlLayoutRowType`

| Value | Description |
|-------|-------------|
| `compact row` |  |
| `tabular row` |  |
| `outline row` |  |

### `XlSubtototalLocationType`

| Value | Description |
|-------|-------------|
| `at top` |  |
| `at bottom` |  |

### `XlAllocation`

| Value | Description |
|-------|-------------|
| `manual allocation` |  |
| `automatic allocation` |  |

### `XlAllocationValue`

| Value | Description |
|-------|-------------|
| `allocate value` |  |
| `allocate increment` |  |

### `XlAllocationMethod`

| Value | Description |
|-------|-------------|
| `equal allocation` |  |
| `weight allocation` |  |

### `XlPivotFieldRepeatLabels`

| Value | Description |
|-------|-------------|
| `do not repeat labels` |  |
| `repeat labels` |  |

### `XlPivotTableMissingItems`

| Value | Description |
|-------|-------------|
| `missing items default` |  |
| `missing items none` |  |
| `missing items max` |  |
| `missing items max2` |  |

### `XlPivotCellType`

| Value | Description |
|-------|-------------|
| `pivot cell value` |  |
| `pivot cell pivot item` |  |
| `pivot cell subtotal` |  |
| `pivot cell grand total` |  |
| `pivot cell data field` |  |
| `pivot cell pivot field` |  |
| `pivot cell page field item` |  |
| `pivot cell custom subtotal` |  |
| `pivot cell data pivot field` |  |
| `pivot cell blank cell` |  |

### `XlCellChangedState`

| Value | Description |
|-------|-------------|
| `cell not changed` |  |
| `cell changed` |  |
| `cell change applied` |  |

### `XlLayoutFormType`

| Value | Description |
|-------|-------------|
| `tabular` |  |
| `outline` |  |

### `XlPivotFilterType`

| Value | Description |
|-------|-------------|
| `pivot top count` |  |
| `pivot bottom count` |  |
| `pivot top percent` |  |
| `pivot bottom percent` |  |
| `pivot top sum` |  |
| `pivot bottom sum` |  |
| `pivot value equals` |  |
| `pivot value is not equal` |  |
| `pivot value is greater than` |  |
| `pivot value is greater than or equal to` |  |
| `pivot value is less than` |  |
| `pivot value is less than or equal to` |  |
| `pivot value is between` |  |
| `pivot value is not between` |  |
| `pivot caption equals` |  |
| `pivot caption does not equal` |  |
| `pivot caption begins with` |  |
| `pivot caption does not begin with` |  |
| `pivot caption ends with` |  |
| `pivot caption does not end with` |  |
| `pivot caption contains` |  |
| `pivot caption does not contain` |  |
| `pivot caption is greater than` |  |
| `pivot caption is greater than or equal to` |  |
| `pivot caption is less than` |  |
| `pivot caption is less than or equal to` |  |
| `pivot caption is between` |  |
| `pivot caption is now between` |  |
| `pivot specific date` |  |
| `pivot not specific date` |  |
| `pivot before` |  |
| `pivot before or equal to` |  |
| `pivot after` |  |
| `pivot after or equal to` |  |
| `pivot between` |  |
| `pivot not between` |  |
| `pivot tomorrow` |  |
| `pivot today` |  |
| `pivot yesterday` |  |
| `pivot next week` |  |
| `pivot this week` |  |
| `pivot last week` |  |
| `pivot next month` |  |
| `pivot this month` |  |
| `pivot last month` |  |
| `pivot next quarter` |  |
| `pivot this quarter` |  |
| `pivot last quarter` |  |
| `pivot next year` |  |
| `pivot this year` |  |
| `pivot last year` |  |
| `pivot year to date` |  |
| `pivot all dates in period quarter1` |  |
| `pivot all dates in period quarter2` |  |
| `pivot all dates in period quarter3` |  |
| `pivot all dates in period quarter4` |  |
| `pivot all dates in period January` |  |
| `pivot all dates in period Feberary` |  |
| `pivot all dates in period March` |  |
| `pivot all dates in period April` |  |
| `pivot all dates in period May` |  |
| `pivot all dates in period June` |  |
| `pivot all dates in period July` |  |
| `pivot all dates in period August` |  |
| `pivot all dates in period September` |  |
| `pivot all dates in period October` |  |
| `pivot all dates in period November` |  |
| `pivot all dates in period December` |  |

### `XlPivotLineType`

| Value | Description |
|-------|-------------|
| `pivot line regular` |  |
| `pivot line subtotal` |  |
| `pivot line grandtotal` |  |
| `pivot line blank` |  |

### `XlCubeFieldType`

| Value | Description |
|-------|-------------|
| `hierarchy` |  |
| `measure` |  |
| `set` |  |

### `XlCubeFieldSubType`

| Value | Description |
|-------|-------------|
| `cube hierarchy` |  |
| `cube measure` |  |
| `cube set` |  |
| `cube attribute` |  |
| `cube calculated measure` |  |
| `cube KPI value` |  |
| `cube KPI goal` |  |
| `cube KPI status` |  |
| `cube KPI trend` |  |
| `cube KPI weight` |  |

### `XlPropertyDisplayedIn`

| Value | Description |
|-------|-------------|
| `display property in pivot table` |  |
| `display property in tooltip` |  |
| `display property in pivot table and tooltip` |  |

### `XlCalculatedMemberType`

| Value | Description |
|-------|-------------|
| `calculated member` |  |
| `calculated set` |  |

### `XlConnectionType`

| Value | Description |
|-------|-------------|
| `connection type OLEDB` |  |
| `connection type ODBC` |  |
| `connection type XMLMAP` |  |
| `connection type TEXT` |  |
| `connection type WEB` |  |

### `XlPasteSpecialOperation`

| Value | Description |
|-------|-------------|
| `paste special operation add` |  |
| `paste special operation divide` |  |
| `paste special operation multiply` |  |
| `paste special operation none` |  |
| `paste special operation subtract` |  |

### `XlPasteType`

| Value | Description |
|-------|-------------|
| `paste all` |  |
| `paste all using source theme` |  |
| `paste all except borders` |  |
| `paste formats` |  |
| `paste formulas` |  |
| `paste comments` |  |
| `paste values` |  |
| `paste column widths` |  |
| `paste validation` |  |
| `paste formulas and number formats` |  |
| `paste values and number formats` |  |

### `XlPhoneticCharacterType`

| Value | Description |
|-------|-------------|
| `phonetic character half width katakana` |  |
| `phonetic character full width katakana` |  |
| `phonetic character hiragana` |  |
| `no phonetic character conversion` |  |

### `XlPhoneticAlignment`

| Value | Description |
|-------|-------------|
| `phonetic align no control` |  |
| `phonetic align left` |  |
| `phonetic align center` |  |
| `phonetic align distributed` |  |

### `XlPictureAppearance`

| Value | Description |
|-------|-------------|
| `printer` |  |
| `screen` |  |

### `XlPivotFieldOrientation`

| Value | Description |
|-------|-------------|
| `orient as column field` |  |
| `orient as data field` |  |
| `orient as hidden` |  |
| `orient as page field` |  |
| `orient as row field` |  |

### `XlPivotFieldCalculation`

| Value | Description |
|-------|-------------|
| `pivot field calculation difference from` |  |
| `pivot field calculation index` |  |
| `pivot field calculation no additional calculation` |  |
| `pivot field calculation percent difference from` |  |
| `pivot field calculation percent of` |  |
| `pivot field calculation percent of column` |  |
| `pivot field calculation percent of row` |  |
| `pivot field calculation percent of total` |  |
| `pivot field calculation running total` |  |

### `XlPlacement`

| Value | Description |
|-------|-------------|
| `placement free floating` |  |
| `placement move` |  |
| `placement move and size` |  |

### `XlPlatform`

| Value | Description |
|-------|-------------|
| `Macintosh` |  |
| `MSDos` |  |
| `MSWindows` |  |

### `XlPrintLocation`

| Value | Description |
|-------|-------------|
| `print sheet end` |  |
| `print in place` |  |
| `print no comments` |  |

### `XlPriority`

| Value | Description |
|-------|-------------|
| `priority high` |  |
| `priority low` |  |
| `priority normal` |  |

### `XlPTSelectionMode`

| Value | Description |
|-------|-------------|
| `selection mode label only` |  |
| `selection mode data and label` |  |
| `selection mode data only` |  |
| `selection mode origin` |  |
| `selection mode button` |  |
| `selection mode blanks` |  |

### `XlRangeAutoFormat`

| Value | Description |
|-------|-------------|
| `range autoformat threeD effects 1` |  |
| `range autoformat threeD effects 2` |  |
| `range autoformat accounting 1` |  |
| `range autoformat accounting 2` |  |
| `range autoformat accounting 3` |  |
| `range autoformat accounting 4` |  |
| `range autoformat classic 1` |  |
| `range autoformat classic 2` |  |
| `range autoformat classic 3` |  |
| `range autoformat color 1` |  |
| `range autoformat color 2` |  |
| `range autoformat color 3` |  |
| `range autoformat list 1` |  |
| `range autoformat list 2` |  |
| `range autoformat list 3` |  |
| `range autoformat local format 1` |  |
| `range autoformat local format 2` |  |
| `range autoformat local format 3` |  |
| `range autoformat local format 4` |  |
| `range autoformat none` |  |
| `range autoformat simple` |  |

### `XlRoutingSlipDelivery`

| Value | Description |
|-------|-------------|
| `all at once` |  |
| `one after another` |  |

### `XlRoutingSlipStatus`

| Value | Description |
|-------|-------------|
| `not yet routed` |  |
| `routing complete` |  |
| `routing in progress` |  |

### `XlRunAutoMacro`

| Value | Description |
|-------|-------------|
| `auto activate` |  |
| `auto close` |  |
| `auto deactivate` |  |
| `auto open` |  |

### `XlSaveAsAccessMode`

| Value | Description |
|-------|-------------|
| `exclusive` |  |
| `no change` |  |
| `shared` |  |

### `XlSaveConflictResolution`

| Value | Description |
|-------|-------------|
| `local session changes` |  |
| `other session changes` |  |
| `user resolution` |  |

### `XlSearchDirection`

| Value | Description |
|-------|-------------|
| `search next` |  |
| `search previous` |  |

### `XlSearchOrder`

| Value | Description |
|-------|-------------|
| `by columns` |  |
| `by rows` |  |

### `XlSheetVisibility`

| Value | Description |
|-------|-------------|
| `sheet visible` |  |
| `sheet hidden` |  |
| `sheet very hidden` |  |

### `XlSortMethod`

| Value | Description |
|-------|-------------|
| `pin yin` | Phonetic Chinese/Japanese sort order for characters. This is the default value. |
| `stroke` | Sort by the quantity of strokes in each character. |

### `XlSortOrder`

| Value | Description |
|-------|-------------|
| `sort ascending` | Sorts the specified field in ascending order. This is the default value. |
| `sort descending` | Sorts the specified field in descending order. |
| `sort manual` | It is not supported. |

### `XlSortOrientation`

| Value | Description |
|-------|-------------|
| `sort rows` | Sorts by row. this is the default value. |
| `sort columns` | Sorts by column. |

### `XlSortType`

| Value | Description |
|-------|-------------|
| `sort labels` | Sorts the PivotTable report by labels. |
| `sort values` | Sorts the PivotTable report by values. |

### `XlSpecialCellsValue`

| Value | Description |
|-------|-------------|
| `errors` |  |
| `logical` |  |
| `numbers` |  |
| `text values` |  |

### `XlSummaryRow`

| Value | Description |
|-------|-------------|
| `summary above` |  |
| `summary below` |  |

### `XlSummaryColumn`

| Value | Description |
|-------|-------------|
| `summary on left` |  |
| `summary on right` |  |

### `XlSummaryReportType`

| Value | Description |
|-------|-------------|
| `summary pivot table` |  |
| `standard summary` |  |

### `XlTextParsingType`

| Value | Description |
|-------|-------------|
| `delimited` |  |
| `fixed width` |  |

### `XlTextQualifier`

| Value | Description |
|-------|-------------|
| `text qualifier double quote` |  |
| `text qualifier none` |  |
| `text qualifier single quote` |  |

### `XlWBATemplate`

| Value | Description |
|-------|-------------|
| `chart` |  |
| `Excel 4 intl macro sheet` |  |
| `Excel 4 macro sheet` |  |
| `worksheet` |  |

### `XlWindowView`

| Value | Description |
|-------|-------------|
| `normal view` |  |
| `page layout view` |  |

### `XlXLMMacroType`

| Value | Description |
|-------|-------------|
| `macro type command` |  |
| `macro type function` |  |
| `macro type not XLM` |  |

### `XlYesNoGuess`

| Value | Description |
|-------|-------------|
| `header guess` | Default value. Excel determines whether there is a header, and where it is, if there is one. |
| `header no` | The entire range should be sorted. |
| `header yes` | The entire range should not be sorted. |

### `XlDisplayDrawingObjects`

| Value | Description |
|-------|-------------|
| `display shapes` |  |
| `hide` |  |
| `placeholders` |  |

### `XlBordersIndex`

| Value | Description |
|-------|-------------|
| `inside horizontal` |  |
| `inside vertical` |  |
| `diagonal down` |  |
| `diagonal up` |  |
| `edge bottom` |  |
| `edge left` |  |
| `edge right` |  |
| `edge top` |  |
| `border bottom` |  |
| `border left` |  |
| `border right` |  |
| `border top` |  |

### `XlToolbarProtection`

| Value | Description |
|-------|-------------|
| `no button changes` |  |
| `no changes` |  |
| `no docking changes` |  |
| `toolbar protection none` |  |
| `no shape changes` |  |

### `XlBuiltInDialog`

| Value | Description |
|-------|-------------|
| `dialog open` |  |
| `dialog open links` |  |
| `dialog save as` |  |
| `dialog file delete` |  |
| `dialog page setup` |  |
| `dialog print` |  |
| `dialog printer setup` |  |
| `dialog arrange all` |  |
| `dialog window size` |  |
| `dialog window move` |  |
| `dialog run` |  |
| `dialog set print titles` |  |
| `dialog font` |  |
| `dialog display` |  |
| `dialog protect document` |  |
| `dialog calculation` |  |
| `dialog extract` |  |
| `dialog data delete` |  |
| `dialog sort` |  |
| `dialog data series` |  |
| `dialog table` |  |
| `dialog format number` |  |
| `dialog alignment` |  |
| `dialog style` |  |
| `dialog border` |  |
| `dialog cell protection` |  |
| `dialog column width` |  |
| `dialog clear` |  |
| `dialog paste special` |  |
| `dialog edit delete` |  |
| `dialog insert` |  |
| `dialog paste names` |  |
| `dialog define name` |  |
| `dialog create names` |  |
| `dialog formula goto` |  |
| `dialog formula find` |  |
| `dialog gallery area` |  |
| `dialog gallery bar` |  |
| `dialog gallery column` |  |
| `dialog gallery line` |  |
| `dialog gallery pie` |  |
| `dialog gallery scatter` |  |
| `dialog combination` |  |
| `dialog gridlines` |  |
| `dialog axes` |  |
| `dialog attach text` |  |
| `dialog patterns` |  |
| `dialog main chart` |  |
| `dialog overlay` |  |
| `dialog scale` |  |
| `dialog format legend` |  |
| `dialog format text` |  |
| `dialog parse` |  |
| `dialog unhide` |  |
| `dialog workspace` |  |
| `dialog activate` |  |
| `dialog copy picture` |  |
| `dialog delete name` |  |
| `dialog delete format` |  |
| `dialog new` |  |
| `dialog row height` |  |
| `dialog format move` |  |
| `dialog format size` |  |
| `dialog formula replace` |  |
| `dialog select special` |  |
| `dialog apply names` |  |
| `dialog replace font` |  |
| `dialog split` |  |
| `dialog outline` |  |
| `dialog save workbook` |  |
| `dialog copy chart` |  |
| `dialog format font` |  |
| `dialog note` |  |
| `dialog set update status` |  |
| `dialog color palette` |  |
| `dialog change link` |  |
| `dialog app move` |  |
| `dialog app size` |  |
| `dialog main chart type` |  |
| `dialog overlay chart type` |  |
| `dialog open mail` |  |
| `dialog send mail` |  |
| `dialog standard font` |  |
| `dialog consolidate` |  |
| `dialog sort special` |  |
| `dialog gallery threeD area` |  |
| `dialog gallery threeD column` |  |
| `dialog gallery threeD line` |  |
| `dialog gallery threeD pie` |  |
| `dialog view threeD` |  |
| `dialog goal seek` |  |
| `dialog workgroup` |  |
| `dialog fill group` |  |
| `dialog update link` |  |
| `dialog promote` |  |
| `dialog demote` |  |
| `dialog show detail` |  |
| `dialog object properties` |  |
| `dialog save new object` |  |
| `dialog apply style` |  |
| `dialog assign to object` |  |
| `dialog object protection` |  |
| `dialog show toolbar` |  |
| `dialog print preview` |  |
| `dialog edit color` |  |
| `dialog format main` |  |
| `dialog format overlay` |  |
| `dialog edit series` |  |
| `dialog define style` |  |
| `dialog gallery radar` |  |
| `dialog zoom` |  |
| `dialog insert object` |  |
| `dialog size` |  |
| `dialog move` |  |
| `dialog format auto` |  |
| `dialog gallery threeD bar` |  |
| `dialog gallery threeD surface` |  |
| `dialog customize toolbar` |  |
| `dialog workbook add` |  |
| `dialog workbook move` |  |
| `dialog workbook copy` |  |
| `dialog workbook options` |  |
| `dialog save workspace` |  |
| `dialog chart wizard` |  |
| `dialog assign to tool` |  |
| `dialog placement` |  |
| `dialog fill workgroup` |  |
| `dialog workbook new` |  |
| `dialog scenario cells` |  |
| `dialog scenario add` |  |
| `dialog scenario edit` |  |
| `dialog scenario summary` |  |
| `dialog pivot table wizard` |  |
| `dialog pivot field properties` |  |
| `dialog options calculation` |  |
| `dialog options edit` |  |
| `dialog options view` |  |
| `dialog add in manager` |  |
| `dialog menu editor` |  |
| `dialog attach toolbars` |  |
| `dialog options chart` |  |
| `dialog vba insert file` |  |
| `dialog vba procedure definition` |  |
| `dialog routing slip` |  |
| `dialog mail logon` |  |
| `dialog insert picture` |  |
| `dialog gallery doughnut` |  |
| `dialog chart trend` |  |
| `dialog workbook insert` |  |
| `dialog options transition` |  |
| `dialog options general` |  |
| `dialog filter advanced` |  |
| `dialog mail next letter` |  |
| `dialog data label` |  |
| `dialog insert title` |  |
| `dialog font properties` |  |
| `dialog macro options` |  |
| `dialog workbook unhide` |  |
| `dialog workbook name` |  |
| `dialog gallery custom` |  |
| `dialog add chart autoformat` |  |
| `dialog chart add data` |  |
| `dialog tab order` |  |
| `dialog subtotal create` |  |
| `dialog workbook tab split` |  |
| `dialog workbook protect` |  |
| `dialog scrollbar properties` |  |
| `dialog pivot show pages` |  |
| `dialog text to columns` |  |
| `dialog format charttype` |  |
| `dialog pivot field group` |  |
| `dialog pivot field ungroup` |  |
| `dialog checkbox properties` |  |
| `dialog label properties` |  |
| `dialog listbox properties` |  |
| `dialog editbox properties` |  |
| `dialog open text` |  |
| `dialog pushbutton properties` |  |
| `dialog filter` |  |
| `dialog function wizard` |  |
| `dialog save copy as` |  |
| `dialog options lists add` |  |
| `dialog series axes` |  |
| `dialog series x` |  |
| `dialog series y` |  |
| `dialog errorbar x` |  |
| `dialog errorbar y` |  |
| `dialog format chart` |  |
| `dialog series order` |  |
| `dialog mail edit mailer` |  |
| `dialog standard width` |  |
| `dialog scenario merge` |  |
| `dialog properties` |  |
| `dialog summary info` |  |
| `dialog find file` |  |
| `dialog active cell font` |  |
| `dialog vba make add in` |  |
| `dialog file sharing` |  |
| `dialog autocorrect` |  |
| `dialog custom views` |  |
| `dialog insert name label` |  |
| `dialog series shape` |  |
| `dialog chart options data labels` |  |
| `dialog chart options data table` |  |
| `dialog set background picture` |  |
| `dialog data validation` |  |
| `dialog chart type` |  |
| `dialog chart location` |  |
| `dialog chart source data` |  |
| `dialog series options` |  |
| `dialog pivot table options` |  |
| `dialog pivot solve order` |  |
| `dialog pivot calculated field` |  |
| `dialog pivot calculated item` |  |
| `dialog conditional formatting` |  |
| `dialog insert hyperlink` |  |
| `dialog protect sharing` |  |
| `dialog phonetic` |  |
| `dialog import text file` |  |
| `dialog web options general` |  |
| `dialog web options pictures` |  |
| `dialog web options files` |  |
| `dialog web options fonts` |  |
| `dialog web options encoding` |  |

### `XlParameterType`

| Value | Description |
|-------|-------------|
| `prompt` |  |
| `constant` |  |
| `range` |  |

### `XlParameterDataType`

| Value | Description |
|-------|-------------|
| `param type unknown` |  |
| `param type char` |  |
| `param type numeric` |  |
| `param type decimal` |  |
| `param type number` |  |
| `param type small int` |  |
| `param type float` |  |
| `param type real` |  |
| `param type double` |  |
| `param type var char` |  |
| `param type date` |  |
| `param type time` |  |
| `param type timestamp` |  |
| `param type long var char` |  |
| `param type binary` |  |
| `param type var binary` |  |
| `param type long var binary` |  |
| `param type big int` |  |
| `param type tiny int` |  |
| `param type bit` |  |

### `XlFormControl`

| Value | Description |
|-------|-------------|
| `button control` |  |
| `check box` |  |
| `drop down` |  |
| `edit box` |  |
| `group box` |  |
| `label` |  |
| `list box` |  |
| `option button` |  |
| `scroll bar` |  |
| `spinner` |  |

### `XlColumnDataType`

| Value | Description |
|-------|-------------|
| `general format` |  |
| `text format` |  |
| `MDY format` |  |
| `DMY format` |  |
| `YMD format` |  |
| `MYD format` |  |
| `DYM format` |  |
| `YDM format` |  |
| `skip column` |  |

### `XlQueryType`

| Value | Description |
|-------|-------------|
| `ODBC query` |  |
| `DAO record set` |  |
| `web query` |  |
| `OLE DB query` |  |
| `text import` |  |
| `ADO recordset` |  |
| `FileMaker query` |  |

### `XlCmdType`

| Value | Description |
|-------|-------------|
| `cmd cube` |  |
| `cmd sql` |  |
| `cmd table` |  |
| `cmd default` |  |
| `cmd list` |  |

### `XlListObjectSourceType`

| Value | Description |
|-------|-------------|
| `src external` |  |
| `src range` |  |
| `src xml` |  |
| `src query` |  |
| `src model` |  |

### `XlFMCriteriaOperator`

| Value | Description |
|-------|-------------|
| `criteria equals` |  |
| `criteria less than or equal to` |  |
| `criteria greater than or equal to` |  |
| `criteria less than` |  |
| `criteria greater than` |  |
| `criteria begins with` |  |
| `criteria ends with` |  |
| `criteria contains` |  |

### `xlFMCriteriaConditional`

| Value | Description |
|-------|-------------|
| `no condition` |  |
| `and condition` |  |
| `or condition` |  |

### `XlRangeValueDataType`

| Value | Description |
|-------|-------------|
| `range value default` |  |
| `range value XML spreadsheet` |  |
| `range value MS persist XML` |  |

### `XLSubTotalsIndex`

| Value | Description |
|-------|-------------|
| `subtotal automatic` |  |
| `subtotal sum` |  |
| `subtotal count` |  |
| `subtotal average` |  |
| `subtotal max` |  |
| `subtotal min` |  |
| `subtotal product` |  |
| `subtotal count numbers` |  |
| `subtotal standard deviation` |  |
| `subtotal standard deviation p` |  |
| `subtotal variable` |  |
| `subtotal variable p` |  |

### `XLDataEntryMode`

| Value | Description |
|-------|-------------|
| `data entry on` |  |
| `data entry strict` |  |
| `data entry off` |  |

### `XLStatusBarState`

| Value | Description |
|-------|-------------|
| `status text` |  |
| `a Boolean` |  |

### `XLTransitionMenuKeyAction`

| Value | Description |
|-------|-------------|
| `excel menus` |  |

### `XLDefaultSheetDir`

| Value | Description |
|-------|-------------|
| `left to right` |  |
| `right to left` |  |
| `context` |  |

### `XLCusorMovement`

| Value | Description |
|-------|-------------|
| `normal cursor` |  |
| `logical cursor` |  |
| `visual cursor` |  |

### `XLRangeReference`

| Value | Description |
|-------|-------------|
| `range object` | range object |
| `A1-style range reference` | range R1C1 reference |
| `named range` | range R1C1 reference |

### `XLSubTotalType`

| Value | Description |
|-------|-------------|
| `automatic subtotal` |  |
| `sum subtotal` |  |
| `count subtotal` |  |
| `average subtotal` |  |
| `maximum value` |  |
| `minimum value` |  |
| `product subtotal` |  |
| `count numbers subtotal` |  |
| `standard deviation` |  |
| `standard deviation P` |  |
| `variance subtotal` |  |
| `variance P subtotal` |  |

### `XLAutoShowType`

| Value | Description |
|-------|-------------|
| `type_automatic` |  |
| `type_manual` |  |

### `XLAutoShowPosition`

| Value | Description |
|-------|-------------|
| `position top` |  |
| `position bottom` |  |

### `XLScrollTabPosition`

| Value | Description |
|-------|-------------|
| `scroll tab position first` |  |
| `scroll tab position last` |  |

### `XLPivotTableWizardSourceData`

| Value | Description |
|-------|-------------|
| `range` |  |
| `a list of ranges` |  |
| `report name` |  |
| `a list of string that is a SQL query` |  |

### `XLDefaultChartTemplate`

| Value | Description |
|-------|-------------|
| `built in chart template` |  |
| `format name` |  |

### `XLDefaultChartType`

| Value | Description |
|-------|-------------|
| `built in chart type` |  |
| `custom chart` |  |

### `XLCustomListType`

| Value | Description |
|-------|-------------|
| `range object` | range object |
| `A1-style range reference` | range R1C1 reference |
| `named range` | range R1C1 reference |
| `list of strings` |  |

### `XLInputDefault`

| Value | Description |
|-------|-------------|
| `range object` | range object |
| `A1-style range reference` | range R1C1 reference |
| `named range` | range R1C1 reference |
| `input default as string` |  |

### `XLInputType`

| Value | Description |
|-------|-------------|
| `a number` | range object |
| `input type as string` | range R1C1 reference |
| `a number or a string` | range R1C1 reference |
| `a bool` |  |
| `range object` |  |
| `list of numbers` |  |
| `list of strings` |  |
| `list of number or string` |  |
| `list of bools` |  |
| `list of range objects` |  |

### `XLZoomType`

| Value | Description |
|-------|-------------|
| `a number` |  |
| `a bool` |  |

### `XLSourceDataLocation`

| Value | Description |
|-------|-------------|
| `range object` | range object |
| `A1-style range reference` | range R1C1 reference |
| `named range` | range R1C1 reference |
| `list of strings` | A list of SQL query strings |

### `XLSourceData`

| Value | Description |
|-------|-------------|
| `percentable` | A percentage between 10 and 400 |
| `a bool` |  |

### `XLOnDataType`

| Value | Description |
|-------|-------------|
| `script` | A script object |
| `script Text` |  |

### `XlRangeTarget`

| Value | Description |
|-------|-------------|
| `application` |  |
| `worksheet` |  |
| `A1-style range reference` |  |

### `XlHorizAlignmentTarget`

| Value | Description |
|-------|-------------|
| `horizontal aligment bottom` |  |
| `horizontal aligment left` |  |
| `horizontal aligment right` |  |
| `horizontal aligment top` |  |

### `XlVerticalAlignmentTarget`

| Value | Description |
|-------|-------------|
| `vertical alignment top` |  |
| `vertical alignment center` |  |
| `vertical alignment bottom` |  |
| `vertical alignment justify` |  |
| `vertical alignment distributed` |  |

### `XlCheckBoxState`

| Value | Description |
|-------|-------------|
| `checkbox off` |  |
| `checkbox on` |  |
| `checkbox mixed` |  |

### `XlEditBoxItem`

| Value | Description |
|-------|-------------|
| `text` |  |
| `a number` |  |
| `xl number` |  |
| `reference` |  |
| `formula` |  |

### `XlMultiSelect`

| Value | Description |
|-------|-------------|
| `select none` |  |
| `select simple` |  |
| `select extended` |  |

### `XlReplacements`

| Value | Description |
|-------|-------------|
| `text to replace` |  |
| `replacement text` |  |

### `XlCategoryNames`

| Value | Description |
|-------|-------------|
| `range object` | range object |
| `A1-style range reference` | range R1C1 reference |
| `named range` | range R1C1 reference |
| `list of category names` | A list category names |

### `MyUDateLinks`

| Value | Description |
|-------|-------------|
| `do not update links` |  |
| `update external links only` |  |
| `update remote links only` |  |
| `update remote and external links` |  |

### `MyODelimiter`

| Value | Description |
|-------|-------------|
| `tab delimiter` |  |
| `commas delimiter` |  |
| `spaces delimiter` |  |
| `semicolon delimiter` |  |
| `no delimiter` |  |
| `custom character delimiter` |  |

### `XlColorVariance`

| Value | Description |
|-------|-------------|
| `vary by color` |  |
| `vary by shade` |  |
| `vary by grayscale` |  |
| `vary by same color` |  |

### `XlTickHAlign`

| Value | Description |
|-------|-------------|
| `align tick label center` |  |
| `align tick label left` |  |
| `align tick label right` |  |

### `XlLanguage`

| Value | Description |
|-------|-------------|
| `Basque` |  |
| `Catalan` |  |
| `Chinese` |  |
| `Chinese Taiwan` |  |
| `Czech` |  |
| `Danish` |  |
| `Dutch` |  |
| `English US` |  |
| `English AUS` |  |
| `English British` |  |
| `English CAN` |  |
| `Finnish` |  |
| `French` |  |
| `French Canadian` |  |
| `German` |  |
| `German Austria` |  |
| `Swiss German` |  |
| `Greek` |  |
| `Hungarian` |  |
| `Italian` |  |
| `Japanese` |  |
| `Korean` |  |
| `Malaysian` |  |
| `Norwegian Bokmal` |  |
| `Norwegian` |  |
| `Polish` |  |
| `Portuguese Brazil` |  |
| `Portuguese Iberian` |  |
| `Russian` |  |
| `Slovak` |  |
| `Slovenian` |  |
| `Spanish` |  |
| `Swedish` |  |
| `Turkish` |  |

### `XlSortOn`

| Value | Description |
|-------|-------------|
| `sort on cell value` | Values. |
| `sort on cell color` | Cell color. |
| `sort on font color` | Font color. |
| `sort on icon` | Icon. |

### `XlSortDataOption`

| Value | Description |
|-------|-------------|
| `sort normal` | Default. Sorts numeric and text data separately. |
| `sort text as numbers` | Treat text as numeric data for the sort. |

### `XlTotalsCalculation`

| Value | Description |
|-------|-------------|
| `none totals calc` |  |
| `sum totals calc` |  |
| `average totals calc` |  |
| `count totals calc` |  |
| `count number totals calc` |  |
| `min totals calc` |  |
| `max totals calc` |  |
| `deviation totals calc` |  |
| `var totals calc` |  |
| `custom totals calc` |  |

### `MsoChartElementType`

| Value | Description |
|-------|-------------|
| `no chart title` |  |
| `chart title centered overlay` |  |
| `chart title above chart` |  |
| `no legend` |  |
| `legend right` |  |
| `legend top` |  |
| `legend left` |  |
| `legend bottom` |  |
| `legend right overlay` |  |
| `legend left overlay` |  |
| `no data label` |  |
| `show data label` |  |
| `data label center` |  |
| `data label inside end` |  |
| `data label inside base` |  |
| `data label outside end` |  |
| `data label left` |  |
| `data label right` |  |
| `data label top` |  |
| `data label bottom` |  |
| `data label best fit` |  |
| `no primary category axis title` |  |
| `primary category axis title adjacent to axis` |  |
| `primary category axis title below axis` |  |
| `primary category axis title rotated` |  |
| `primary category axis title vertical` |  |
| `primary category axis title horizontal` |  |
| `no primary value axis title` |  |
| `primary value axis title adjacent to axis` |  |
| `primary value axis title below axis` |  |
| `primary value axis title rotated` |  |
| `primary value axis title vertical` |  |
| `primary value axis title horizontal` |  |
| `no secondary category axis title` |  |
| `secondary category axis title adjacent to axis` |  |
| `secondary category axis title below Axis` |  |
| `secondary category axis title rotated` |  |
| `secondary category axis title vertical` |  |
| `secondary category axis title horizontal` |  |
| `no secondary value axis title` |  |
| `secondary value axis title Adjacent to axis` |  |
| `secondary value axis title below axis` |  |
| `secondary value axis title rotated` |  |
| `secondary value axis title vertical` |  |
| `secondary value axis title horizontal` |  |
| `no series axis title` |  |
| `series axis title rotated` |  |
| `series axis title vertical` |  |
| `series axis title horizontal` |  |
| `no primary value grid lines` |  |
| `primary value grid lines minor` |  |
| `primary value grid lines major` |  |
| `primary value grid lines minor major` |  |
| `no primary category grid lines` |  |
| `primary category grid lines minor` |  |
| `primary category grid lines major` |  |
| `primary category grid lines minor major` |  |
| `no secondary value grid lines` |  |
| `secondary value grid lines minor` |  |
| `secondary value grid lines major` |  |
| `secondary value grid lines minor major` |  |
| `no secondary category grid lines` |  |
| `secondary category grid lines minor` |  |
| `secondary category grid lines major` |  |
| `secondary category grid lines minor major` |  |
| `no series axis grid lines` |  |
| `series axis grid lines minor` |  |
| `series axis grid lines major` |  |
| `series axis grid lines minor major` |  |
| `no primary category axis` |  |
| `primary category axis show` |  |
| `primary category axis without labels` |  |
| `primary category axis reverse` |  |
| `no primary value axis` |  |
| `show primary value axis` |  |
| `primary value axis thousands` |  |
| `primary value axis millions` |  |
| `primary value axis billions` |  |
| `primary value axis log scale` |  |
| `no secondary category axis` |  |
| `show secondary category axis` |  |
| `secondary category axis without labels` |  |
| `secondary category axis reverse` |  |
| `no secondary value axis` |  |
| `show secondary value axis` |  |
| `secondary value axis thousands` |  |
| `secondary value axis millions` |  |
| `secondary value axis billions` |  |
| `secondary value axis log scale` |  |
| `no series axis` |  |
| `show series axis` |  |
| `series axis without labeling` |  |
| `series axis reverse` |  |
| `primary category axis thousands` |  |
| `primary category axis millions` |  |
| `primary category axis billions` |  |
| `primary category axis log scale` |  |
| `secondary category axis thousands` |  |
| `secondary category axis millions` |  |
| `secondary category axis billions` |  |
| `secondary category axis log scale` |  |
| `no data table` |  |
| `show data table` |  |
| `data table with legend keys` |  |
| `no Trendline` |  |
| `trendline add linear` |  |
| `trendline add exponential` |  |
| `trendline add linear forecast` |  |
| `trendline add two period moving average` |  |
| `no error bar` |  |
| `error bar standard error` |  |
| `error bar percentage` |  |
| `error bar standard deviation` |  |
| `no line` |  |
| `line drop line` |  |
| `line hiLo line` |  |
| `line series line` |  |
| `line drop hilo line` |  |
| `no up down bars` |  |
| `show up down bars` |  |
| `no plot area` |  |
| `show plot area` |  |
| `no chart wall` |  |
| `show chart wall` |  |
| `no chart floor` |  |
| `show chart floor` |  |

### `XlDynamicFilterCriteria`

| Value | Description |
|-------|-------------|
| `filter above average` | Filter all above-average values. |
| `filter all dates in april` | Filter all dates in April. |
| `filter all dates in august` | Filter all dates in August. |
| `filter all dates in december` | Filter all dates in December. |
| `filter all dates in february` | Filter all dates in February |
| `filter all dates in january` | Filter all dates in January. |
| `filter all dates in july` | Filter all dates in July. |
| `filter all dates in june` | Filter all dates in June. |
| `filter all dates in march` | Filter all dates in March. |
| `filter all dates in may` | Filter all dates in May. |
| `filter all dates in november` | Filter all dates in November. |
| `filter all dates in october` | Filter all dates in October. |
| `filter all dates in quarter1` | Filter all dates in Quarter1. |
| `filter all dates in quarter2` | Filter all dates in Quarter2. |
| `filter all dates in quarter3` | Filter all dates in Quarter3. |
| `filter all dates in quarter4` | Filter all dates in Quarter4. |
| `filter all dates in september` | Filter all dates in September. |
| `filter below average` | Filter all below-average values. |
| `filter last month` | Filter all values related to last month. |
| `filter last quarter` | Filter all values related to last quarter. |
| `filter last week` | Filter all values related to last week. |
| `filter last year` | Filter all values related to last year. |
| `filter next month` | Filter all values related to next month. |
| `filter next quarter` | Filter all values related to next quarter. |
| `filter next week` | Filter all values related to next week. |
| `filter next year` | Filter all values related to next year. |
| `filter this month` | Filter all values related to the current month. |
| `filter this quarter` | Filter all values related to the current quarter. |
| `filter this week` | Filter all values related to the current week. |
| `filter this year` | Filter all values related to the current year. |
| `filter today` | Filter all values related to the current date. |
| `filter tomorrow` | Filter all values related to tomorrow. |
| `filter year to date` | Filter all values from today until a year ago. |
| `filter yesterday` | Filter all values related to yesterday. |

### `XlThemeFont`

| Value | Description |
|-------|-------------|
| `theme font index none` |  |
| `theme font index major` |  |
| `theme font index minor` |  |

### `XlThemeColor`

| Value | Description |
|-------|-------------|
| `color index none` |  |
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

### `XlCheckInVersionType`

| Value | Description |
|-------|-------------|
| `minor version` |  |
| `major version` |  |
| `overwrite current version` |  |

### `XlWebSelectionType`

| Value | Description |
|-------|-------------|
| `entire page` |  |
| `all tables` |  |
| `specified tables` |  |

### `XlWebFormatting`

| Value | Description |
|-------|-------------|
| `web formatting all` |  |
| `web formatting rtf` |  |
| `web formatting none` |  |

### `XlRobustConnect`

| Value | Description |
|-------|-------------|
| `as required` |  |
| `always` |  |
| `never` |  |

### `XlConditionValueTypes`

| Value | Description |
|-------|-------------|
| `condition value none` |  |
| `condition value number` |  |
| `condition value lowest value` |  |
| `condition value highest value` |  |
| `condition value percent` |  |
| `condition value formula` |  |
| `condition value percentile` |  |
| `condition value automatic minimum` |  |
| `condition value automatic maximum` |  |

### `XlPivotConditionScope`

| Value | Description |
|-------|-------------|
| `pivot condition selection scope` |  |
| `pivot condition fields scope` |  |
| `pivot condition data field scope` |  |

### `XlDataBarFillType`

| Value | Description |
|-------|-------------|
| `databar fill solid` |  |
| `databar fill gradient` |  |

### `XlDataBarBorderType`

| Value | Description |
|-------|-------------|
| `databar border none` |  |
| `databar border solid` |  |

### `XlDataBarAxisPosition`

| Value | Description |
|-------|-------------|
| `databar axis automatic` |  |
| `databar axis midpoint` |  |
| `databar axis none` |  |

### `XlDataBarNegativeFormatType`

| Value | Description |
|-------|-------------|
| `databar automatic` |  |
| `databar positive format` |  |
| `databar custom format` |  |

### `XlIcon`

| Value | Description |
|-------|-------------|
| `format condition icon no cell icon` |  |
| `format condition icon green up arrow` |  |
| `format condition icon yellow side arrow` |  |
| `format condition icon red down arrow` |  |
| `format condition icon gray up arrow` |  |
| `format condition icon gray side arrow` |  |
| `format condition icon gray down arrow` |  |
| `format condition icon green flag` |  |
| `format condition icon yellow flag` |  |
| `format condition icon red flag` |  |
| `format condition icon green circle` |  |
| `format condition icon yellow circle` |  |
| `format condition icon red circle with border` |  |
| `format condition icon black circle with border` |  |
| `format condition icon green traffic light` |  |
| `format condition icon yellow traffic light` |  |
| `format condition icon red traffic light` |  |
| `format condition icon yellow triangle` |  |
| `format condition icon red diamond` |  |
| `format condition icon green check symbol` |  |
| `format condition icon yellow exclamation symbol` |  |
| `format condition icon red cross symbol` |  |
| `format condition icon green check` |  |
| `format condition icon yellow exclamation` |  |
| `format condition icon red cross` |  |
| `format condition icon yellow up incline arrow` |  |
| `format condition icon yellow down incline arrow` |  |
| `format condition icon gray up incline arrow` |  |
| `format condition icon gray down incline arrow` |  |
| `format condition icon red circle` |  |
| `format condition icon pink circle` |  |
| `format condition icon gray circle` |  |
| `format condition icon black circle` |  |
| `format condition icon circle with one white quarter` |  |
| `format condition icon circle with two white quarters` |  |
| `format condition icon circle with three white quarters` |  |
| `format condition icon white circle all white quarters` |  |
| `format condition icon 0 bars` |  |
| `format condition icon 1 bar` |  |
| `format condition icon 2 bars` |  |
| `format condition icon 3 bars` |  |
| `format condition icon 4 bars` |  |
| `format condition icon gold star` |  |
| `format condition icon half gold star` |  |
| `format condition icon silver star` |  |
| `format condition icon green up triangle` |  |
| `format condition icon yellow dash` |  |
| `format condition icon red down triangle` |  |
| `format condition icon 4 filled boxes` |  |
| `format condition icon 3 filled boxes` |  |
| `format condition icon 2 filled boxes` |  |
| `format condition icon 1 filled box` |  |
| `format condition icon 0 filled boxes` |  |

### `XlIconSet`

| Value | Description |
|-------|-------------|
| `icon set custom` |  |
| `icon set 3 arrows` |  |
| `icon set 3 arrows gray` |  |
| `icon set 3 flags` |  |
| `icon set 3 traffic lights 1` |  |
| `icon set 3 traffic lights 2` |  |
| `icon set 3 signs` |  |
| `icon set 3 symbols` |  |
| `icon set 3 symbols 2` |  |
| `icon set 4 arrows` |  |
| `icon set 4 arrows gray` |  |
| `icon set 4 red to black` |  |
| `icon set 4 CRV` |  |
| `icon set 4 traffic lights` |  |
| `icon set 5 arrows` |  |
| `icon set 5 arrows gray` |  |
| `icon set 5 CRV` |  |
| `icon set 5 quarters` |  |
| `icon set 3 stars` |  |
| `icon set 3 triangles` |  |
| `icon set 5 boxes` |  |

### `XlTopBottom`

| Value | Description |
|-------|-------------|
| `top 10 top` |  |
| `top 10 bottom` |  |

### `XlCalcFor`

| Value | Description |
|-------|-------------|
| `calc for all values` |  |
| `calc for row groups` |  |
| `calc for col groups` |  |

### `XlAboveBelow`

| Value | Description |
|-------|-------------|
| `format above average` |  |
| `format below average` |  |
| `format equal above average` |  |
| `format equal below average` |  |
| `format above standard deviation` |  |
| `format below standard deviation` |  |

### `XlDupeUnique`

| Value | Description |
|-------|-------------|
| `format unique values` |  |
| `format duplicate values` |  |

### `XlContainsOperator`

| Value | Description |
|-------|-------------|
| `text contains` |  |
| `text does not contain` |  |
| `text begins with` |  |
| `text ends with` |  |

### `XlTimePeriods`

| Value | Description |
|-------|-------------|
| `date is today` |  |
| `date is yesterday` |  |
| `date is within the last seven days` |  |
| `date is this week` |  |
| `date is last week` |  |
| `date is last month` |  |
| `date is tomorrow` |  |
| `date is next week` |  |
| `date is next month` |  |
| `date is this month` |  |

### `XlDataBarNegativeColorType`

| Value | Description |
|-------|-------------|
| `databar color type color` |  |
| `databar color type same as positive` |  |

### `large scroll`

| Value | Description |
|-------|-------------|
| `window` |  |
| `pane` |  |

### `print out`

| Value | Description |
|-------|-------------|
| `window` |  |
| `sheet` |  |
| `workbook` |  |

### `print preview`

| Value | Description |
|-------|-------------|
| `window` |  |
| `sheet` |  |
| `workbook` |  |

### `small scroll`

| Value | Description |
|-------|-------------|
| `window` |  |
| `pane` |  |

### `calculate`

| Value | Description |
|-------|-------------|
| `application` |  |
| `sheet` |  |

### `check spelling`

| Value | Description |
|-------|-------------|
| `sheet` |  |
| `button` |  |
| `checkbox` |  |
| `option button` |  |
| `groupbox` |  |
| `label` |  |
| `textbox` |  |

### `copy picture`

| Value | Description |
|-------|-------------|
| `button` |  |
| `checkbox` |  |
| `option button` |  |
| `scrollbar` |  |
| `listbox` |  |
| `groupbox` |  |
| `dropdown` |  |
| `spinner` |  |
| `label` |  |
| `textbox` |  |

### `cut`

| Value | Description |
|-------|-------------|
| `button` |  |
| `checkbox` |  |
| `option button` |  |
| `scrollbar` |  |
| `listbox` |  |
| `groupbox` |  |
| `dropdown` |  |
| `spinner` |  |
| `label` |  |
| `textbox` |  |

### `show`

| Value | Description |
|-------|-------------|
| `dialog` |  |
| `scenario` |  |

### `unprotect`

| Value | Description |
|-------|-------------|
| `sheet` |  |
| `workbook` |  |

### `bring to front`

| Value | Description |
|-------|-------------|
| `button` |  |
| `checkbox` |  |
| `option button` |  |
| `scrollbar` |  |
| `listbox` |  |
| `groupbox` |  |
| `dropdown` |  |
| `spinner` |  |
| `label` |  |
| `textbox` |  |

### `send to back`

| Value | Description |
|-------|-------------|
| `button` |  |
| `checkbox` |  |
| `option button` |  |
| `scrollbar` |  |
| `listbox` |  |
| `groupbox` |  |
| `dropdown` |  |
| `spinner` |  |
| `label` |  |
| `textbox` |  |

### `set first priority`

| Value | Description |
|-------|-------------|
| `format condition` |  |
| `color scale format condition` |  |
| `databar format condition` |  |
| `icon set format condition` |  |
| `top 10 format condition` |  |
| `above average format condition` |  |
| `unique values format condition` |  |

### `set last priority`

| Value | Description |
|-------|-------------|
| `format condition` |  |
| `color scale format condition` |  |
| `databar format condition` |  |
| `icon set format condition` |  |
| `top 10 format condition` |  |
| `above average format condition` |  |
| `unique values format condition` |  |

### `modify applies to range`

| Value | Description |
|-------|-------------|
| `format condition` |  |
| `color scale format condition` |  |
| `databar format condition` |  |
| `icon set format condition` |  |
| `top 10 format condition` |  |
| `above average format condition` |  |
| `unique values format condition` |  |

### `clear all filters`

| Value | Description |
|-------|-------------|
| `pivot table` |  |
| `pivot field` |  |

### `delete`

| Value | Description |
|-------|-------------|
| `cube field` |  |
| `calculated member` |  |
| `pivot filter` |  |
| `value change` |  |

### `drill to`

| Value | Description |
|-------|-------------|
| `pivot field` |  |
| `pivot item` |  |

### `clear manual filter`

| Value | Description |
|-------|-------------|
| `cube field` |  |
| `pivot field` |  |

### `reset timer`

| Value | Description |
|-------|-------------|
| `pivot cache` |  |
| `query table` |  |

### `add item to list`

| Value | Description |
|-------|-------------|
| `listbox` |  |
| `dropdown` |  |

### `get border`

| Value | Description |
|-------|-------------|
| `format condition` |  |
| `display format` |  |
| `top 10 format condition` |  |
| `above average format condition` |  |
| `unique values format condition` |  |

### `set list item`

| Value | Description |
|-------|-------------|
| `listbox` |  |
| `dropdown` |  |

### `copy object`

| Value | Description |
|-------|-------------|
| `button` |  |
| `checkbox` |  |
| `option button` |  |
| `scrollbar` |  |
| `listbox` |  |
| `groupbox` |  |
| `dropdown` |  |
| `spinner` |  |
| `label` |  |
| `textbox` |  |

### `get list item`

| Value | Description |
|-------|-------------|
| `listbox` |  |
| `dropdown` |  |

### `remove all items`

| Value | Description |
|-------|-------------|
| `listbox` |  |
| `dropdown` |  |

### `remove item`

| Value | Description |
|-------|-------------|
| `listbox` |  |
| `dropdown` |  |

### `activate object`

| Value | Description |
|-------|-------------|
| `window` |  |
| `sheet` |  |
| `workbook` |  |
| `pane` |  |

---

## Drawing Suite

Classes and Methods used for Graphic Objects

### Commands

### `apply`

Applies to the specified shape formatting that's been copied by using the pick up method.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `automatic length`

Specifies that the first segment of the callout line, the segment attached to the text callout box, be scaled automatically when the callout is moved.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `callout format` | no |  |

### `begin connect`

Attaches the beginning of the specified connector to a specified shape. If there's already a connection between the beginning of the connector and another shape, that connection is broken.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `connector format` | no |  |
| `connected shape` | `shape` | no | The shape to attach the beginning of the connector to. |
| `connection site` | `integer` | no | A connection site on the shape specified by connected shape. Must be an integer between 1 and the integer returned by the connection site count property of the specified shape. |

### `begin disconnect`

Detaches the beginning of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `connector format` | no |  |

### `bring to front`

Bring the object to the front of the z-order of objects.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4029` | no |  |

### `check spelling`

Checks the spelling of an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4034` | no |  |
| `custom dictionary` | `text` | yes | A string that indicates the file name of the custom dictionary to examine if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used. |
| `ignore uppercase` | `boolean` | yes | Set to true to have Microsoft Excel ignore words that are all uppercase. False to have Microsoft Excel check words that are all uppercase. |
| `always suggest` | `boolean` | yes | Set to true to have Microsoft Excel display a list of suggested alternate spellings when an incorrect spelling is found. False to have Microsoft Excel wait for you to input the correct spelling. |

### `copy object`

Copies the object to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4030` | no |  |

### `copy picture`

Copies the selected object to the clipboard as a picture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4031` | no |  |
| `appearance` | `XlPictureAppearance` | yes | Specifies how the picture should be copied. |
| `format` | `XlCopyPictureFormat` | yes | The format of the picture. |

### `custom drop`

Sets the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4035` | no |  |
| `drop` | `real` | no | The drop distance, in points. |

### `custom length`

Specifies that the first segment of the callout line, the segment attached to the text callout box, retain a fixed length whenever the callout is moved.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4036` | no |  |
| `length` | `real` | no | The length of the first segment of the callout, in points. |

### `cut`

Cuts the object to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4032` | no |  |

### `delete gradient stop`

Removes a gradient stop.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
| `stop index` | `integer` | no | The index number of the stop. |

### `end connect`

Attaches the end of the specified connector to a specified shape. If there's already a connection between the end of the connector and another shape, that connection is broken.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `connector format` | no |  |
| `connected shape` | `shape` | no | The shape to attach the end of the connector to. |
| `connection site` | `integer` | no | A connection site on the shape specified by connected shape. Must be an integer between 1 and the integer returned by the connection site count property of the specified shape. |

### `end disconnect`

Detaches the end of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `connector format` | no |  |

### `flip`

Flips the specified shape around its horizontal or vertical axis.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |
| `flip cmd` | `MsoFlipCmd` | no |  |

### `get custom color`

Returns the custom color for the specified Microsoft Office theme.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme color scheme` | no |  |
| `name` | `text` | no |  |

**Returns:** `integer`

### `insert gradient stop`

Adds a stop to a gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
| `custom color` | `integer` | no | Sets the color of the stop within a gradient. |
| `position` | `real` | no | Sets the position of the stop within the gradient expressed as a percent. |
| `transparency` | `real` | yes | Sets the transparency of the stop within the gradient. |
| `stop index` | `integer` | yes | The index number of the stop. |

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
| `(direct)` | `fill format` | no |  |
| `gradient style` | `MsoGradientStyle` | no | The gradient style. |
| `variant` | `integer` | no | The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the gradient tab in the fill effects dialog box. If Style is from center gradient, this argument can be either 1 or 2. |
| `degree` | `real` | no | The gradient degree. Can be a value from 0.0, dark to 1.0, light. |

### `patterned`

Sets the specified fill to a pattern.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
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
| `(direct)` | `4037` | no |  |
| `drop type` | `MsoCalloutDropType` | no | The starting position of the callout line relative to the text bounding box. |

### `preset gradient`

Sets the specified fill to a preset gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
| `gradient style` | `MsoGradientStyle` | no | Specifies the gradient style. |
| `variant` | `integer` | no | The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the gradient tab in the fill effects dialog box. If Style is from center gradient, this argument can be either 1 or 2. |
| `preset gradient type` | `MsoPresetGradientType` | no | The gradient type. |

### `preset textured`

Sets the specified fill to a preset texture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
| `preset texture` | `MsoPresetTexture` | no | The preset texture. |

### `reroute connections`

Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the reroute connections method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `reset rotation`

Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `threeD format` | no |  |

### `save as picture`

Saves the shape in the requested file using the stated graphic format

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |
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

Scales the height of the shape by a specified factor.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `picture` | no |  |
| `factor` | `real` | no | Specifies the ratio between the height of the shape after you resize it and the current or original height. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument. |
| `relative to original size` | `boolean` | no | Set to true to scale the shape relative to its original size. False to scale it relative to its current size. |
| `scale` | `MsoScaleFrom` | yes | Specifies which part of the shape retains its position when the shape is scaled. |

### `scale width`

Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original size or relative to the current size.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `picture` | no |  |
| `factor` | `real` | no | Specifies the ratio between the width of the shape after you resize it and the current or original width. For example, to make a rectangle 50 percent larger, specify 1.5 for this argument. |
| `relative to original size` | `boolean` | no | Set to true to scale the shape relative to its original size. False to scale it relative to its current size. You can specify true for this argument only if the specified shape is a picture. |
| `scale` | `MsoScaleFrom` | yes | Specifies which part of the shape retains its position when the shape is scaled. |

### `send to back`

Sends the object to the back of the z-order.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4033` | no |  |

### `set bullet picture`

Sets the graphics file to be used for bullets in a bulleted list.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `bullet format` | no |  |
| `FileName` | `text` | no |  |

### `set shapes default properties`

Applies the formatting for the specified shape to the default shape. Shapes created after this method has been used will have this formatting applied by default.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `solid`

Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |

### `toggle vertical text`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `word art format` | no |  |

### `two color gradient`

Sets the specified fill to a two-color gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
| `gradient style` | `MsoGradientStyle` | no | The gradient style. |
| `variant` | `integer` | no | The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the gradient tab in the fill effects dialog box. If Style is from center gradient, this argument can be either 1 or 2. |

### `ungroup`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape text frame` | no |  |

### `user picture`

Fills the specified shape with one large image.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
| `picture file` | `text` | no | The name of the picture file. |

### `user textured`

Fills the specified shape with small tiles of an image.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
| `texture file` | `text` | no | The name of the picture file. |

### `z order`

Moves the specified shape in front of or behind other shapes in the collection, that is, changes the shape's position in the z-order.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |
| `z order command` | `MsoZOrderCmd` | no | Specifies where to move the specified shape relative to the other shapes. |

### Classes

### `adjustment`

*Plural:* `adjustments`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `adjustment_value` | `real` | r/w |  |

### `arc`

Represents an arc graphic.

*Plural:* `arcs`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `add indent` | `integer` | r/w | Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. |
| `auto scale font` | `boolean` | r/w | Returns or sets  if the text in the object changes font size when the object size changes. |
| `auto size` | `integer` | r/w | Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bottom right cell` | `integer` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `integer` | r/w | Returns or sets the caption for this object. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `font object` | `integer` | r/o | Returns a font object that represents the font of the specified object. |
| `formula` | `integer` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `height` | `integer` | r/w | Returns or set the height of the object. |
| `horizontal alignment` | `integer` | r/w | Returns or sets the horizontal alignment for the object. |
| `interior object` | `integer` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `integer` | r/w | Returns or sets the position of the specified object, in points. |
| `locked` | `integer` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `locked text` | `integer` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `integer` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `orientation` | `integer` | r/w | May also be a number value from -90 to 90 degrees. |
| `placement` | `integer` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `integer` | r/w | Returns or sets if this object is printed. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `string value` | `integer` | r/w | Returns or sets the text of the specified object. |
| `top` | `integer` | r/w | Returns the top position of the specified object, in points. |
| `top left cell` | `integer` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `vertical alignment` | `integer` | r/w | Returns or sets the vertical alignment of the object. |
| `visible` | `integer` | r/w | Returns or sets if the object is visible. |
| `width` | `integer` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `bullet format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bullet character` | `text` | r/w | Returns or sets the Unicode character value that is used for bullets in the specified text. |
| `bullet font` | `shape font` | r/o | Returns a font object that represents character formatting for a bullet format object. |
| `bullet number` | `integer` | r/o | Returns the bullet number of a paragraph. |
| `bullet start value` | `integer` | r/w | Gets or sets the beginning value of a bulleted list. |
| `bullet style` | `MsoNumberedBulletStyle` | r/w | Returns or sets a constant that represents the style of a bullet. |
| `bullet type` | `MsoBulletType` | r/w | Returns or sets a constant that represents the type of bullet. |
| `relative size` | `real` | r/w | Returns or sets the bullet size relative to the size of the first text character in the paragraph. |
| `use text color` | `boolean` | r/w | Determines whether the specified bullets are set to the color of the first text character in the paragraph. |
| `use text font` | `boolean` | r/w | Determines whether the specified bullets are set to the font of the first text character in the paragraph. |
| `visible` | `boolean` | r/w | Returns or sets a value that specifies whether the bullet is visible. |

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
| `border` | `boolean` | r/w | Returns or sets whether the text in the specified callout is surrounded by a border. |
| `callout format length` | `real` | r/o | When the auto length property of the specified callout is set to false, the length property returns the length in points of the first segment of the callout line, the segment attached to the text callout box. |
| `callout format type` | `MsoCalloutType` | r/w | Returns or sets the callout type. |
| `drop` | `real` | r/o | For callouts with an explicitly set drop value, this property returns the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box. |
| `drop type` | `MsoCalloutDropType` | r/o | Returns a value that indicates where the callout line attaches to the callout text box. |
| `gap` | `real` | r/w | Returns or sets the horizontal distance in points between the end of the callout line and the text bounding box. |

### `callout`

*Plural:* `callouts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `callout format` | `callout format` | r/o | Returns a connector format object that contains connector formatting properties. |
| `callout type` | `MsoCalloutType` | r/o | Returns the type of callout. |

### `connector format`

Contains properties and methods that apply to connectors. A connector is a line that attaches two other shapes at points called connection sites.

*Plural:* `connector formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `begin connected` | `boolean` | r/o | Returns true if the beginning of the specified connector is connected to a shape. |
| `begin connected shape` | `shape` | r/o | Returns a shape object that represents the shape that the beginning of the specified connector is attached to. |
| `begin connection site` | `integer` | r/o | Returns an integer that specifies the connection site that the beginning of a connector is connected to. |
| `connector format type` | `MsoConnectorType` | r/w | Returns or sets the connector type. |
| `end connected` | `boolean` | r/o | Returns true if the end of the specified connector is connected to a shape. |
| `end connected shape` | `shape` | r/o | Returns a shape object that represents the shape that the end of the specified connector is attached to. |
| `end connection site` | `integer` | r/o | Returns an integer that specifies the connection site that the end of a connector is connected to. |

### `fill format`

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.

*Plural:* `fill formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `back color` | `integer` | r/w | Returns or sets a RGB color that represents the background color for the specified fill or patterned line. |
| `back color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the specified fill background color. |
| `fill format type` | `MsoFillType` | r/o | Returns the shape fill format type. |
| `fore color` | `integer` | r/w | Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow. |
| `fore color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the specified foreground fill or solid color. |
| `gradient color type` | `MsoGradientColorType` | r/o | Returns the gradient color type for the specified fill. |
| `gradient degree` | `real` | r/o | Returns a value that indicates how dark or light a one-color gradient fill is. A value of zero means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in. Values between 1 and zero blend. |
| `gradient style` | `MsoGradientStyle` | r/o | Returns the gradient style for the specified fill. |
| `gradient variant` | `integer` | r/o | Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. If the gradient style is from center gradient, this property returns either 1 or 2. |
| `pattern` | `MsoPatternType` | r/o | Returns a value that represents the pattern applied to the specified fill or line. |
| `preset gradient type` | `MsoPresetGradientType` | r/o | Returns the preset gradient type for the specified fill. |
| `preset texture` | `MsoPresetTexture` | r/o | Returns the preset texture for the specified fill. |
| `rotate with object` | `boolean` | r/w | Returns or sets whether the fill rotates with the specified shape. |
| `texture alignment` | `MsoTextureAlignment` | r/w | Returns or sets the texture alignment for the specified object. |
| `texture horizontal scale` | `real` | r/w | Returns or sets the texture alignment for the specified object. |
| `texture name` | `text` | r/o | Returns the name of the custom texture file for the specified fill. |
| `texture offset X` | `real` | r/w | Returns or sets the texture alignment for the specified object. |
| `texture offset Y` | `real` | r/w | Returns or sets the texture alignment for the specified object. |
| `texture tile` | `boolean` | r/w | Returns the texture tile style for the specified fill. |
| `texture type` | `MsoTextureType` | r/o | Returns the texture type for the specified fill. |
| `texture vertical scale` | `real` | r/w | Returns or sets the texture alignment for the specified object. |
| `transparency` | `real` | r/w | Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear. |
| `visible` | `boolean` | r/w | Returns or sets if the specified object, or the formatting applied to it, is visible. |

### `glow format`

Represents the glow formatting for a shape or range of shapes

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `integer` | r/w | Returns or sets the color for the specified glow format. |
| `color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified glow format. |
| `radius` | `real` | r/w | Returns or sets the length of the radius for the specified glow format. |

### `gradient stop`

Represents one gradient stop.

*Plural:* `gradient stops`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `integer` | r/w | Returns or sets the color for the specified the gradient stop. |
| `color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified gradient stop. |
| `position` | `real` | r/w | Returns or sets the position for the specified gradient stop expressed as a percent. |
| `transparency` | `real` | r/w | Returns or sets a value representing the transparency of the gradient fill expressed as a percent. |

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

### `line`

Represents a line graphic object.

*Plural:* `lines`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `arrowhead length` | `XlArrowHeadLength` | r/w | Returns or sets the length of an arrowhead |
| `arrowhead style` | `XlArrowHeadStyle` | r/w | Returns or sets the style of an arrowhead. |
| `arrowhead width` | `XlArrowHeadWidth` | r/w | Returns or sets the width of an arrowhead. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bottom right cell` | `range` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or set the height of the object. |
| `left position` | `real` | r/w | Returns or sets the position of the specified object, in points. |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `top` | `real` | r/w | Returns the top position of the specified object, in points. |
| `top left cell` | `range` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `major theme font`

Represents a container for the font schemes of a Microsoft Office theme.

*Plural:* `major theme fonts`

### `minor theme font`

Represents a container for the font schemes of a Microsoft Office theme.

*Plural:* `minor theme fonts`

### `office theme`

Represents a Microsoft Office theme.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `theme color scheme` | `theme color scheme` | r/o | Returns the color scheme of a Microsoft Office theme. |
| `theme effect scheme` | `theme effect scheme` | r/o | Returns the effects scheme of a Microsoft Office theme. |
| `theme font scheme` | `theme font scheme` | r/o | Returns the font scheme of a Microsoft Office theme. |

### `oval`

Represents an oval graphic.

*Plural:* `ovals`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `add indent` | `integer` | r/w | Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. |
| `auto scale font` | `boolean` | r/w | Returns or sets  if the text in the object changes font size when the object size changes. |
| `auto size` | `integer` | r/w | Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bottom right cell` | `integer` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `integer` | r/w | Returns or sets the caption for this object. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `font object` | `integer` | r/o | Returns a font object that represents the font of the specified object. |
| `formula` | `integer` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `height` | `integer` | r/w | Returns or set the height of the object. |
| `horizontal alignment` | `integer` | r/w | Returns or sets the horizontal alignment for the object. |
| `interior object` | `integer` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `integer` | r/w | Returns or sets the position of the specified object, in points. |
| `locked` | `integer` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `locked text` | `integer` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `integer` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `orientation` | `integer` | r/w | May also be a number value from -90 to 90 degrees. |
| `placement` | `integer` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `integer` | r/w | Returns or sets if this object is printed. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `shadow` | `integer` | r/w | Returns or sets if the object has a shadow. |
| `string value` | `integer` | r/w | Returns or sets the text of the specified object. |
| `top` | `integer` | r/w | Returns the top position of the specified object, in points. |
| `top left cell` | `integer` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `vertical alignment` | `integer` | r/w | Returns or sets the vertical alignment of the object. |
| `visible` | `integer` | r/w | Returns or sets if the object is visible. |
| `width` | `integer` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `paragraph format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `MsoParagraphAlignment` | r/w | Returns or sets a value specifying the alignment of the paragraph. |
| `baseline alignment` | `MsoBaselineAlignment` | r/w | Returns or sets a constant that represents the vertical position of fonts in a paragraph. |
| `bullet` | `bullet format` | r/o | Returns a bullet format object for the paragraph. |
| `east asian line break level` | `boolean` | r/w | Returns or sets the East Asian line break control level for the specified paragraph. |
| `first line indent` | `real` | r/w | Returns or sets the value, in points, for a first line or hanging indent. |
| `hanging punctuation` | `boolean` | r/w | Determines whether hanging punctuation is enabled for the specified paragraphs. |
| `indent level` | `integer` | r/w | Returns or sets a value representing the indent level assigned to text in the selected paragraph. |
| `left indent` | `real` | r/w | Returns or sets a value that represents the left indent value, in points, for the specified paragraphs. |
| `line rule after` | `boolean` | r/w | Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines. |
| `line rule before` | `boolean` | r/w | Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines. |
| `line rule within` | `boolean` | r/w | Determines whether line spacing between base lines is set to a specific number of points or lines. |
| `right indent` | `real` | r/w | Returns or sets the right indent, in points, for the specified paragraphs. |
| `space after` | `real` | r/w | Returns or sets the amount of spacing, in points, after the specified paragraph. |
| `space before` | `real` | r/w | Returns or sets the spacing, in points, before the specified paragraphs. |
| `space within` | `real` | r/w | Returns or sets the amount of space between base lines in the specified paragraph, in points or lines. |
| `text direction` | `mTxD` | r/w | Returns or sets the text direction for the specified paragraph. |
| `word wrap` | `boolean` | r/w | Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs. |

### `picture format`

Contains properties and methods that apply to pictures.

*Plural:* `picture formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `brightness` | `real` | r/w | Returns or sets the brightness of the specified picture . The value for this property must be a number from 0.0, dimmest to 1.0, brightest. |
| `color type` | `MsoPictureColorType` | r/w | Returns or sets the type of color transformation applied to the specified picture. |
| `contrast` | `real` | r/w | Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast. |
| `crop bottom` | `real` | r/w | Returns or sets the number of points that are cropped off the bottom of the specified picture. |
| `crop left` | `real` | r/w | Returns or sets the number of points that are cropped off the left side of the specified picture. |
| `crop right` | `real` | r/w | Returns or sets the number of points that are cropped off the right side of the specified picture. |
| `crop top` | `real` | r/w | Returns or sets the number of points that are cropped off the top of the specified picture. |
| `transparency color` | `cRGBColor` | r/w | Returns or sets the transparent color for the specified picture as aRGB color. For this property to take effect, the transparent background property must be set to true. |
| `transparent background` | `boolean` | r/w | Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent. |

### `picture`

*Plural:* `pictures`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `text` | r/o | Returns he name of the file that has the picture. |
| `link to file` | `boolean` | r/o | Returns if the picture is lined to the file. |
| `picture format` | `picture format` | r/o | Returns a picture format object for this picture. |
| `save with document` | `boolean` | r/o | Returns if the picture should be saved with the document. |

### `rectangle`

Represents a rectangle graphic object.

*Plural:* `rectangles`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `add indent` | `integer` | r/w | Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. |
| `auto scale font` | `boolean` | r/w | Returns or sets  if the text in the object changes font size when the object size changes. |
| `auto size` | `integer` | r/w | Returns or sets if the size of the specified object is changed automatically to fit text within its boundaries. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bottom right cell` | `integer` | r/o | Returns the bottom right cell of the range the control is occupying. |
| `caption` | `integer` | r/w | Returns or sets the caption for this object. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `font object` | `integer` | r/o | Returns a font object that represents the font of the specified object. |
| `formula` | `integer` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `height` | `integer` | r/w | Returns or set the height of the object. |
| `horizontal alignment` | `integer` | r/w | Returns or sets the horizontal alignment for the object. |
| `interior object` | `integer` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `integer` | r/w | Returns or sets the position of the specified object, in points. |
| `locked` | `integer` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `locked text` | `integer` | r/w | Returns or sets whether the control's text is locked for editing. |
| `name` | `integer` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `orientation` | `integer` | r/w | May also be a number value from -90 to 90 degrees. |
| `placement` | `integer` | r/w | Returns or sets how the object is placed on the worksheet. |
| `print object` | `integer` | r/w | Returns or sets if this object is printed. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `rounded corners` | `integer` | r/w | Returns or sets if the rectangle has rounded corners |
| `shadow` | `integer` | r/w | Returns or sets if the object has a shadow. |
| `string value` | `integer` | r/w | Returns or sets the text of the specified object. |
| `top` | `integer` | r/w | Returns the top position of the specified object, in points. |
| `top left cell` | `integer` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `vertical alignment` | `integer` | r/w | Returns or sets the vertical alignment of the object. |
| `visible` | `integer` | r/w | Returns or sets if the object is visible. |
| `width` | `integer` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `reflection format`

Represents the reflection effect in Office graphics.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `reflection type` | `MsoReflectionType` | r/w | Returns or sets the type of the reflection format object. |

### `ruler level`

*Plural:* `ruler levels`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `first margin` | `real` | r/w | Returns or sets the first-line indent for the specified outline level, in points. |
| `left margin` | `real` | r/w | Returns or sets the left indent for the specified outline level, in points. |

### `ruler`

Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.

**Contains:** `ruler level`

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
| `offset X` | `real` | r/w | Returns or sets the horizontal offset in points of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left. |
| `offset Y` | `real` | r/w | Returns or sets the vertical offset in points of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape. |
| `rotate with shape` | `boolean` | r/w | Returns or sets whether to rotate the shadow when rotating the shape. |
| `shadow style` | `MsoShadowStyle` | r/w | Returns or sets the style of shadow formatting to apply to a shape. |
| `shadow type` | `MsoShadowType` | r/w | Returns or sets the shape shadow type. |
| `size` | `real` | r/w | Returns or sets the width of the shadow. |
| `transparency` | `real` | r/w | Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear. |
| `visible` | `boolean` | r/w | Returns or sets if the specified object, or the formatting applied to it, is visible. |

### `shape connector`

*Plural:* `shape connectors`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `connector format` | `connector format` | r/o | Returns a connector format object that contains connector formatting properties. |
| `connector type` | `MsoConnectorType` | r/o | Returns the connector type. |

### `shape font`

Contains font attributes such as font name, size, and color, for an object.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `ASCII name` | `text` | r/w | Returns or sets the font used for Latin text; characters with character codes from 0 through 127. |
| `base line offset` | `real` | r/w | Returns or sets a value specifying the horizontaol offset of the selected font. |
| `bold` | `boolean` | r/w | Returns or sets a value specifying whether the font should be bold. |
| `caps type` | `MsoTextCaps` | r/w | Returns or sets a value specifying how the text should be capitalized. |
| `east asian name` | `text` | r/w | Returns or sets the font name used for Asian text. |
| `embedable` | `boolean` | r/o | Returns a value indicating whether the font can be embedded in a page. |
| `embedded` | `boolean` | r/o | Returns a value specifying whether the font is embedded in a page. |
| `equalize character height` | `boolean` | r/w | Returns or sets a value specifying whether the text should have the same horizontal height. |
| `fill format` | `fill format` | r/o | Returns a fill format object that contains fill formatting properties for the specified font. |
| `font color` | `integer` | r/w |  |
| `font color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified font. |
| `font name` | `text` | r/w | Returns or sets a value specifying the font to use for a selection. |
| `font name other` | `text` | r/w | Returns or sets the font used for characters whose character set numbers are greater than 127. |
| `font size` | `real` | r/w | Returns or sets the font size. |
| `glow format` | `glow format` | r/o | Returns the formatting properties for a glow effect. |
| `highlight color` | `integer` | r/w | Returns or sets the text highlight color for object. |
| `highlight color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the highlight color for the specified text. |
| `italic` | `boolean` | r/w | Returns or sets a value specifying whether the text for a selection is italic. |
| `kerning` | `real` | r/w | Returns or sets a value specifying the amount of spacing between text characters. |
| `line format` | `line format` | r/o | Returns a value specifiying the format of a line. |
| `reflection format` | `reflection format` | r/o | Returns the formatting properties for a reflection effect. |
| `shadow format` | `shadow format` | r/o | Returns the value specifying the type of shadow effect for the selection of text. |
| `soft edge type` | `MsoSoftEdgeType` | r/w | Returns or sets the type soft edge format object. |
| `spacing` | `real` | r/w | Returns or sets a value specifying the spacing between characters in a selection of text. |
| `strike type` | `MsoTextStrike` | r/w | Returns or sets a value specifying the strike format used for a selection of text. |
| `subscript` | `boolean` | r/w | Returns or sets a value specifying that the selected text should be displayed a subscript. |
| `superscript` | `boolean` | r/w | Returns or sets a value specifying that the selected text should be displayed a superscript. |
| `underline color` | `integer` | r/w | Returns a value specifying the color of the underline for the selected text. |
| `underline color theme index` | `MsoThemeColorIndex` | r/w | Returns a value specifying the color of the underline for the selected text. |
| `underline style` | `MsoTextUnderlineType` | r/w | Returns or sets a value specifying the text effect for the selected text. |
| `word art styles format` | `MsoPresetTextEffect` | r/w | Returns or sets a value specifying the text effect for the selected text. |

### `shape line`

The line shape uses begin line X, begin line Y, end line X, and end line Y when created

*Plural:* `shape lines`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `begin line X` | `real` | r/w | Returns or sets the beginning X position of the line. |
| `begin line Y` | `real` | r/w | Returns or sets the beginning Y position of the line. |
| `end line X` | `real` | r/w | Returns or sets the ending X position of the line. |
| `end line Y` | `real` | r/w | Returns or sets the ending Y position of the line. |

### `shape text frame`

Represents the shape text frame in a shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.

*Plural:* `shape text frames`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `has text` | `boolean` | r/o | Returns whether the specified text frame has text. |
| `horizontal anchor` | `MsoHorizontalAnchor` | r/w |  |
| `margin bottom` | `real` | r/w | Returns or sets the distance, in points, between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. |
| `margin left` | `real` | r/w | Returns or sets the distance, in points, between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. |
| `margin right` | `real` | r/w | Returns or sets the distance, in points, between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. |
| `margin top` | `real` | r/w | Returns or sets the distance, in points, between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. |
| `orientation` | `MsoTextOrientation` | r/w | Returns or sets the text orientation. |
| `path format` | `MsoPathFormat` | r/w | Returns or sets the path type for the specified text frame. |
| `ruler` | `ruler` | r/o |  |
| `text column` | `text column` | r/o | Returns the text column object that represents the columns within the text frame. |
| `text range` | `text range` | r/o |  |
| `threeD format` | `threeD format` | r/o | Returns the 3-D-effect formatting properties for the specified text. |
| `vertical anchor` | `MsoVerticalAnchor` | r/w |  |
| `warp format` | `MsoWarpFormat` | r/w | Returns or sets the warp type for the specified text frame. |
| `word wrap` | `boolean` | r/w | Returns or sets text break lines within or past the boundaries of the shape. |
| `wordart auto size` | `MsoAutoSize` | r/w | The size of the specified object that changes automatically to fit text within its boundaries. |

### `shape textbox`

*Plural:* `shape textboxes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `text orientation` | `MsoTextOrientation` | r/o | Returns the text orientation of the object. |

### `shape`

Represents an object in the drawing layer.

*Plural:* `shapes`

**Contains:** `adjustment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alternative text` | `text` | r/w | Returns or sets the descriptive alternative text string for a Shape object when the object is saved to a Web page. |
| `auto shape type` | `MsoAutoShapeType` | r/w | Returns or sets the shape type for the specified shape object, which must represent an auto-shape. |
| `background style` | `MsoBackgroundStyleIndex` | r/w | Returns or sets the background style. |
| `black white mode` | `MsoBlackWhiteMode` | r/w | Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. |
| `bottom right cell` | `range` | r/o | Returns a range object that represents the cell that lies under the lower-right corner of the object. |
| `chart` | `chart` | r/o | Returns a chart object that represents the chart contained in the shape. |
| `child` | `boolean` | r/o | True if the shape is a child shape. |
| `connection site count` | `integer` | r/o | Returns the number of connection sites on the specified shape. |
| `connector` | `boolean` | r/o | Returns true if the specified shape is a connector. |
| `connector format` | `connector format` | r/o | Returns a connector format object that contains connector formatting properties if this shape is a connector. |
| `connector type` | `MsoConnectorType` | r/o | Returns the connector type if this shape is a connector. |
| `fill format` | `fill format` | r/o | Returns a fill format object that contains fill formatting properties for the specified shape. |
| `glow format` | `glow format` | r/o | Returns the formatting properties for a glow effect. |
| `has chart` | `boolean` | r/o | True if the specified shape has a chart. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `horizontal flip` | `boolean` | r/o | Returns true if the specified shape is flipped around the horizontal axis. |
| `hyperlink` | `hyperlink` | r/o | Returns a hyperlink object that represents the hyperlink for the shape. |
| `left position` | `real` | r/w | Returns or sets the left position of the specified object, in points. |
| `line format` | `line format` | r/o | Returns a line format object for this shape. |
| `lock aspect ratio` | `boolean` | r/w | Returns or sets if the specified shape retains its original proportions when you resize it. If false, you can change the height and width of the shape independently of one another when you resize it. |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked. If false, the object can be modified when the sheet is protected. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `parentgroup` | `shape` | r/o | Returns a Shape object that represents the common parent shape of a child shape or a range of child shapes. |
| `placement` | `XlPlacement` | r/w | Returns or sets how the object is placed on the worksheet. |
| `reflection format` | `reflection format` | r/o | Returns the formatting properties for a reflection effect. |
| `rotation` | `real` | r/w | Returns or sets the rotation of the shape, in degrees. |
| `shadow format` | `shadow format` | r/o | Returns a shadow format object for this shape. |
| `shape on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `shape style` | `MsoShapeStyleIndex` | r/w | Returns or sets the shape style corresponding to the Quick Styles. |
| `shape text frame` | `shape text frame` | r/o | Returns a shape text frame object that contains the alignment and anchoring properties for the specified shape. |
| `shape type` | `MsoShapeType` | r/o | Returns the shape type. |
| `soft edge format` | `soft edge format` | r/o | Returns the formatting properties for a soft edge effect. |
| `text frame` | `text frame` | r/o | Returns a text frame object that contains the alignment and anchoring properties for the specified shape. |
| `threeD format` | `threeD format` | r/o | Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape. |
| `top` | `real` | r/w | Returns or sets the top position of the specified object, in points. |
| `top left cell` | `range` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `vertical flip` | `boolean` | r/o | Returns true if the specified shape is flipped around the vertical axis. |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `word art format` | `word art format` | r/o | Returns the formatting properties for a word art effect. |
| `z order position` | `integer` | r/o | Returns the position of the specified shape in the z-order. To set the shape's position in the z-order, use the z order method. |

### `soft edge format`

Represents the soft edge formatting for a shape or range of shapes

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `soft edge type` | `MsoSoftEdgeType` | r/w | Returns or sets the type soft edge format object. |

### `tab stop`

Represents a single tab stop.

*Plural:* `tab stops`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `tab position` | `real` | r/w | Returns or sets the position of a tab stop relative to the left margin. |
| `tab stop type` | `MsoTabStopType` | r/w | Returns or sets the type of the tab stop object. |

### `text column`

Represents a single text column.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `column number` | `integer` | r/w | Returns or sets the index of the text column object. |
| `spacing` | `real` | r/w | Returns or sets the spacing between text columns in a text column object. |
| `text direction` | `mTxD` | r/w | Returns or sets the direction of text in the text column object. |

### `text frame`

Represents the text frame in a shape object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.

*Plural:* `text frames`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `margin bottom` | `real` | r/w | Returns or sets the distance, in points, between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. |
| `margin left` | `real` | r/w | Returns or sets the distance, in points, between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. |
| `margin right` | `real` | r/w | Returns or sets the distance, in points, between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. |
| `margin top` | `real` | r/w | Returns or sets the distance, in points, between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. |
| `orientation` | `MsoTextOrientation` | r/w | Returns or sets the text orientation. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |

### `theme color scheme`

Represents the color scheme of an Office Theme

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

Represents a container for the font schemes of a Microsoft Office theme.

*Plural:* `theme fonts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | Returns or sets a value specifying the font to use for a selection. |

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
| `contour color` | `integer` | r/o | Returns or sets the color of the contour of an object or text. |
| `contour color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified contour. |
| `contour width` | `real` | r/w | Returns or sets the width of the contour of an object or text. |
| `depth` | `real` | r/w | Returns or sets the depth of the shape's extrusion. |
| `extrusion color` | `integer` | r/o | Returns or sets a RGB color that represents the color of the shape's extrusion. |
| `extrusion color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified extrusion. |
| `extrusion color type` | `MsoExtrusionColorType` | r/w | Returns or sets a value that indicates what will determine the extrusion color. |
| `field of view` | `real` | r/w | Returns or sets the amount of perspective for an object or text. |
| `light angle` | `real` | r/w | Returns or sets the angle of the lighting. |
| `perspective` | `boolean` | r/w | Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point. |
| `preset camera` | `MsoPresetCamera` | r/o | Returns a constant that represents the camera preset. |
| `preset extrusion direction` | `MsoPresetExtrusionDirection` | r/o | Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion. |
| `preset lighting direction` | `MsoPresetLightingDirection` | r/w | Returns or sets the position of the light source relative to the extrusion. |
| `preset lighting rig` | `MsoLightRigType` | r/w | Returns a constant that represents the lighting preset. |
| `preset lighting softness` | `MsoPresetLightingSoftness` | r/w | Returns or sets the intensity of the extrusion lighting. |
| `preset material` | `MsoPresetMaterial` | r/w | Returns or sets the extrusion surface material. |
| `preset threeD format` | `MsoPresetThreeDFormat` | r/o |  |
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
| `bold` | `boolean` | r/w | Returns if the text of the word art shape is formatted as bold. |
| `font name` | `text` | r/w | Returns or sets the font name of the font used by this word art shape. |
| `font size` | `real` | r/w | Returns or sets the font size of the font used by this word art shape. |
| `italic` | `boolean` | r/w | Returns if the text of the word art shape is formatted as italic. |
| `kerned pairs` | `boolean` | r/w | Returns or sets if character pairs in a WordArt object have been kerned. |
| `normalized height` | `boolean` | r/w | Returns or sets if all characters, both uppercase and lowercase, in the specified WordArt are the same height. |
| `preset shape` | `MsoPresetTextEffectShape` | r/w | Returns or sets the shape of the specified WordArt. |
| `preset word art effect` | `MsoPresetTextEffect` | r/w | Returns or sets the style of the specified WordArt. |
| `rotated chars` | `boolean` | r/w | Returns or sets if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. If false, characters in the specified WordArt retain their original orientation relative to the bounding shape. |
| `tracking` | `real` | r/w | Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Can be a value from zero through 5. |
| `word art text` | `text` | r/w | Returns or sets the text associated with this word art shape. |

### Enumerations

### `XlOartVerticalOverflow`

| Value | Description |
|-------|-------------|
| `overflow` |  |
| `clip` |  |
| `ellipsis` |  |

### `XlOartHorizontalOverflow`

| Value | Description |
|-------|-------------|
| `overflow` |  |
| `clip` |  |

### `custom drop`

| Value | Description |
|-------|-------------|
| `callout format` |  |
| `callout` |  |

### `custom length`

| Value | Description |
|-------|-------------|
| `callout format` |  |
| `callout` |  |

### `preset drop`

| Value | Description |
|-------|-------------|
| `callout format` |  |
| `callout` |  |

### `check spelling`

| Value | Description |
|-------|-------------|
| `rectangle` |  |
| `oval` |  |
| `arc` |  |

### `copy picture`

| Value | Description |
|-------|-------------|
| `line` |  |
| `rectangle` |  |
| `oval` |  |
| `arc` |  |
| `shape` |  |

### `cut`

| Value | Description |
|-------|-------------|
| `line` |  |
| `rectangle` |  |
| `oval` |  |
| `arc` |  |
| `shape` |  |

### `bring to front`

| Value | Description |
|-------|-------------|
| `line` |  |
| `rectangle` |  |
| `oval` |  |
| `arc` |  |

### `send to back`

| Value | Description |
|-------|-------------|
| `line` |  |
| `rectangle` |  |
| `oval` |  |
| `arc` |  |

### `copy object`

| Value | Description |
|-------|-------------|
| `line` |  |
| `rectangle` |  |
| `oval` |  |
| `arc` |  |
| `shape` |  |

---

## Text Suite

Classes and types for scripting text manipulations

### Commands

### `add periods to`

Adds period punctuation to the end of the text contained in text range object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `change case`

Changes the case of a text range object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `to` | `MsoTextChangeCase` | no |  |

### `copy text range`

Copies a text range object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `cut text range`

Removes a portion or all of the text from a range of text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `get border`

Returns the specified border object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `style` | no |  |
| `which border` | `XlBordersIndex` | no | This specifies which border object should be retrieved. |

**Returns:** `border`

### `get rotated text bounds`

Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

**Returns:** `any`

### `insert into`

Inserts a string preceding the selected characters.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `character` | no |  |
| `string` | `text` | no | The string to insert. |

### `insert text text range`

Adds new text to the text range object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `insert where` | `MsoTextRangeInsertPosition` | no | Sets whether the text will be inserted before or after the text range. |
| `new text` | `text` | no | Contains the text to be inserted. |

### `paste text range`

Pastes the contents of the Clipboard into the text range object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `remove periods from`

Removes all period punctuation from the text in the text range object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### Classes

### `character`

Represents characters in an object that contains text.

*Plural:* `characters`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `content` | `text` | r/w | Returns or sets the text for the specified object. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `phonetic characters` | `text` | r/w | Returns or sets the phonetic text in the specified characters object. |

### `font`

Contains font attributes, such as font name, size, and color, for an object.

*Plural:* `fonts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bold` | `boolean` | r/w | Returns or sets if the font is formatted as bold. |
| `color` | `integer` | r/w | Returns or sets the color for the font. |
| `font background` | `XlBackground` | r/w | Returns or sets the text background type. |
| `font color index` | `XlColorIndex` | r/w | Returns or sets the color index of the font. |
| `font size` | `integer` | r/w | Returns or sets the font size. |
| `font style` | `text` | r/w | Returns or sets the font style. |
| `italic` | `boolean` | r/w | Returns or sets if the font is formatted as italic. |
| `name` | `text` | r/w | Returns or sets the font name associated with this font object. |
| `outline font` | `boolean` | r/w | Returns or sets if the specified font is formatted as outline. |
| `shadow` | `boolean` | r/w | Returns or sets if the specified font is formatted as shadowed. |
| `strikethrough` | `boolean` | r/w | Returns or sets if the font is formatted as strikethrough text. |
| `subscript` | `boolean` | r/w | Returns or sets if the font is formatted as subscript. |
| `superscript` | `boolean` | r/w | Returns or sets if the font is formatted as superscript. |
| `theme color` | `XlThemeColor` | r/w | Returns or sets the theme color in the applied color scheme that is associated with the specified object. |
| `theme font index` | `XlThemeFont` | r/w | Returns or sets the theme font in the applied font scheme that is associated with the specified object. |
| `tint and shade` | `real` | r/w | Returns or sets a Single that lightens or darkens a color. |
| `underline` | `XlUnderlineStyle` | r/w | Returns or sets the type of underline applied to the font. |

### `paragraph`

*Plural:* `paragraphs`

### `sentence`

*Plural:* `sentences`

### `style`

Represents a style description for a range.

*Plural:* `styles`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `add indent` | `boolean` | r/w | Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. |
| `built in` | `boolean` | r/o | Returns true if the style is a built-in style. |
| `font object` | `font` | r/o | Returns the font object associated with this style. |
| `formula hidden` | `boolean` | r/w | Returns or sets if the formula will be hidden when the worksheet is protected. |
| `horizontal alignment` | `XlHAlign` | r/w | Returns or sets the horizontal alignment for the object. |
| `include alignment` | `boolean` | r/w | Returns or sets if the style includes the add indent, horizontal alignment, vertical alignment, wrap text, and orientation properties. |
| `include border` | `boolean` | r/w | Returns or sets if the style includes the color, color index, line style, and weight border properties. |
| `include font` | `boolean` | r/w | Returns or sets if the style includes the background, bold, color, color index, font style, italic, name, outline font, shadow, size, strikethrough, subscript, superscript, and underline font properties. |
| `include number` | `boolean` | r/w | Returns or sets if the style includes the number format property. |
| `include patterns` | `boolean` | r/w | Returns or sets if the style includes the color, color index, invert if negative, pattern, pattern color, and pattern color index interior properties. |
| `include protection` | `boolean` | r/w | Returns or sets if the style includes the formula hidden and locked protection properties. |
| `indent level` | `integer` | r/w | Returns or sets the indent level for the style. Can be an integer from 0 to 15. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified style. |
| `locked` | `boolean` | r/w | Returns or sets if the range using this style is locked. |
| `merged cells` | `boolean` | r/w | Returns true if the style contains merged cells. |
| `name` | `text` | r/o | Returns the name of the style. |
| `name local` | `text` | r/o | Returns the name of the style, in the language of the user. |
| `number format` | `text` | r/w | Returns or sets the format code for the object. |
| `number format local` | `text` | r/w | Returns or sets the format code for the object as a string in the language of the user. |
| `orientation` | `XlOrientation` | r/w | May also be a number value from -90 to 90 degrees. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `shrink to fit` | `boolean` | r/w | Returns or sets if text automatically shrinks to fit in the available column width. |
| `value` | `text` | r/o | Return the name of the specified style. |
| `vertical alignment` | `XlVAlign` | r/w | Returns or sets the vertical alignment of the style. |
| `wrap text` | `boolean` | r/w | Returns or sets if Microsoft Excel wraps the text in the object. |

### `text flow`

*Plural:* `text flows`

### `text range character`

*Plural:* `text range characters`

### `text range line`

*Plural:* `text range lines`

### `text range`

Represents the text frame in a shape object.

**Contains:** `text range character`, `word`, `text range line`, `sentence`, `paragraph`, `text flow`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bound height` | `real` | r/o | Returns the height, in points, of the text bounding box for the specified text. |
| `bound left` | `real` | r/o | Returns the left coordinate, in points, of the text bounding box for the specified text. |
| `bound top` | `real` | r/o | Returns the top coordinate, in points, of the text bounding box for the specified text. |
| `bound width` | `real` | r/o | Returns the width, in points, of the text bounding box for the specified text. |
| `content` | `text` | r/w | Returns or sets the text in a text range. |
| `font` | `shape font` | r/o | Returns the character formatting for the object. |
| `paragraph format` | `paragraph format` | r/o | Returns the paragraph formatting for the specified text. |
| `text length` | `integer` | r/o | Returns the length of a text range. |
| `text start` | `integer` | r/o | Returns the starting point of the specified text range. |

### `word`

*Plural:* `words`

---

## Table Suite

Classes and types for scripting table manipulations

### Commands

### `activate object`

Activates the object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `add comment`

Adds a comment to the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `comment text` | `text` | yes | The comment text. |

**Returns:** `Excel comment` — The newly created comment object.

### `advanced filter`

Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `action` | `XlFilterAction` | no | The filter operation. |
| `criteria range` | `XlCategoryNames` | yes | The criteria range. If this argument is omitted, there are no criteria. |
| `copy to range` | `XlCategoryNames` | yes | The destination range for the copied rows if action is filter copy. Otherwise, this argument is ignored. |
| `unique` | `boolean` | yes | Set to true to filter unique records only. False to filter all records that meet the criteria. The default value is false. |

### `apply names`

Applies names to the cells in the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `names` | `text` | yes | A list of the names to be applied. If this argument is omitted, all names on the sheet are applied to the range. |
| `ignore relative absolute` | `boolean` | yes | Set to true to replace references with names, regardless of the reference types of either the names or references. False to replace absolute references with absolute names, relative references with relative names, and mixed references with mixed names. |
| `use row column names` | `boolean` | yes | Set to true to use the names of row and column ranges that contain the specified range if names for the range cannot be found. False to ignore the omit column and omit row arguments. The default value is true. |
| `omit column` | `boolean` | yes | Set to true to replace the entire reference with the row-oriented name. The column-oriented name can be omitted only if the referenced cell is in the same column as the formula and is within a row-oriented named range. The default value is true. |
| `omit row` | `boolean` | yes | Set to true to replace the entire reference with the column-oriented name. The row-oriented name can be omitted only if the referenced cell is in the same row as the formula and is within a column-oriented named range. The default value is true. |
| `order` | `XlApplyNamesOrder` | yes | Determines which range name is listed first when a cell reference is replaced by a row-oriented and column-oriented range name. |
| `append last` | `boolean` | yes | Set to true to replace the definitions of the names in names and also replace the definitions of the last names that were defined. False to replace the definitions of the names in names only. The default value is false. |

### `apply outline styles`

Applies outlining styles to the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `auto outline`

Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `autocomplete`

Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `string` | `text` | no | The string to complete. |

**Returns:** `text` — The completed string.

### `autofill`

Performs an autofill on the cells in the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `destination` | `XlCategoryNames` | no | The cells to be filled. The destination must include the source range. |
| `type` | `XlAutoFillType` | yes | Specifies the fill type. |

### `autofilter range`

Filters a list using the AutoFilter.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `field` | `integer` | yes | The integer offset of the field on which you want to base the filter, from the left of the list, the leftmost field is field one. |
| `criteria1` | `any` | yes | The criteria, a string, for example, 101. Use = to find blank fields, or use <> to find nonblank fields. If this argument is omitted, the criteria is all. |
| `operator` | `XlAutoFilterOperator` | yes | Specifies the operator to be used. |
| `criteria2` | `any` | yes | The second criteria. Used with criteria1 and operator to construct compound criteria. |
| `visible drop down` | `boolean` | yes | Set to true to display the AutoFilter drop-down arrow for the filtered field. False to hide the AutoFilter drop-down arrow for the filtered field. |

### `autofit`

Changes the width of the specified column range or the height of the specified row range to achieve the best fit.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `autoformat`

Automatically formats a range of cells, using a predefined format.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `format` | `XlRangeAutoFormat` | yes | The range format to be used. |
| `include number` | `boolean` | yes | Set to true to include number formats in the autoformat. The default value is true. |
| `font` | `boolean` | yes | set to true to include font formats in the autoformat. The default value is true. |
| `alignment` | `boolean` | yes | Set to true to include alignment in the autoformat. The default value is true. |
| `border` | `boolean` | yes | Set to true to include border formats in the autoformat. The default value is true. |
| `pattern` | `boolean` | yes | set to true to include pattern formats in the autoformat. The default value is true. |
| `width` | `boolean` | yes | Set to true to include column width and row height in the autoformat. The default value is true. |

### `border around`

Adds a border to a range and sets the color, line style, and weight properties for the new border.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `line style` | `XlLineStyle` | yes | The line style for the border. |
| `weight` | `XlBorderWeight` | yes | The border weight. |
| `color index` | `XlColorIndex` | yes | The border color, as an index into the current color palette. |
| `color` | `integer` | yes | The border color, as an RGB value. |

### `calculate`

Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `calculate row major order`

Calculates a specfied range of cells.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `check spelling`

Checks the spelling of an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `custom dictionary` | `text` | yes | A string that indicates the file name of the custom dictionary to examine if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used. |
| `ignore uppercase` | `boolean` | yes | Set to true to have Microsoft Excel ignore words that are all uppercase. False to have Microsoft Excel check words that are all uppercase. |
| `always suggest` | `boolean` | yes | Set to true to have Microsoft Excel display a list of suggested alternate spellings when an incorrect spelling is found. False to have Microsoft Excel wait for you to input the correct spelling. |

### `clear Excel comments`

Clears all cell comments from the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `clear contents`

Clears all cell comments from the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `clear hyperlinks`

clear all the hyperlinks in a range

### `clear outline`

Clears the outline for the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `clear range`

Clears the entire object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `clear range formats`

Clears the formatting of the object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `column differences`

Returns a range object that represents all the cells whose contents are different from the comparison cell in each column.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `comparison` | `XlCategoryNames` | no | . A single cell to compare to the specified range. |

**Returns:** `range`

### `consolidate`

Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `sources` | `text` | yes | The sources of the consolidation as a list of text reference strings in R1C1-style notation. The references must include the full path of sheets to be consolidated. |
| `consolidation function` | `XlConsolidationFunction` | yes | The consolidation function. |
| `top row` | `boolean` | yes | Set to true to consolidate data based on column titles in the top row of the consolidation ranges. False to consolidate data by position. The default value is false. |
| `left column` | `boolean` | yes | Set to true to consolidate data based on row titles in the left column of the consolidation ranges. False to consolidate data by position. The default value is false. |
| `create links` | `boolean` | yes | Set to true to have the consolidation use worksheet links. False to have the consolidation copy the data. The default value is false. |

### `copy picture`

Copies the selected object to the clipboard as a picture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `appearance` | `XlPictureAppearance` | yes | Specifies how the picture should be copied. |
| `format` | `XlCopyPictureFormat` | yes | The format of the picture. |

### `copy range`

Copies the range to the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `destination` | `range` | yes |  |

### `create names`

Creates names in the specified range, based on text labels in the sheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `top` | `boolean` | yes | Set to true to create names by using labels in the top row. The default value is false. |
| `left position` | `boolean` | yes | Set to true to create names by using labels in the left column. The default value is false. |
| `bottom` | `boolean` | yes | Set to true to create names by using labels in the bottom row. The default value is false. |
| `right` | `boolean` | yes | Set to true to create names by using labels in the right column. The default value is false. |

### `cut range`

Cuts the range to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `destination of cut` | `XlCategoryNames` | yes |  |

### `data series`

Creates a data series in the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `rowcol` | `XlRowCol` | yes | Can be the by rows or by columns to have the data series entered in rows or columns, respectively. If this argument is omitted, the size and shape of the range is used. |
| `data series type` | `XlDataSeriesType` | yes | Sets the type of the data series.  The default is data series linear. |
| `date` | `XlDataSeriesDate` | yes | If the data series type argument is chronological series, the date argument indicates the step date unit. The default value is series date day. |
| `increment` | `integer` | yes | The step value for the series. The default value is 1. |
| `stop` | `integer` | yes | The stop value for the series. If this argument is omitted, Microsoft Excel fills to the end of the range. |
| `trend` | `boolean` | yes | Set to true to create a linear trend or growth trend. False to create a standard data series. The default value is false. |

### `data table`

Creates a data table based on input values and formulas that you define on a worksheet.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `row input` | `cell` | yes | A single cell to use as the row input for your table. |
| `column input` | `cell` | yes | A single cell to use as the column input for your table. |

### `delete range`

Deletes the range

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `shift` | `XlDeleteShiftDirection` | yes | Specifies how to shift cells to replace deleted cells. |

### `dirty`

Designates a range to be recalculated when the next recalculation occurs.

### `fill down`

Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `fill left`

Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `fill right`

Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `fill up`

Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `find`

Finds specific information in a range, and returns a range object that represents the first cell where that information is found. Returns nothing if no match is found. Doesn't affect the selection or the active cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `what` | `text` | no | The data to search for. |
| `after` | `XlCategoryNames` | yes | The cell after which you want to search. Note that after must be a single cell in the range.  If this argument isn't specified, the search starts after the cell in the upper-left corner of the range. |
| `look in` | `XlFindLookIn` | yes | Specifies where the find method should look. |
| `look at` | `XlLookAt` | yes | Specifies the part that should be looked at. |
| `search order` | `XlSearchOrder` | yes | Specifies if the search should be rows or columns. |
| `search direction` | `XlSearchDirection` | yes | Specifies the search direction either next or previous. |
| `match case` | `boolean` | yes | Set to true to make the search case sensitive. |
| `match byte` | `boolean` | yes | Used only in the East Asian version of Microsoft Excel. Set to true to have double-byte characters match only double-byte characters. False to have double-byte characters match their single-byte equivalents. |

**Returns:** `range` — The range which the specified data was found.

### `find next`

Continues a search that was begun with the find method. Finds the next cell that matches those same conditions and returns a range object that represents that cell. Doesn't affect the selection or the active cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `after` | `XlCategoryNames` | yes | The cell after which you want to search. Note that after must be a single cell in the range.  If this argument isn't specified, the search starts after the cell in the upper-left corner of the range. |

**Returns:** `range` — The range which the specified data was found.

### `find previous`

Continues a search that was begun with the find method. Finds the previous cell that matches those same conditions and returns a range object that represents that cell. Doesn't affect the selection or the active cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `after` | `XlCategoryNames` | yes | The cell after which you want to search. Note that after must be a single cell in the range.  If this argument isn't specified, the search starts after the cell in the upper-left corner of the range. |

**Returns:** `range` — The range which the specified data was found.

### `function wizard`

Starts the Function Wizard for the upper-left cell of the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `get XML value`

Returns the value of the specified range as XML.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

**Returns:** `any`

### `get address`

Returns the range reference in the language of the macro.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `row absolute` | `boolean` | yes | Set to true to return the row part of the reference as an absolute reference. The default value is true. |
| `column absolute` | `boolean` | yes | Set to true to return the column part of the reference as an absolute reference. The default value is true. |
| `reference style` | `XlReferenceStyle` | yes | Specifies what reference style should be returned. The default is A1. |
| `external` | `boolean` | yes | Set to true to return an external reference. False to return a local reference. The default value is false. |
| `relative to` | `XlCategoryNames` | yes | If row absolute and column absolute are false, and reference style is R1C1, you must include a starting point for the relative reference. This argument is a range object that defines the starting point. |

**Returns:** `text` — The address string

### `get address local`

Returns the range reference in the language of the user.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `row absolute` | `boolean` | yes | Set to true to return the row part of the reference as an absolute reference. The default value is true. |
| `column absolute` | `boolean` | yes | Set to true to return the column part of the reference as an absolute reference. The default value is true. |
| `reference style` | `XlReferenceStyle` | yes | Specifies what reference style should be returned. The default is A1. |
| `external` | `boolean` | yes | Set to true to return an external reference. False to return a local reference. The default value is false. |
| `relative to` | `XlCategoryNames` | yes | If row absolute and column absolute are false, and reference style is R1C1, you must include a starting point for the relative reference. This argument is a range object that defines the starting point. |

**Returns:** `text` — The address string

### `get border`

Returns the specified border object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `which border` | `XlBordersIndex` | no | This specifies which border object should be retrieved. |

**Returns:** `border`

### `get end`

Returns a range object that represents the cell at the end of the region that contains the source range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `direction` | `XlDirection` | no | The direction in which to move. |

**Returns:** `range`

### `get offset`

Returns a range object that represents a range that's offset from the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `row offset` | `integer` | yes | The number of rows, positive, negative, or zero, by which the range is to be offset. The default value is 0. |
| `column offset` | `integer` | yes | The number of columns, positive, negative, or zero, by which the range is to be offset. The default value is 0. |

**Returns:** `range`

### `get resize`

Resizes the specified range. Returns a Range object that represents the resized range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `row size` | `integer` | yes | The number of rows in the new range. If this argument is omitted, the number of rows in the range remains the same. |
| `column size` | `integer` | yes | The number of columns in the new range. If this argument is omitted, the number of columns in the range remains the same. |

**Returns:** `range` — A range object that represents the resized range.

### `goal seek`

Calculates the values necessary to achieve a specific goal. If the goal is an amount returned by a formula, this calculates a value that, when supplied to your formula, causes the formula to return the number you want. Returns True if successful.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `goal` | `real` | no | The value you want returned in this cell. |
| `changing cell` | `range` | no | Specifies which cell should be changed to achieve the target value. |

**Returns:** `boolean` — Specifies if the goal was met.

### `group`

When the Range object represents a single cell in a pivot table field's data range, the group method performs numeric or date-based grouping in that field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `start` | `integer` | yes | The first value to be grouped. If this argument is omitted or true, the first value in the field is used. |
| `end` | `integer` | yes | The last value to be grouped. If this argument is omitted or true, the last value in the field is used. |
| `by` | `integer` | yes | If the field is numeric, this argument specifies the size of each group. |
| `periods` | `boolean` | yes | An array of boolean values that specify the period for the group, as shown in the following list item 1 seconds, 2 minutes, 3 hours, 4 days, 5 months, 6 quarters, 7 years.  It the items is set to true, a group is created for the corresponding time. |

**Returns:** `boolean` — Specifies success or failure.

### `insert indent`

Adds an indent to the specified range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `insert amount` | `integer` | no | The amount to be added to the current indent. |

### `insert into range`

Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `shift` | `XlInsertShiftDirection` | yes | Specifies which way to shift the cells. |

### `justify`

Rearranges the text in a range so that it fills the range evenly.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `list names`

Pastes a list of all non-hidden names onto the worksheet, beginning with the first cell in the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `merge`

Creates a merged cell from the specified Range object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `across` | `boolean` | yes |  |

### `navigate arrow`

Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a range object that represents the new selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `toward precedent` | `boolean` | yes | Specifies the direction to navigate: True to navigate toward precedents, False to navigate toward dependent. |
| `arrow number` | `integer` | yes | Specifies the arrow number to navigate, corresponds to the numbered reference in the cell's formula. |
| `link number` | `integer` | yes | If the arrow is an external reference arrow, this argument indicates which external reference to follow. If this argument is omitted, the first external reference is followed. |

### `parse`

Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `parse line` | `text` | yes | A string that contains left and right brackets to indicate where the cells should be split. |
| `destination` | `XlCategoryNames` | yes | A range object that represents the upper-left corner of the destination range for the parsed data. If this argument is omitted, Microsoft Excel parses in place. |

### `paste special`

Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `what` | `XlPasteType` | yes | The part of the range to be pasted. |
| `operation` | `XlPasteSpecialOperation` | yes | The paste operation. |
| `skip blanks` | `boolean` | yes | Set to true to have blank cells in the range on the clipboard not be pasted into the destination range. The default value is false. |
| `transpose` | `boolean` | yes | set to true to transpose rows and columns when the range is pasted. The default value is false. |

### `print out`

Prints the object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `from` | `integer` | yes | The number of the page at which to start printing. If this argument is omitted, printing starts at the beginning. |
| `to` | `integer` | yes | The number of the last page to print. If this argument is omitted, printing ends with the last page. |
| `copies` | `integer` | yes | The number of copies to print. If this argument is omitted, one copy is printed. |
| `preview` | `boolean` | yes | Set to true to have Microsoft Excel invoke print preview before printing the object. |
| `active printer` | `text` | yes | Sets the name of the active printer. |
| `print to file` | `boolean` | yes | Set to true to print to a file. |
| `collate` | `boolean` | yes | Set to true to collate multiple copies. |

### `print preview`

Shows a preview of the object as it would look when printed. This function has been deprecated.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `enable changes` | `boolean` | yes | Controls access to the page setup dialog and the ability to change margins from the preview window by enabling or disabling the setup and margins buttons, respectively. |

### `remove duplicates`

Removes duplicate values from a range of values.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `remove subtotal`

Removes subtotals from a list.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `replace`

Finds and replaces characters in cells within a range. Doesn't change the selection or the active cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `what` | `text` | no | The string to search for. |
| `replacement` | `text` | no | The replacement string. |
| `look at` | `XlLookAt` | yes | Specifies the part that should be looked at. |
| `search order` | `XlSearchOrder` | yes | Specifies if the search should be rows or columns. |
| `match case` | `boolean` | yes | Set to true to make the search case sensitive. |
| `match byte` | `boolean` | yes | Used only in the East Asian version of Microsoft Excel. Set to true to have double-byte characters match only double-byte characters. False to have double-byte characters match their single-byte equivalents. |
| `match control characters` | `boolean` | yes | Used only in the East Asian version of Microsoft Excel. Set to true to have double-byte characters match only double-byte characters. False to have double-byte characters match their single-byte equivalents. |
| `match diacritics` | `boolean` | yes | Used only in the East Asian version of Microsoft Excel. Set to true to have double-byte characters match only double-byte characters. False to have double-byte characters match their single-byte equivalents. |
| `formula version` | `XlFormulaVersion` | yes | Specifies if the search/replace should be using formula or formula2 |

**Returns:** `boolean` — Specifies success or failure.

### `row differences`

Returns a range object that represents all the cells whose contents are different from those of the comparison cell in each row.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `comparison` | `cell` | no | A single cell to compare with the specified range. |

**Returns:** `range` — The result range.

### `run VB macro`

Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `arg1` | `any` | yes | An argument to be passed to the macro. |
| `arg2` | `any` | yes | An argument to be passed to the macro. |
| `arg3` | `any` | yes | An argument to be passed to the macro. |
| `arg4` | `any` | yes | An argument to be passed to the macro. |
| `arg5` | `any` | yes | An argument to be passed to the macro. |
| `arg6` | `any` | yes | An argument to be passed to the macro. |
| `arg7` | `any` | yes | An argument to be passed to the macro. |
| `arg8` | `any` | yes | An argument to be passed to the macro. |
| `arg9` | `any` | yes | An argument to be passed to the macro. |
| `arg10` | `any` | yes | An argument to be passed to the macro. |
| `arg11` | `any` | yes | An argument to be passed to the macro. |
| `arg12` | `any` | yes | An argument to be passed to the macro. |
| `arg13` | `any` | yes | An argument to be passed to the macro. |
| `arg14` | `any` | yes | An argument to be passed to the macro. |
| `arg15` | `any` | yes | An argument to be passed to the macro. |
| `arg16` | `any` | yes | An argument to be passed to the macro. |
| `arg17` | `any` | yes | An argument to be passed to the macro. |
| `arg18` | `any` | yes | An argument to be passed to the macro. |
| `arg19` | `any` | yes | An argument to be passed to the macro. |
| `arg20` | `any` | yes | An argument to be passed to the macro. |
| `arg21` | `any` | yes | An argument to be passed to the macro. |
| `arg22` | `any` | yes | An argument to be passed to the macro. |
| `arg23` | `any` | yes | An argument to be passed to the macro. |
| `arg24` | `any` | yes | An argument to be passed to the macro. |
| `arg25` | `any` | yes | An argument to be passed to the macro. |
| `arg26` | `any` | yes | An argument to be passed to the macro. |
| `arg27` | `any` | yes | An argument to be passed to the macro. |
| `arg28` | `any` | yes | An argument to be passed to the macro. |
| `arg29` | `any` | yes | An argument to be passed to the macro. |
| `arg30` | `any` | yes | An argument to be passed to the macro. |

**Returns:** `text` — The result of the macro.

### `run XLM macro`

Runs a macro or calls a function. This can be used to run a macro written in the Microsoft Excel 4.0 macro language, or to run a function in a DLL or XLL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `arg1` | `any` | yes | An argument to be passed to the macro. |
| `arg2` | `any` | yes | An argument to be passed to the macro. |
| `arg3` | `any` | yes | An argument to be passed to the macro. |
| `arg4` | `any` | yes | An argument to be passed to the macro. |
| `arg5` | `any` | yes | An argument to be passed to the macro. |
| `arg6` | `any` | yes | An argument to be passed to the macro. |
| `arg7` | `any` | yes | An argument to be passed to the macro. |
| `arg8` | `any` | yes | An argument to be passed to the macro. |
| `arg9` | `any` | yes | An argument to be passed to the macro. |
| `arg10` | `any` | yes | An argument to be passed to the macro. |
| `arg11` | `any` | yes | An argument to be passed to the macro. |
| `arg12` | `any` | yes | An argument to be passed to the macro. |
| `arg13` | `any` | yes | An argument to be passed to the macro. |
| `arg14` | `any` | yes | An argument to be passed to the macro. |
| `arg15` | `any` | yes | An argument to be passed to the macro. |
| `arg16` | `any` | yes | An argument to be passed to the macro. |
| `arg17` | `any` | yes | An argument to be passed to the macro. |
| `arg18` | `any` | yes | An argument to be passed to the macro. |
| `arg19` | `any` | yes | An argument to be passed to the macro. |
| `arg20` | `any` | yes | An argument to be passed to the macro. |
| `arg21` | `any` | yes | An argument to be passed to the macro. |
| `arg22` | `any` | yes | An argument to be passed to the macro. |
| `arg23` | `any` | yes | An argument to be passed to the macro. |
| `arg24` | `any` | yes | An argument to be passed to the macro. |
| `arg25` | `any` | yes | An argument to be passed to the macro. |
| `arg26` | `any` | yes | An argument to be passed to the macro. |
| `arg27` | `any` | yes | An argument to be passed to the macro. |
| `arg28` | `any` | yes | An argument to be passed to the macro. |
| `arg29` | `any` | yes | An argument to be passed to the macro. |
| `arg30` | `any` | yes | An argument to be passed to the macro. |

**Returns:** `text` — The result of the macro.

### `set XML value`

Set the value of the specified range using XML.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `range value` | `any` | no |  |

### `show`

Scrolls through the contents of the active window to move the range into view. The range must consist of a single cell in the active document.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `show dependents`

Draws tracer arrows to the direct dependents of the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `remove` | `boolean` | yes | Set to true to remove one level of tracer arrows to direct dependents. False to expand one level of tracer arrows. The default value is false. |

### `show errors`

Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `show precedents`

Draws tracer arrows to the direct precedents of the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `remove` | `boolean` | yes | Set to true to remove one level of tracer arrows to direct precedents. False to expand one level of tracer arrows. The default value is false. |

### `sort`

Sorts a pivot table, a range, or the current region, if the specified range contains only one cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `key1` | `range` | yes | The first sort field. |
| `order1` | `XlSortOrder` | yes | Specifies the sort order for the first sort field. |
| `key2` | `range` | yes | The second sort field. |
| `sort type` | `XlSortType` | yes | Specifies which elements are sorted. |
| `order2` | `XlSortOrder` | yes | Specifies the sort order for the second sort field. |
| `key3` | `range` | yes | The third sort field. |
| `order3` | `XlSortOrder` | yes | Specifies the sort order for the third sort field. |
| `header` | `XlYesNoGuess` | yes | Specifies whether the first row contains headers. |
| `order custom` | `integer` | yes | A 1-based integer offset into the list of custom sort orders. If this argument is omitted, 1, normal, is used. |
| `match case` | `boolean` | yes | Set to true to do a case-sensitive sort. |
| `orientation` | `XlSortOrientation` | yes | Specifies if the sort is done top to bottom or left to right. |
| `sort method` | `XlSortMethod` | yes | Specifies the type of sort.  The default is pin yin. |
| `ignore control characters` | `boolean` | yes | It is not supported any more. |
| `ignore diacritics` | `boolean` | yes | It is not supported any more. |
| `dataoption1` | `XlSortDataOption` | yes | Specifies how to sort text in the range specified in Key1. Does not apply to pivot table sorting. |
| `dataoption2` | `XlSortDataOption` | yes | Specifies how to sort text in the range specified in Key2. Does not apply to pivot table sorting. |
| `dataoption3` | `XlSortDataOption` | yes | Specifies how to sort text in the range specified in Key3. Does not apply to pivot table sorting. |

### `sort special`

Uses East Asian sorting methods to sort the range, or uses the current region if the range contains only one cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `sort method` | `XlSortMethod` | yes | Specifies how to sort. |
| `key1` | `range` | yes | The first sort field. |
| `order1` | `XlSortOrder` | yes | Specifies the sort order for the first sort field. |
| `type` | `any` | yes | Specifies which elements are sorted. |
| `key2` | `range` | yes | The second sort field. |
| `order2` | `XlSortOrder` | yes | Specifies the sort order for the second sort field. |
| `key3` | `range` | yes | The third sort field. |
| `order3` | `XlSortOrder` | yes | Specifies the sort order for the third sort field. |
| `header` | `XlYesNoGuess` | yes | Specifies whether the first row contains headers. |
| `order custom` | `integer` | yes | A 1-based integer offset into the list of custom sort orders. If this argument is omitted, 1, normal, is used. |
| `match case` | `boolean` | yes | Set to true to do a case-sensitive sort. |
| `orientation` | `XlSortOrientation` | yes | Specifies if the sort is done top to bottom or left to right. |
| `dataoption1` | `XlSortDataOption` | yes | Specifies how to sort text in the range specified in Key1. Does not apply to pivot table sorting. |
| `dataoption2` | `XlSortDataOption` | yes | Specifies how to sort text in the range specified in Key2. Does not apply to pivot table sorting. |
| `dataoption3` | `XlSortDataOption` | yes | Specifies how to sort text in the range specified in Key3. Does not apply to pivot table sorting. |

### `special cells`

Returns a Range object that represents all the cells that match the specified type and value.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `type` | `XlCellType` | no | Specifies the cells to include. |
| `value` | `XlSpecialCellsValue` | yes | If type is either cell type constants or cell type formulas, this argument is used to determine which types of cells to include in the result.  The default is to select all constants or formulas, no matter what the type. |

**Returns:** `range` — A range containing the special cells.

### `subtotal`

Creates subtotals for the range, or the current region, if the range is a single cell.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `group by` | `integer` | no | The field to group by, as a 1-based integer offset. |
| `function` | `XlConsolidationFunction` | no | The subtotal function. |
| `total list` | `integer` | no | An array of 1-based field offsets, indicating the fields to which the subtotals are added. |
| `replace` | `boolean` | yes | Set to true to replace existing subtotals. The default value is false |
| `page breaks` | `boolean` | yes | Set to true to add page breaks after each group. The default value is false. |
| `summary below data` | `XlSummaryRow` | yes | Specifies if the summary data id placed above or below the range.  The default is below. |

### `text to columns`

Parses a column of cells that contain text into several columns.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |
| `destination` | `XlCategoryNames` | yes | A range object that specifies where Microsoft Excel will place the results. If the range is larger than a single cell, the top left cell is used. |
| `data type` | `XlTextParsingType` | yes | The format of the text to be split into columns. |
| `text qualifier` | `XlTextQualifier` | yes | Specifies the text qualifier. |
| `consecutive delimiter` | `boolean` | yes | Set to true to have Microsoft Excel consider consecutive delimiters as one delimiter. The default value is false. |
| `tab` | `boolean` | yes | Set to true to have data type be delimited and to have the tab character be a delimiter. The default value is false. |
| `semicolon` | `boolean` | yes | Set to true to have data type be delimited and to have the semicolon character be a delimiter. The default value is false. |
| `comma` | `boolean` | yes | Set to true to have data type be delimited and to have the comma character be a delimiter. The default value is false. |
| `space` | `boolean` | yes | Set to true to have data type be delimited and to have the space character be a delimiter. The default value is false. |
| `use other` | `boolean` | yes | Set to true to have the character specified by the other char argument be the delimiter.  The data type argument must be delimited. The default is false. |
| `other char` | `text` | yes | This is required if the use other argument is true.  Specifies the delimiter character when Other is true. If more than one character is specified, only the first character of the string is used, the remaining characters are ignored. |
| `field info` | `any` | yes | A list contain parse information for the individual columns of data. Formats are general format, text format, MDY format, DMY format, YMD format, MYD format, DYM format, YDM format, skip column. |
| `decimal separator` | `text` | yes | A string specifying whether a comma or period is used in the text file as the separator for decimal numbers. |
| `thousands separator` | `text` | yes | A string specifying whether a comma,  period, or apostrophe is used in the text file as the separator for thousands. |

### `ungroup`

Promotes a range in an outline, that is, decreases its outline level. The specified range must be a row or column, or a range of rows or columns. If the range is in a pivot table, this method ungroups the items contained in the range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### `unmerge`

Separates a merged area into individual cells.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `range` | no |  |

### Classes

### `cell`

*Plural:* `cells`

### `column`

*Plural:* `columns`

### `range`

Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.

*Plural:* `ranges`

**Contains:** `range`, `cell`, `row`, `column`, `character`, `format condition`, `hyperlink`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `Excel comment` | `Excel comment` | r/o | Returns a comment object that represents the comment associated with the cell in the upper-left corner of the range. |
| `add indent` | `boolean` | r/w | Returns or sets if text is automatically indented when the text alignment in a cell is set to equal distribution either horizontally or vertically. |
| `areas` | `specifier` | r/o | Returns a list of  range objects  that represents all the ranges in a multiple-area selection. |
| `column width` | `real` | r/w | Returns or sets the width of all columns in the specified range. |
| `count large` | `null` | r/o | Counts the largest value in a given range of values. Read-only Variant. |
| `current array` | `range` | r/o | If the specified cell is part of an array, returns a range object that represents the entire array. |
| `current region` | `range` | r/o | Returns a range object that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns. |
| `dependents` | `range` | r/o | Returns a range object that represents the range containing all the dependents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one dependent. |
| `direct dependents` | `range` | r/o | Returns a range object that represents the range containing all the direct dependents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one dependent. |
| `direct precedents` | `range` | r/o | Returns a range object that represents the range containing all the direct precedents of a cell. This can be a multiple selection, a union of Range objects,  if there's more than one precedent. |
| `display format` | `display format` | r/o | Returns the display format associated with this range. |
| `entire column` | `range` | r/o | Returns a range object that represents the entire column, or columns,  that contains the specified range. |
| `entire row` | `range` | r/o | Returns a range object that represents the entire row, or rows,  that contains the specified range. |
| `first column index` | `integer` | r/o | Returns the number of the first column in the first area in the specified range. |
| `first row index` | `integer` | r/o | Returns the number of the first row of the first area in the range. |
| `font object` | `font` | r/o | Returns the font object associated with this range. |
| `formula` | `any` | r/w | Returns or sets the object's formula, in A1-style notation. |
| `formula array` | `text` | r/w | Returns or sets the array formula of a range. |
| `formula hidden` | `boolean` | r/w | Returns or sets if the formula will be hidden when the workbook or worksheet is protected. |
| `formula label` | `XlFormulaLabel` | r/w | Returns or sets the formula label type for the specified range. |
| `formula local` | `text` | r/w | Returns or sets the formula for the object, using A1-style references in the language of the user. |
| `formula r1c1` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation. |
| `formula r1c1 local` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation in the language of the user. |
| `formula2` | `any` | r/w | Returns or sets the object's formula, in A1-style notation. |
| `formula2 local` | `text` | r/w | Returns or sets the formula for the object, using A1-style references in the language of the user. |
| `formula2 r1c1` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation. |
| `formula2 r1c1 local` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation in the language of the user. |
| `has array` | `boolean` | r/o | Returns true if the specified cell is part of an array formula. |
| `has formula` | `boolean` | r/o | Returns true if all cells in the range contain formulas. |
| `height` | `real` | r/o | Returns the height of the range. |
| `hidden` | `boolean` | r/w | Returns or sets if the rows or columns are hidden. The specified range must span an entire column or row. |
| `horizontal alignment` | `XlHAlign` | r/w | Returns or sets the horizontal alignment for the range. |
| `indent level` | `integer` | r/w | Returns or sets the indent level for the range or style. Can be an integer from 0 to 15. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/o | The distance from the left edge of column A to the left edge of the range. If the range is discontinuous, the first area is used. If the range is more than one column wide, the leftmost column in the range is used. |
| `list header rows` | `integer` | r/o | Returns the number of header rows for the specified range. |
| `list object` | `list object` | r/o | Returns the list object associated with this range. |
| `location in table` | `XlLocationInTable` | r/o | Returns an enumeration that describes the part of the pivot table that contains the upper-left corner of the specified range. |
| `locked` | `boolean` | r/w | Returns or sets  if the range is locked. |
| `merge area` | `range` | r/o | Returns a range object that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell. |
| `merge cells` | `boolean` | r/w | Returns or sets if the range contains merged cells. |
| `name` | `text` | r/w | Returns or sets the name of the range. |
| `named item` | `text` | r/o | Returns the named item associated with this range. |
| `number format` | `text` | r/w | Returns or sets the format code for the object. |
| `number format local` | `text` | r/w | Returns or sets the format code for the object as a string in the language of the user. |
| `outline level` | `integer` | r/w | Returns or sets the current outline level of the specified row or column |
| `page break` | `XlPageBreak` | r/w | Returns or sets the location of a page break. |
| `phonetic object` | `phonetic` | r/o | Returns the phonetic object, which contains information about a specific phonetic text string in a specified range. |
| `pivot field` | `pivot field` | r/o | Returns a pivot field object that represents the pivot field containing the upper-left corner of the specified range. |
| `pivot item` | `pivot item` | r/o | Returns a pivot item object that represents the pivot item containing the upper-left corner of the specified range. |
| `pivot table` | `pivot table` | r/o | Returns a pivot table object that represents the pivot table containing the upper-left corner of the specified range. |
| `precedents` | `range` | r/o | Returns a range object that represents all the precedents of a cell. This can be a multiple selection, a union of Range objects, if there's more than one precedent. |
| `prefix character` | `text` | r/o | Returns the prefix character for the cell. |
| `query table` | `query table` | r/o | Returns a query table object that represents the query table that intersects the specified range. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `row height` | `real` | r/w | Returns  or sets the height of all the rows in the range specified, measured in points. |
| `show detail` | `boolean` | r/w | Returns or sets if the outline is expanded for the specified range, so that the detail of the column or row is visible. The specified range must be a single summary column or row in an outline. |
| `shrink to fit` | `boolean` | r/w | Returns or sets if text automatically shrinks to fit in the available column width. |
| `string value` | `any` | r/o | Returns the text for the range. |
| `style object` | `style` | r/w | Returns or sets a style object that represents the style of the specified range. |
| `summary` | `boolean` | r/o | Returns true if the range is an outlining summary row or column. The range should be a row or a column. |
| `text orientation` | `XlOrientation` | r/w | The text orientation. Can be a number value from -90 to 90 degrees. |
| `top` | `real` | r/o | The distance from the top edge of row 1 to the top edge of the range. If the range is discontinuous, the first area is used. If the range is more than one row high, the top, lowest numbered, row in the range is used. |
| `use standard height` | `boolean` | r/w | Returns or sets if the row height of the range object equals the standard height of the sheet. |
| `use standard width` | `boolean` | r/w | Returns or sets if the column width of the range object equals the standard width of the sheet. |
| `validation` | `validation` | r/o | Returns the validation object that represents data validation for the specified range. |
| `value` | `any` | r/w | Returns or sets the value of the range. |
| `value2` | `any` | r/w | Returns or sets the value of the range. The difference between this property and Value is that Value2 uses the numerical representation of cells that are formatted as currency and date. |
| `vertical alignment` | `XlVerticalAlignmentTarget` | r/w | Returns or sets the vertical alignment of the range. |
| `width` | `real` | r/o | Returns the width of the range. |
| `worksheet object` | `sheet` | r/o | Returns a worksheet object that represents the worksheet containing the specified range. |
| `wrap text` | `boolean` | r/w | Returns or sets if Microsoft Excel wraps the text in the object. |

### `row`

*Plural:* `rows`

---

## Proofing Suite

Classes and types for scripting text proofing

### Commands

### `add replacement`

Adds an entry to the array of autocorrect replacements.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `autocorrect` | no |  |
| `text to replace` | `text` | no | The text to be replaced. If this string already exists in the array of autocorrect replacements, the existing substitute text is replaced by the new text. |
| `replacement text` | `text` | no | The replacement text. |

### `delete replacement`

Deletes an entry from the array of autocorrect replacements.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `autocorrect` | no |  |
| `what` | `text` | no | The text to be replaced, as it appears in the row to be deleted from the array of autocorrect replacements. If this string doesn't exist in the array of autocorrect replacements, this method fails. |

### `get replacement list`

Returns a list of autocorrect replacements.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `autocorrect` | no |  |

**Returns:** `any`

### Classes

### `autocorrect`

Contains Microsoft Excel autocorrect attributes, capitalization of names of days, correction of two initial capital letters, automatic correction list, and so on.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `Automatically expand tables as I type` | `boolean` | r/w | Returns or sets if resizes the table to include data entered into a neighboring column or row. |
| `Automatically fill formulas` | `boolean` | r/w | Returns or sets if we automatically copies the formula to all the cells in the column when a formula is entered into a cell. |
| `correct caps lock` | `boolean` | r/w | Returns or sets if Microsoft Excel automatically corrects accidental use of the CAPS LOCK key. |
| `correct days` | `boolean` | r/w | Returns or sets if the first letter of day names is capitalized automatically. |
| `correct initial caps` | `boolean` | r/w | Returns or sets if words that begin with two capital letters are corrected automatically. |
| `correct sentence caps` | `boolean` | r/w | Returns or sets if Microsoft Excel automatically corrects sentence, first word, capitalization. |
| `replace text` | `boolean` | r/w | Returns or sets if text in the list of autocorrect replacements is replaced automatically. |

---

## Chart Suite

Classes and types for scripting charts

### Commands

### `activate object`

Activates the object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4038` | no |  |

### `apply custom chart type`

Applies a standard or custom chart type to a chart or series.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `chart type` | `XlChartGallery` | no | A standard chart type, see the chart type property for the list of available constants. For Chart objects, this argument can also be one of the options specified for this argument. |
| `chart name` | `text` | yes | The name of the custom chart type if chart type specifies a custom chart gallery. |

### `apply data labels`

Applies data labels to all the series in a chart, a single series or a series point.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4048` | no |  |
| `type` | `XlDataLabelsType` | yes | Specifies the data label type.  The default is 'data labels show value' |
| `legend key` | `boolean` | yes | Set to true to show the legend key next to the point. The default value is false. |
| `auto text` | `boolean` | yes | Set to true if the object automatically generates appropriate text based on content. |
| `has leader lines` | `boolean` | yes | True for chart and series point objects if the series has leader lines. |
| `show series name` | `boolean` | yes | Set to true to show the series name for the data label. |
| `show category name` | `boolean` | yes | Set to true to show the category name for the data label. |
| `show value` | `boolean` | yes | Set to true to show the value for the data label. |
| `show percentage` | `boolean` | yes | Set to true to show the percentage for the data label. |
| `show bubble size` | `boolean` | yes | Set to true to show the bubble size for the data label. |
| `separator` | `text` | yes | The separator for the data label. |

### `apply layout`

Applies a chart layout.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `layout` | `integer` | no | Specifies the type of layout, denoted by a number from 1 to 10. |
| `chart type` | `XlChartType` | yes | If specified, uses the layout from this chart type. For example, you can apply the layouts that are available from a line chart to a column chart. The layout only adds chart elements that are available for that particular chart type. |

### `bring to front`

Brings the object to the front of the z-order.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart object` | no |  |

### `chart location`

Moves the chart to a new location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `where` | `XlChartLocation` | no | Where to move the chart. |
| `name` | `text` | yes | The name of the sheet where the chart will be embedded if where is location as object or the name of the new sheet if where is location as new sheet. |

**Returns:** `chart`

### `chart one color gradient`

Sets the specified fill to a one-color gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4041` | no |  |
| `gradient style` | `MsoGradientStyle` | no | The gradient style. |
| `variant` | `integer` | no | The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the gradient tab in the fill effects dialog box. If gradient style is from center gradient, the variant argument can only be 1 or 2. |
| `degree` | `real` | no | The gradient degree. Can be a value from 0.0 dark, through 1.0 light. |

### `chart patterned`

Sets the specified fill to a pattern.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4040` | no |  |
| `pattern` | `MsoPatternType` | no | The pattern type. |

### `chart solid`

Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4045` | no |  |

### `chart two color gradient`

Sets the specified fill to a two-color gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4044` | no |  |
| `gradient style` | `MsoGradientStyle` | no | The gradient style. |
| `variant` | `integer` | no | The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the gradient tab in the fill effects dialog box. If gradient style is from center gradient, the variant argument can only be 1 or 2. |

### `chart user picture`

Fills the specified shape with an image.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4043` | no |  |
| `picture file` | `text` | no | The name of the picture file. |
| `picture format` | `XlChartPictureType` | yes | The picture format. |
| `picture stack unit` | `integer` | yes | The picture stack or scale unit, depends on the picture format argument. |
| `picture placement` | `XlChartPicturePlacement` | yes | The picture placement. |

### `chart user textured`

Fills the specified shape with small tiles of an image.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4042` | no |  |
| `texture file` | `text` | no | The name of the picture file. |

### `chart wizard`

Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `source` | `range` | yes | The range that contains the source data for the new chart. If this argument is omitted, Microsoft Excel edits the active chart sheet or the selected chart on the active worksheet. |
| `gallery` | `XlChartType` | yes | Specifies he chart type. |
| `format` | `integer` | yes | The option number for the built-in autoformats. Can be a number from 1 through 10, depending on the gallery type. If this argument is omitted, Microsoft Excel chooses a default value based on the gallery type and data source. |
| `plot by` | `XlRowCol` | yes | Specifies whether the data for each series is in rows or columns. |
| `category labels` | `integer` | yes | An integer specifying the number of rows or columns within the source range that contain category labels. Legal values are from zero through one less than the maximum number of the corresponding categories or series. |
| `series labels` | `integer` | yes | An integer specifying the number of rows or columns within the source range that contain series labels. Legal values are from zero through one less than the maximum number of the corresponding categories or series. |
| `has legend` | `boolean` | yes | Set to true to include a legend. |
| `title` | `text` | yes | Specifies the chart title text. |
| `category title` | `text` | yes | Specifies the category axis title text. |
| `value title` | `text` | yes | Specifies the value axis title text. |
| `extra title` | `text` | yes | Specifies the series axis title for 3-D charts or the second value axis title for 2-D charts. |

### `check spelling`

Checks the spelling of an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `custom dictionary` | `text` | yes | A string that indicates the file name of the custom dictionary to examine if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used. |
| `ignore uppercase` | `boolean` | yes | Set to true to have Microsoft Excel ignore words that are all uppercase. False to have Microsoft Excel check words that are all uppercase. |
| `always suggest` | `boolean` | yes | Set to true to have Microsoft Excel display a list of suggested alternate spellings when an incorrect spelling is found. False to have Microsoft Excel wait for you to input the correct spelling. |

### `clear`

Clear the object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4052` | no |  |

### `clear contents`

Clears the data from a chart but leaves the formatting.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart area` | no |  |

### `clear formats`

Clears the formatting of the object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4050` | no |  |

### `clear to match style`

Sets the chart formatting to the last predefined chart style applied.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |

### `copy chart as picture`

Copies the selected object to the clipboard as a picture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `appearance` | `XlPictureAppearance` | yes | Specifies how the picture should be copied. |
| `format` | `XlCopyPictureFormat` | yes | The format of the picture. |
| `output size` | `XlPictureAppearance` | yes | The size of the copied picture when the object is a chart on a chart sheet, not embedded on a worksheet. |

### `copy object`

Copies the object to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4049` | no |  |

### `copy picture`

Copies the selected object to the clipboard as a picture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart object` | no |  |
| `appearance` | `XlPictureAppearance` | yes | Specifies how the picture should be copied. |
| `format` | `XlCopyPictureFormat` | yes | The format of the picture. |

### `cut`

Cuts the object to the clipboard.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart object` | no |  |

### `delete object`

Deletes the object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4039` | no |  |

### `deselect`

Cancels the selection for the specified chart.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |

### `error bar`

Applies error bars to the series.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `series` | no |  |
| `direction` | `XlErrorBarDirection` | no | The error bar direction. |
| `include` | `XlErrorBarInclude` | no | The error bar parts to include. |
| `type` | `XlErrorBarType` | no | The error bar type. |
| `amount` | `real` | yes | The error amount. Used for only the positive error amount when type is error bar type custom. |
| `minus values` | `real` | yes | The negative error amount when type is error bar type custom. |

### `get axis`

Returns an object that represents a specified axis object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `axis type` | `XlAxisType` | no |  |
| `which axis` | `XlAxisGroup` | yes |  |

**Returns:** `axis`

### `get chart element`

Returns information about the chart element at specified X and Y coordinates. This method returns a list  of three items.  Please refer to the documentation on the meaning of the returned values.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `x` | `integer` | no | The X coordinate of the chart element. |
| `y` | `integer` | no | The Y coordinate of the chart element. |

**Returns:** `XlChartItem` — The list of returned values.

### `get has axis`

Returns which axes exist on the chart.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `axis type` | `XlAxisType` | yes | Specifies the axis type. |
| `axis group` | `XlAxisGroup` | yes | Specifies the axis group. |

**Returns:** `boolean` — Returns true if the specified axis exists.

### `paste`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4051` | no |  |

### `paste chart`

Pastes chart data from the clipboard into the specified chart.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `format` | `XlPasteType` | yes |  |

### `paste series`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `series` | no |  |

### `preset chart gradient`

Sets the specified fill to a preset gradient.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4047` | no |  |
| `gradient style` | `MsoGradientStyle` | no | The gradient style. |
| `variant` | `integer` | no | The gradient variant. Can be a value from 1 through 4, corresponding to one of the four variants on the gradient tab in the fill effects dialog box. If gradient style is from center gradient, the variant argument can only be 1 or 2. |
| `preset gradient type` | `MsoPresetGradientType` | no | The gradient type. |

### `preset chart textured`

Sets the specified fill format to a preset texture.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4046` | no |  |
| `preset texture for chart` | `MsoPresetTexture` | no | The preset texture. |

### `print out`

Prints the object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `from` | `integer` | yes | The number of the page at which to start printing. If this argument is omitted, printing starts at the beginning. |
| `to` | `integer` | yes | The number of the last page to print. If this argument is omitted, printing ends with the last page. |
| `copies` | `integer` | yes | The number of copies to print. If this argument is omitted, one copy is printed. |
| `preview` | `boolean` | yes | Set to true to have Microsoft Excel invoke print preview before printing the object. |
| `active printer` | `text` | yes | Sets the name of the active printer. |
| `print to file` | `boolean` | yes | Set to true to print to a file. |
| `collate` | `boolean` | yes | Set to true to collate multiple copies. |

### `print preview`

Shows a preview of the object as it would look when printed. This function is deprecated.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `enable changes` | `boolean` | yes | Controls access to the page setup dialog and the ability to change margins from the preview window by enabling or disabling the setup and margins buttons, respectively. |

### `protect chart`

Protects a chart so that it cannot be modified.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `password` | `text` | yes | A string that specifies a case-sensitive password for the sheet or workbook. If this argument is omitted, you can unprotect the sheet or workbook without using a password. Otherwise, you must specify the password to unprotect the sheet or workbook. |
| `drawing objects` | `boolean` | yes | Set to true to protect shapes. The default value is false. |
| `chart contents` | `boolean` | yes | Set to true to protect contents. The default value is true. |
| `user interface only` | `boolean` | yes | Set to true to protect the user interface, but not macros. If this argument is omitted, protection applies both to macros and to the user interface. |

### `refresh`

Updates the cache of the chart object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |

### `save as`

Saves changes into a different file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `filename` | `text` | no | A string that indicates the name of the file to be saved. You can include a full path, if you don't, Microsoft Excel saves the file in the current folder. |
| `file format` | `XlFileFormat` | yes | Specifies the file format to use when you save the file. |
| `password` | `text` | yes | A case-sensitive string, no more than 15 characters, that indicates the protection password to be given to the file. |
| `write reservation password` | `text` | yes | A string that indicates the write-reservation password for this file. If a file is saved with the password and the password isn't supplied when the file is opened, the file is opened as read-only. |
| `read only recommended` | `boolean` | yes | Set to true to display a message when the file is opened, recommending that the file be opened as read-only. |
| `create backup` | `boolean` | yes | Set to true to create a backup file. |
| `add to most recently used list` | `boolean` | yes | Set to true to add this workbook to the list of recently used files. The default value is false. |
| `save as local language` | `boolean` | yes | True saves files against the language of Microsoft Excel. False is the default, which saves files against the language of Visual Basic for Applications |

### `save as picture`

Saves the shape in the requested file using the stated graphic format

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart object` | no |  |
| `picture type` | `MsoPictureType` | yes | Specifies the graphic format in which the file is saved |
| `file name` | `text` | yes | The name and path for the picture being saved |

### `send to back`

Sends the object to the back of the z-order.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart object` | no |  |

### `set background picture`

Sets the background graphic for a chart.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `picture file name` | `text` | yes | The name of the graphic file. |

### `set chart element`

Sets chart elements on a chart.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `chart element` | `null` | no | Specifies the chart element type. |

### `set has axis`

Sets which axes exist on the chart.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `axis exists` | `boolean` | no | Specifies if the specified axis should exist. |
| `axis type` | `XlAxisType` | yes | Specifies the axis type. |
| `axis group` | `XlAxisGroup` | yes | Specifies the axis group. |

### `set source data`

Sets the source data range for the chart.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `source` | `range` | no | The range that contains the source data. |
| `plot by` | `XlRowCol` | yes | Specifies the way the data is to be plotted. |

### `unprotect`

Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `chart` | no |  |
| `password` | `text` | yes | A string that denotes the case-sensitive password to use to unprotect the sheet or workbook.  If you omit this argument for a sheet that's protected with a password, you'll be prompted for the password. |

### Classes

### `area group`

*Plural:* `area groups`

### `axis title`

Represents a chart axis title.

*Plural:* `axis titles`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets if the text in the object changes font size when the object size changes. |
| `axis title text` | `text` | r/w | Returns or sets the text associated with this object. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `caption` | `text` | r/w | Returns or sets the caption for this object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `horizontal alignment` | `XlHAlign` | r/w | Returns or sets the horizontal alignment for the object. |
| `include in layout` | `boolean` | r/w | Returns or sets if a axis title will occupy the chart layout space when a chart layout is being determined. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the left position of the specified object, in points. |
| `name` | `text` | r/o | Returns the name of the object. |
| `orientation` | `integer` | r/w | The text orientation.  Can be a number value from -90 to 90 degrees. |
| `position` | `XlChartElementPosition` | r/w | Returns or sets the position of the axis title on the chart. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `top` | `real` | r/w | Returns or sets the top position of the specified object, in points. |
| `vertical alignment` | `XlVerticalAlignmentTarget` | r/w | Returns or sets the vertical alignment of the object. |

### `axis`

Represents a single axis in a chart.

*Plural:* `axes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `axis between categories` | `boolean` | r/w | Returns or sets if the value axis crosses the category axis between categories. |
| `axis group` | `XlAxisGroup` | r/o | Returns the group for the specified axis. |
| `axis title` | `axis title` | r/o | Returns an axis title object that represents the title of the specified axis. |
| `axis type` | `XlAxisType` | r/w | Returns or sets the axis type |
| `base unit` | `XlTimeUnit` | r/w | Returns or sets the base unit for the specified category axis. |
| `base unit is auto` | `boolean` | r/w | Returns or sets if Microsoft Excel chooses appropriate base units for the specified category axis. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `category names` | `XlCategoryNames` | r/w | Returns or sets all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a range object that contains the category names. |
| `category type` | `XlCategoryType` | r/w | Returns or sets the category axis type. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the axis. |
| `crosses` | `XlAxisCrosses` | r/w | Returns or sets the point on the specified axis where the other axis crosses. |
| `crosses at` | `real` | r/w | Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. |
| `display unit` | `XlDisplayUnit` | r/w | Returns or sets the unit label for the specified axis. |
| `display unit custom` | `real` | r/w | If the value of the display unit property is custom display unit, the display unit custom property returns or sets the value of the displayed units. |
| `display unit label` | `display unit label` | r/o | Returns the display unit label object for the specified axis |
| `has display unit label` | `boolean` | r/w | Returns or sets if the label specified by the display unit or display unit custom property is displayed on the specified axis. The default value is true. |
| `has major gridlines` | `boolean` | r/w | Returns or sets if the axis has major gridlines. Only axes in the primary axis group can have gridlines. |
| `has minor gridlines` | `boolean` | r/w | Returns or sets if the axis has minor gridlines. Only axes in the primary axis group can have gridlines. |
| `has title` | `boolean` | r/w | Returns or sets if the axis or chart has a visible title. |
| `height` | `real` | r/o | Returns the height of the object. |
| `left position` | `real` | r/o | Returns the left position of the specified object, in points. |
| `log base` | `real` | r/w | Returns or sets the base of the logarithm when you are using log scales. |
| `major gridlines` | `gridlines` | r/o | Returns a gridlines object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines. |
| `major tick mark` | `XlTickMark` | r/w | Returns or sets the type of major tick mark for the specified axis. |
| `major unit` | `real` | r/w | Returns or sets the major units for the axis. |
| `major unit is auto` | `boolean` | r/w | Returns or sets if Microsoft Excel calculates the major units for the axis. |
| `major unit scale` | `XlTimeUnit` | r/w | Returns or sets the major unit scale value for the category axis when the category type property is set to time scale. |
| `maximum scale` | `real` | r/w | Returns or sets the maximum value on the axis. |
| `maximum scale is auto` | `boolean` | r/w | Returns or sets if Microsoft Excel calculates the maximum value for the axis. |
| `minimum scale` | `real` | r/w | Returns or sets the minimum value on the axis. |
| `minimum scale is auto` | `boolean` | r/w | Returns or sets if Microsoft Excel calculates the minimum value for the axis. |
| `minor gridlines` | `gridlines` | r/o | Returns a gridlines object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines. |
| `minor tick mark` | `XlTickMark` | r/w | Returns or sets the type of minor tick mark for the specified axis. |
| `minor unit` | `real` | r/w | Returns or sets the minor units on the axis. |
| `minor unit is auto` | `boolean` | r/w | Returns or sets if Microsoft Excel calculates minor units for the axis. |
| `minor unit scale` | `XlTimeUnit` | r/w | Returns or sets the minor unit scale value for the category axis when the category type property is set to time scale. |
| `reverse plot order` | `boolean` | r/w | Returns or sets if Microsoft Excel plots data points from last to first. |
| `scale type` | `XlScaleType` | r/w | Returns or sets the value axis scale type. |
| `tick label position` | `XlTickLabelPosition` | r/w | Describes the position of tick-mark labels on the specified axis. |
| `tick label spacing` | `integer` | r/w | Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes. |
| `tick label spacing is auto` | `boolean` | r/w | Returns or sets whether or not the tick label spacing is automatic. |
| `tick labels` | `tick labels` | r/o | Returns a tick labels object that represents the tick-mark labels for the specified axis. |
| `tick mark spacing` | `integer` | r/w | Returns or sets the number of categories or series between tick marks. Applies only to category and series axes. |
| `top` | `real` | r/o | Returns the top position of the specified object, in points. |
| `width` | `real` | r/o | Returns the width of the object. |

### `bar group`

*Plural:* `bar groups`

### `chart area`

Represents the chart area of a chart. The chart area on a 2-D chart contains the axes, the chart title, the axis titles, and the legend. The chart area on a 3-D chart contains the chart title and the legend; it doesn't include the plot area .

*Plural:* `chart areas`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets if the text in the object changes font size when the object size changes. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the chart area. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the left position of the specified object, in points. |
| `name` | `text` | r/o | Returns the name of the object. |
| `rounded corners` | `boolean` | r/o | Returns or sets if the chart area has rounded corners. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `top` | `real` | r/w | Returns or sets the top position of the specified object, in points. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |

### `chart fill format`

Used only with charts. Represents fill formatting for chart elements.

*Plural:* `chart fill formats`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `back color` | `integer` | r/o | Returns the background color. |
| `background scheme color` | `integer` | r/w | Returns or sets the background color as an index in the current color scheme. |
| `chart fill format type` | `MsoFillType` | r/o | The fill type. |
| `fore color` | `integer` | r/o | Returns the foreground color. |
| `foreground scheme color` | `integer` | r/w | Returns or sets the foreground color as an index in the current color scheme. |
| `gradient color type` | `MsoGradientColorType` | r/o | Returns the gradient color type for the specified fill. |
| `gradient degree` | `real` | r/o | Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 dark through 1.0 light. |
| `gradient style` | `MsoGradientStyle` | r/o | Returns the gradient style for the specified fill. |
| `gradient variant` | `integer` | r/o | Returns the shade variant for the specified fill as an integer value from 1 through 4. |
| `pattern` | `MsoPatternType` | r/o | Returns the fill pattern. |
| `preset gradient type` | `MsoPresetGradientType` | r/o | Returns the preset gradient type for the specified fill. |
| `preset texture` | `MsoPresetTexture` | r/o | Returns the preset texture for the specified fill. |
| `texture name` | `text` | r/o | Returns the name of the custom texture file for the specified fill. |
| `texture type` | `MsoTextureType` | r/o | Returns the texture type for the specified fill. |
| `visible` | `boolean` | r/w | Returns or sets if the specified object is visible. |

### `chart format`

Provides access to the Office Art formatting for chart elements.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `fill format` | `fill format` | r/o | Returns a fill format object for the parent chart element that contains fill formatting properties for the chart element. |
| `glow format` | `glow format` | r/o | Returns a glow format object for a specified chart that contains glow formatting properties for the chart. |
| `line format` | `line format` | r/o | Returns a line format object that contains line formatting properties for the specified chart element. |
| `shadow format` | `shadow format` | r/o | Returns a shadow format object that contains shadow formatting properties for the chart element. |
| `shape text frame` | `shape text frame` | r/o | Returns a shape text frame object that contains the alignment and anchoring properties for the specified chart. |
| `soft edge format` | `soft edge format` | r/o | Returns the formatting properties for a soft edge effect. |
| `threeD format` | `threeD format` | r/o | Returns a threeD format object that contains 3-D-effect formatting properties for the specified chart. |

### `chart group`

Represents one or more series plotted in a chart with the same format. A chart contains one or more chart groups, each chart group contains one or more series, and each series contains one or more points.

*Plural:* `chart groups`

**Contains:** `series`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `axis group` | `XlAxisGroup` | r/w |  |
| `bubble scale` | `integer` | r/w | Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from zero to 300, corresponding to a percentage of the default size. Applies only to bubble charts. |
| `doughnut hole size` | `integer` | r/w | Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent. |
| `down bars object` | `down bars` | r/o | Returns a down bars object that represents the down bars on a line chart. Applies only to line charts. |
| `drop lines object` | `drop lines` | r/o | Returns a drop lines object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `first slice angle` | `integer` | r/w | Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees, clockwise from vertical. Applies only to pie, 3-D pie, and doughnut charts. |
| `gap width` | `integer` | r/w | Returns or sets: For bar and column charts, the space between bar or column clusters, as a percentage of the bar or column width. For pie of pie and bar of pie charts, the space between the primary and secondary sections of the chart. |
| `has drop lines` | `boolean` | r/w | Returns or sets if the line chart or area chart has drop lines. Applies only to line and area charts. |
| `has hi lo lines` | `boolean` | r/w | Returns or sets if the line chart has high-low lines. Applies only to line charts. |
| `has radar axis labels` | `boolean` | r/w | Returns or sets if a radar chart has axis labels. Applies only to radar charts. |
| `has series lines` | `boolean` | r/w | Returns or sets if a stacked column chart or bar chart has series lines or if a pie of pie chart or bar of pie chart has connector lines between the two sections. Applies only to stacked column charts, bar charts, pie of pie charts, or bar of pie charts. |
| `has threeD shading` | `boolean` | r/w | Returns or sets if the chart group has three-dimensional shading. |
| `has up down bars` | `boolean` | r/w | Returns or sets if a line chart has up and down bars. Applies only to line charts. |
| `hilo lines Object` | `hilo lines` | r/o | Returns a hilo lines object that represents the high-low lines for a series on a line chart. |
| `overlap` | `integer` | r/w | Returns or sets how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2-D bar and 2-D column charts. |
| `radar axis labels` | `tick labels` | r/o | Returns a tick labels object that represents the radar axis labels for the specified chart group. |
| `second plot size` | `integer` | r/w | Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. |
| `series lines object` | `series lines` | r/o | Returns a series lines object that represents the series lines for a stacked bar chart or a stacked column chart. Applies only to stacked bar and stacked column charts. |
| `show negative bubbles` | `boolean` | r/w | Returns or sets if negative bubbles are shown for the chart group. Valid only for bubble charts. |
| `size represents` | `XlSizeRepresents` | r/w | Returns or sets what the bubble size represents on a bubble chart. |
| `split type` | `XlChartSplitType` | r/w | Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. |
| `split value` | `integer` | r/w | Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. |
| `up bars object` | `up bars` | r/o | Returns an up bars object that represents the up bars on a line chart. Applies only to line charts. |
| `vary by categories` | `boolean` | r/w | Returns or sets if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. |

### `chart object`

Represents an embedded chart on a worksheet. The chart object object acts as a container for a chart object. Properties and methods for the chart object object control the appearance and size of the embedded chart on the worksheet.

*Plural:* `chart objects`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bottom right cell` | `range` | r/o | Returns a range object that represents the cell that lies under the lower-right corner of the object. |
| `chart` | `chart` | r/o | Returns a chart object that represents the chart contained in the object. |
| `enabled` | `boolean` | r/w | Returns or sets if the object is enabled. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/w | Returns or set the height of the object. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the position of the specified object, in points. |
| `locked` | `boolean` | r/w | Returns or sets if the object is locked, if false the object can be modified when the sheet is protected. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `on action` | `text` | r/w | Returns or sets the name of a string of AppleScript commands that will be executed when the object is clicked on. Please note that if AppleScript commands are set they will not be saved with the document. |
| `placement` | `XlPlacement` | r/w | Returns or sets the way the object is attached to the cells below it. |
| `print object` | `boolean` | r/w | Returns or sets if this object is printed. |
| `protect chart object` | `boolean` | r/w | Returns or sets if the embedded chart frame cannot be moved, resized, or deleted. |
| `rounded corners` | `boolean` | r/w | Returns or sets if the embedded chart has rounded corners. |
| `shadow` | `boolean` | r/w | Returns or sets if the if the object has a shadow. |
| `top` | `real` | r/w | Returns or sets the distance from the top edge of the object to the top of row 1, on a worksheet, or the top of the chart area, on a chart. |
| `top left cell` | `range` | r/o | Returns a range object that represents the cell that lies under the upper-left corner of the specified object. |
| `visible` | `boolean` | r/w | Returns or sets if the object is visible. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |
| `z order position` | `integer` | r/o | Returns the z-order position of the object. |

### `chart sheet`

A chart sheet is a worksheet that contains only a chart.

*Plural:* `chart sheets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `chart` | `chart` | r/o | Returns the chart for this chart sheet. |
| `worksheet type` | `XlSheetType` | r/o | Returns the type of this worksheet. |

### `chart title`

Represents the chart title.

*Plural:* `chart titles`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets  if the text in the object changes font size when the object size changes. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `caption` | `text` | r/w | Returns or sets the chart title text. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the chart title. |
| `chart title text` | `text` | r/w | Returns or sets the text for the specified object. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `formula` | `text` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `formula local` | `text` | r/w | Returns or sets the formula for the object, using A1-style references in the language of the user. |
| `formula r1c1` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. |
| `formula r1c1 local` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation in the language of the user. |
| `height` | `real` | r/o | Returns the height of the object. |
| `horizontal alignment` | `XlHAlign` | r/w | Returns or sets the horizontal alignment for the object. |
| `include in layout` | `boolean` | r/w | Returns or sets if a chart title will occupy the chart layout space when a chart layout is being determined. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the left position of the specified object, in points. |
| `name` | `text` | r/o | Returns the name of the object. |
| `orientation` | `XlOrientation` | r/w | May also be a number value from -90 to 90 degrees. |
| `position` | `XlChartElementPosition` | r/w | Returns or sets the position of the chart title on the chart. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `top` | `real` | r/w | Returns or sets the top position of the specified object, in points. |
| `vertical alignment` | `XlVerticalAlignmentTarget` | r/w | Returns or sets the vertical alignment of the object. |
| `width` | `real` | r/o | Returns the width of the object. |

### `chart`

Represents a chart in a workbook. The chart can be either an embedded chart, contained in a chart object, or a separate chart sheet.

*Plural:* `charts`

**Contains:** `shape`, `arc`, `area group`, `bar group`, `button`, `chart group`, `chart object`, `checkbox`, `column group`, `doughnut group`, `dropdown`, `groupbox`, `hyperlink`, `label`, `line group`, `line`, `listbox`, `option button`, `oval`, `pie group`, `radar group`, `rectangle`, `scrollbar`, `series`, `spinner`, `textbox`, `xy group`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `area threeD group` | `chart group` | r/o | Returns a chart group object that represents the area chart group on a 3-D chart. |
| `auto scaling` | `boolean` | r/w | Returns or sets if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The chart's right angle axes property must be true. |
| `back wall` | `walls` | r/o | Returns a walls object that allows the user to individually format the back wall of a 3-D chart. |
| `bar shape` | `XlBarShape` | r/w | Returns or sets the shape used with the 3-D bar or column chart. |
| `bar threeD group` | `chart group` | r/o | Returns a chart group object that represents the bar chart group on a 3-D chart. |
| `chart area object` | `chart area` | r/o | Returns a chart area object that represents the complete chart area for the chart. |
| `chart style` | `integer` | r/w | Returns or sets the chart type. |
| `chart title` | `chart title` | r/o | Returns a chart title object that represents the title of the specified chart. |
| `chart type` | `XlChartType` | r/w | Returns or sets the chart type. |
| `column threeD group` | `chart group` | r/o | Returns a chart group object that represents the column chart group on a 3-D chart. |
| `corners object` | `corners` | r/o | Returns a corners object that represents the corners of a 3-D chart. |
| `data table object` | `data table` | r/o | Returns a data table object that represents the chart data table. |
| `depth percent` | `integer` | r/w | Returns or sets the depth of a 3-D chart as a percentage of the chart width, between 20 and 2000 percent. |
| `display blanks as` | `XlDisplayBlanksAs` | r/w | Returns or sets the way that blank cells are plotted on a chart. |
| `elevation` | `integer` | r/w | Returns or sets the elevation of the 3-D chart view, in degrees. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `floor object` | `floor` | r/o | Returns a floor object that represents the floor of the 3-D chart. |
| `gap depth` | `integer` | r/w | Returns or sets the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500. |
| `has data table` | `boolean` | r/w | Returns or sets if the chart has a data table. |
| `has legend` | `boolean` | r/w | Returns or sets if the chart has a legend. |
| `has title` | `boolean` | r/w | Returns or sets if the chart has a title. |
| `height percent` | `integer` | r/w | Returns or sets the height of a 3-D chart as a percentage of the chart width, between 5 and 500 percent. |
| `legend object` | `legend` | r/o | Returns a legend object that represents the legend for the chart. |
| `line threeD group` | `chart group` | r/o | Returns a chart group object that represents the line chart group on a 3-D chart. |
| `name` | `text` | r/w | Returns or sets the name of the chart. |
| `next` | `chart` | r/o | Returns a worksheet object that represents the next sheet. |
| `page setup object` | `page setup` | r/o | Returns the page setup object associated with this chart. |
| `perspective` | `integer` | r/w | Returns or sets the perspective for the 3-D chart view. Must be between 0 and 100. This property is ignored if the right angle axes property is true. |
| `pie threeD group` | `chart group` | r/o | Returns a chart group object that represents the pie chart group on a 3-D chart. |
| `plot area object` | `plot area` | r/o | Returns a plot area object that represents the plot area of a chart. |
| `plot by` | `XlRowCol` | r/w | Returns or sets the way columns or rows are used as data series on the chart. |
| `plot visible only` | `boolean` | r/w | Returns or sets if only visible cells are plotted. False if both visible and hidden cells are plotted. |
| `previous` | `null` | r/o | Returns a worksheet object that represents the previous sheet. |
| `protect contents` | `boolean` | r/o | Returns true if the contents of the sheet are protected. |
| `protect data` | `boolean` | r/w | Returns or sets if series formulas cannot be modified by the user. |
| `protect drawing objects` | `boolean` | r/o | Returns true if shapes are protected. |
| `protect formatting` | `boolean` | r/w | Returns or sets if chart formatting cannot be modified by the user. |
| `protect goal seek` | `boolean` | r/w | Returns or sets if the user cannot modify chart data points with mouse actions. |
| `protect selection` | `boolean` | r/w | Returns or sets if chart elements cannot be selected. |
| `protection mode` | `boolean` | r/o | Returns true if user-interface-only protection is turned on. To turn on user interface protection, use the protect method with the user interface only argument set to true. |
| `right angle axes` | `boolean` | r/w | Returns or sets if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts. |
| `rotation` | `integer` | r/w | The rotation of the 3D chart view.  The value of must be from 0 to 360. |
| `sheet tab` | `tab` | r/o | Returns the sheet tab of the chart sheet |
| `show data labels over maximum` | `boolean` | r/w | Returns or sets whether to show the data labels when the value is greater than the maximum value on the value axis. |
| `show window` | `boolean` | r/w | Returns or sets if the embedded chart is displayed in a separate window. The Chart object used with this property must refer to an embedded chart. |
| `side wall` | `walls` | r/o | Returns a walls object that allows the user to individually format the side wall of a 3-D chart. |
| `size with window` | `boolean` | r/w | Returns or sets if Microsoft Excel resizes the chart to match the size of the chart sheet window. False if the chart size isn't attached to the window size. Applies only to chart sheets, doesn't apply to embedded charts. |
| `surface group` | `chart group` | r/o | Returns a chart group object that represents the surface chart group of a 3-D chart. |
| `visible` | `XlSheetVisibility` | r/w | Returns or sets if the chart is visible. |
| `walls and gridlines twoD` | `boolean` | r/w | Returns or sets if gridlines are drawn two-dimensionally on a 3-D chart. |
| `walls object` | `walls` | r/o | Returns a walls object that represents the walls of the 3-D chart. |

### `column group`

*Plural:* `column groups`

### `corners`

Represents the corners of a 3-D chart.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | Returns the name of this object. |

### `data label`

Represents the data label on a chart point or trendline.

*Plural:* `data labels`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets if the text in the object changes font size when the object size changes. |
| `auto text` | `boolean` | r/w | Returns or sets if the object automatically generates appropriate text based on context. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `caption` | `text` | r/w | Returns or sets the caption of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the data label. |
| `data label text` | `text` | r/w | Returns or sets the text for this object. |
| `data label type` | `XlDataLabelsType` | r/w | Returns or sets the type of the data label. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `formula` | `text` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `formula local` | `text` | r/w | Returns or sets the formula for the object, using A1-style references in the language of the user. |
| `formula r1c1` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. |
| `formula r1c1 local` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation in the language of the user. |
| `height` | `real` | r/o | Returns the height of the object. |
| `horizontal alignment` | `XlHAlign` | r/w | Returns or sets the horizontal alignment for the object. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the left position of the specified object, in points. |
| `name` | `text` | r/o | Returns the name of the object. |
| `number format` | `text` | r/w | Returns or sets the format code for the object. |
| `number format linked` | `boolean` | r/w | Returns or sets if the number format is linked to the cells, so that the number format changes in the labels when it changes in the cells. |
| `number format local` | `text` | r/w | Returns or sets the format code for the object as a string in the language of the user. |
| `orientation` | `integer` | r/w | The text orientation.  Can be a number value from -90 to 90 degrees. |
| `position` | `XlDataLabelPosition` | r/w | Returns or sets the position of the specified object. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `show legend key` | `boolean` | r/w | Returns or sets if the data label legend key is visible. |
| `top` | `real` | r/w | Returns or sets the top position of the specified object, in points. |
| `vertical alignment` | `XlVerticalAlignmentTarget` | r/w | Returns or sets the vertical alignment of the object. |
| `width` | `real` | r/o | Returns the width of the object. |

### `data table`

Represents a chart data table.

*Plural:* `data tables`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets if the text in the object changes font size when the object size changes. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the data table. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `has border horizontal` | `boolean` | r/w | Returns or sets if the chart data table has horizontal cell borders. |
| `has border outline` | `boolean` | r/w | Returns or sets if the chart data table has outline borders. |
| `has border vertical` | `boolean` | r/w | Returns or sets if the chart data table has vertical cell borders. |
| `show legend key` | `boolean` | r/w | Returns or sets if the data label legend key is visible. |

### `display unit label`

Represents a unit label on an axis in the specified chart. Unit labels are useful for charting large values-- for example, in the millions or billions.

*Plural:* `display unit labels`

**Contains:** `character`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets if the text in the object changes font size when the object size changes. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `caption` | `text` | r/w | Returns or sets the caption for the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the chart title. |
| `display label unit text` | `text` | r/w | Returns or sets the text for this object. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `horizontal alignment` | `XlHAlign` | r/w | Returns or sets the horizontal alignment for the object. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the left position of the specified object, in points. |
| `name` | `text` | r/o | Returns the name of the object. |
| `orientation` | `integer` | r/w | The text orientation.  Can be a number value from -90 to 90 degrees. |
| `position` | `XlChartElementPosition` | r/w | Returns or sets the position of the chart title on the chart. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `top` | `real` | r/w | Returns or sets the top position of the specified object, in points. |
| `vertical alignment` | `XlVerticalAlignmentTarget` | r/w | Returns or sets the vertical alignment of the object. |

### `doughnut group`

*Plural:* `doughnut groups`

### `down bars`

Represents the down bars in a chart group. Down bars connect points on the first series in the chart group with lower values on the last series, the lines go down from the first series.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the down bars. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `name` | `text` | r/o | Returns the name of the object. |

### `drop lines`

Represents the drop lines in a chart group. Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the drop lines. |
| `name` | `text` | r/o | Returns the name of this object. |

### `error bars`

Represents the error bars on a chart series. Error bars indicate the degree of uncertainty for chart data.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the error bars. |
| `end style` | `XlEndStyleCap` | r/w | Returns or sets the end style for the error bars. |
| `name` | `text` | r/o | Returns the name of the object. |

### `floor`

Represents the floor of a 3-D chart.

*Plural:* `floors`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the floor. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `name` | `text` | r/o | Returns the name of the object. |
| `picture type` | `XlChartPictureType` | r/w | Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart. |
| `thickness` | `integer` | r/w | Returns or sets  the thickness of the floor. |

### `gridlines`

Represents major or minor gridlines on a chart axis. Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the gridlines. |
| `name` | `text` | r/o | Returns the name of this object. |

### `hilo lines`

Represents the high-low lines in a chart group. High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2-D line groups can have high-low lines.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the hilo lines. |
| `name` | `text` | r/o | Returns the name of this object. |

### `interior`

Represents the interior of an object.

*Plural:* `interiors`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `integer` | r/w | Returns or sets the primary color of the object. |
| `color index` | `XlColorIndex` | r/w | Returns or sets the color of the interior. The color is specified as an index value into the current color palette. |
| `invert if negative` | `boolean` | r/w | Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. |
| `linear gradient` | `linear gradient` | r/o | Returns or sets the Gradient property of an Interior object of a selection. |
| `pattern` | `XlPattern` | r/w | Returns or sets the interior pattern. |
| `pattern color` | `integer` | r/w | Returns or sets the color of the interior pattern as an RGB value. |
| `pattern color index` | `XlColorIndex` | r/w | Returns or sets the color of the interior pattern as an index into the current color palette. |
| `pattern theme color` | `XlThemeColor` | r/w | Returns or sets a theme color pattern for an Interior object. |
| `pattern tint and shade` | `real` | r/w | Returns or sets a tint and shade pattern for an Interior object. |
| `rectangular gradient` | `rectangular gradient` | r/o | Returns or sets the Gradient property of an Interior object of a selection. |
| `theme color` | `XlThemeColor` | r/w | Returns or sets the theme color in the applied color scheme that is associated with the specified object. |
| `tint and shade` | `real` | r/w | Returns or sets a Single that lightens or darkens a color. |

### `leader lines`

Represents leader lines on a chart. Leader lines connect data labels to data points.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the leader lines. |

### `legend entry`

Represents a legend entry in a chart legend.

*Plural:* `legend entries`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets if the text in the object changes font size when the object size changes. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `height` | `real` | r/o | Returns the height of the object. |
| `left position` | `real` | r/o | Returns the left position of the specified object, in points. |
| `top` | `real` | r/o | Returns the top position of the specified object, in points. |
| `width` | `real` | r/o | Returns the width of the object. |

### `legend key`

Represents a legend key in a chart legend. Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart.

*Plural:* `legend keys`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the legend key. |
| `height` | `real` | r/o | Returns the height of the object. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `invert if negative` | `boolean` | r/w | Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. |
| `left position` | `real` | r/o | Returns the left position of the specified object, in points. |
| `marker background color` | `integer` | r/w | Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts. |
| `marker background color index` | `XlColorIndex` | r/w | Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts. |
| `marker foreground color` | `integer` | r/w | Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts. |
| `marker foreground color index` | `XlColorIndex` | r/w | Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts. |
| `marker size` | `integer` | r/w | Returns or sets the data-marker size, in points. |
| `marker style` | `XlMarkerStyle` | r/w | Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. |
| `picture type` | `XlChartPictureType` | r/w | Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart. |
| `picture unit` | `real` | r/w | Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `smooth` | `boolean` | r/w | Returns or sets if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts. |
| `top` | `real` | r/o | Returns the top position of the specified object, in points. |
| `width` | `real` | r/o | Returns the width of the object. |

### `legend`

Represents the legend in a chart. Each chart can have only one legend.

*Plural:* `legends`

**Contains:** `legend entry`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets if the text in the object changes font size when the object size changes. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the legend. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `include in layout` | `boolean` | r/w | Returns or sets if a legend will occupy the chart layout space when a chart layout is being determined. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the left position of the specified object, in points. |
| `name` | `text` | r/o | Returns the name of the object. |
| `position` | `XlLegendPosition` | r/w | Returns or sets the position of the legend on the chart. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `top` | `real` | r/w | Returns or sets the top position of the specified object, in points. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |

### `line group`

*Plural:* `line groups`

### `pie group`

*Plural:* `pie groups`

### `plot area`

Represents the plot area of a chart. This is the area where your chart data is plotted.

*Plural:* `plot areas`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the plot area. |
| `height` | `real` | r/w | Returns or sets the height of the object. |
| `inside height` | `real` | r/o | Returns the inside height of the plot area, in points. |
| `inside left` | `real` | r/o | Returns the distance from the chart edge to the inside left edge of the plot area, in points. |
| `inside top` | `real` | r/o | Returns the distance from the chart edge to the inside top edge of the plot area, in points. |
| `inside width` | `real` | r/o | Returns the inside width of the plot area, in points. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `left position` | `real` | r/w | Returns or sets the left position of the specified object, in points. |
| `name` | `text` | r/o | Returns the name of the object. |
| `position` | `XlChartElementPosition` | r/w | Returns or sets the position of the plot area on the chart. |
| `top` | `real` | r/w | Returns or sets the top position of the specified object, in points. |
| `width` | `real` | r/w | Returns or sets  the width of the object. |

### `radar group`

*Plural:* `radar groups`

### `series lines`

Represents series lines in a chart group. Series lines connect the data values from each series. Only 2-D stacked bar or column chart groups can have series lines.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the series lines. |
| `name` | `text` | r/o | Returns the name of this object. |

### `series point`

Represents a single point in a series in a chart.

*Plural:* `series points`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `apply pict to end` | `boolean` | r/w | Returns or sets if a picture is applied to the end of the point or all points in the series. |
| `apply pict to front` | `boolean` | r/w | Returns or sets if a picture is applied to the front of the point or all points in the series. |
| `apply pict to sides` | `boolean` | r/w | Returns or sets if a picture is applied to the sides of the point or all points in the series. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the point. |
| `data label object` | `data label` | r/o | Returns a data label object that represents the data label associated with the point or trendline. |
| `explosion` | `integer` | r/w | Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns zero if there's no explosion, the tip of the slice is in the center of the pie. |
| `has data label` | `boolean` | r/w | Returns or sets if the point has a data label. |
| `has threeD effect` | `boolean` | r/w | Returns or sets if a point has a three-dimensional appearance. Applies only to bubble charts. |
| `height` | `real` | r/o | Returns the height of the object. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `invert if negative` | `boolean` | r/w | Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. |
| `left position` | `real` | r/o | Returns the left position of the specified object, in points. |
| `marker background color` | `integer` | r/w | Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts. |
| `marker background color index` | `XlColorIndex` | r/w | Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts. |
| `marker foreground color` | `integer` | r/w | Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts. |
| `marker foreground color index` | `XlColorIndex` | r/w | Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts. |
| `marker size` | `integer` | r/w | Returns or sets the data-marker size, in points. |
| `marker style` | `XlMarkerStyle` | r/w | Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. |
| `name` | `text` | r/o | Returns the name of the object. |
| `picture type` | `XlChartPictureType` | r/w | Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart. |
| `picture unit` | `real` | r/w | Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored. |
| `secondary plot` | `boolean` | r/w | Returns or sets if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `top` | `real` | r/o | Returns the top position of the specified object, in points. |
| `width` | `real` | r/o | Returns the width of the object. |

### `series`

Represents a series in a chart.

*Plural:* `series collection`

**Contains:** `data label`, `series point`, `trendline`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `apply picture to end` | `boolean` | r/w | Returns or sets if a picture is applied to the end of the point or all points in the series. |
| `apply picture to front` | `boolean` | r/w | Returns or sets if a picture is applied to the front of the point or all points in the series. |
| `apply picture to sides` | `boolean` | r/w | Returns or sets if a picture is applied to the sides of the point or all points in the series. |
| `axis group` | `XlAxisGroup` | r/w | Returns or sets the axis group for the specified series. |
| `bar shape` | `XlBarShape` | r/w | Returns or sets the shape used with the 3-D bar or column chart. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `bubble sizes` | `text` | r/w | Returns or sets a string in A1-style notation that refers to the worksheet cells containing the size data for the bubble chart. Applies only to bubble charts. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the series. |
| `chart type` | `XlChartType` | r/w | Returns or sets the chart type. |
| `error bars` | `error bars` | r/o | Returns an error bars object that represents the error bars for the series. |
| `explosion` | `integer` | r/w | Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns zero if there's no explosion, the tip of the slice is in the center of the pie. |
| `formula` | `text` | r/w | Returns or sets the object's formula, in A1-style notation and in the language of the macro. |
| `formula local` | `text` | r/w | Returns or sets the formula for the object, using A1-style references in the language of the user. |
| `formula r1c1` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. |
| `formula r1c1 local` | `text` | r/w | Returns or sets the formula for the object, using R1C1-style notation in the language of the user. |
| `has data labels` | `boolean` | r/w | Returns or sets if the series has data labels. |
| `has error bars` | `boolean` | r/w | Returns or set if the series has error bars. This property isn't available for 3-D charts. |
| `has leader lines` | `boolean` | r/w | Returns or sets if the series has leader lines. |
| `has threeD effect` | `boolean` | r/w | Returns or sets if the series has a three-dimensional appearance. Applies only to bubble charts. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `invert if negative` | `boolean` | r/w | Returns or sets if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. |
| `leader lines` | `leader lines` | r/o | Returns a leader lines object that represents the leader lines for the series. |
| `marker background color` | `integer` | r/w | Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts. |
| `marker background color index` | `XlColorIndex` | r/w | Returns or sets the marker background color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts. |
| `marker foreground color` | `integer` | r/w | Returns or sets the foreground color of the marker as an RGB value. Applies only to line, scatter, and radar charts. |
| `marker foreground color index` | `XlColorIndex` | r/w | Returns or sets the marker foreground color as an index into the current color palette, or as one of the following constants: color index automatic or color index none. Applies only to line, scatter, and radar charts. |
| `marker size` | `integer` | r/w | Returns or sets the data-marker size, in points. |
| `marker style` | `XlMarkerStyle` | r/w | Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. |
| `name` | `text` | r/w | Returns or sets the name of the object. |
| `picture type` | `XlChartPictureType` | r/w | Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart. |
| `picture unit` | `real` | r/w | Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored. |
| `plot color index` | `integer` | r/o | Returns the plot color index of the series. |
| `plot order` | `integer` | r/w | Returns or sets the plot order for the selected series within the chart group. |
| `series values` | `XlCategoryNames` | r/w | Returns or sets a list of all the values in the series. This can be a range on a worksheet or a list of constant values. |
| `shadow` | `boolean` | r/w | Returns or sets if the object has a shadow. |
| `smooth` | `boolean` | r/w | Returns or sets if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts. |
| `xvalues` | `XlCategoryNames` | r/w | Returns or sets a list of x values for a chart series. The xvalues property can be set to a range on a worksheet or to a list of values. |

### `tick labels`

Represents the tick-mark labels associated with tick marks on a chart axis.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto scale font` | `boolean` | r/w | Returns or sets if the text in the object changes font size when the object size changes. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the tick labels. |
| `depth` | `integer` | r/o | Returns the number of levels of category tick labels. |
| `font object` | `font` | r/o | Returns a font object that represents the font of the specified object. |
| `multi level` | `boolean` | r/w | Returns or sets whether an axis is multilevel or not. |
| `name` | `text` | r/o | Returns the name of the object. |
| `number format` | `text` | r/w | Returns or sets the format code for the object. |
| `number format linked` | `boolean` | r/w | Returns or sets if the number format is linked to the cells, so that the number format changes in the labels when it changes in the cells. |
| `number format local` | `text` | r/w | Returns or sets the format code for the object as a string in the language of the user. |
| `offset` | `integer` | r/w | Returns or sets the distance between the levels of labels, and the distance between the first level and the axis line. The value can be an integer percentage from 0 through 1000, relative to the axis label's font size. |
| `orientation` | `XlTickLabelOrientation` | r/w | The text orientation.  Can be a number value from -90 to 90 degrees. |
| `reading order` | `XLDefaultSheetDir` | r/w | Returns or sets the reading order for the specified object. |
| `tick alignment` | `XlTickHAlign` | r/w | Returns or sets the alignment for the specified tick label. |

### `trendline`

Represents a trendline in a chart. A trendline shows the trend, or direction, of data in a series.

*Plural:* `trendlines`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `backward` | `real` | r/w | Returns or sets the number of periods, or units on a scatter chart, that the trendline extends backward. |
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the trendline. |
| `data label object` | `data label` | r/o | Returns a data label object that represents the data label associated with the point or trendline. |
| `display R squared` | `boolean` | r/w | Returns or sets if the R-squared value of the trendline is displayed on the chart, in the same data label as the equation. Setting this property to true automatically turns on data labels. |
| `display equation` | `boolean` | r/w | Returns or sets if the equation for the trendline is displayed on the chart, in the same data label as the R-squared value. Setting this property to true automatically turns on data labels. |
| `entry_index` | `integer` | r/o | Returns the index number of the object within the elements of the parent object. |
| `forward` | `real` | r/w | Returns or sets the number of periods, or units on a scatter chart, that the trendline extends forward. |
| `intercept` | `real` | r/w | Returns or sets the point where the trendline crosses the value axis. |
| `intercept is auto` | `boolean` | r/w | Returns or sets if the point where the trendline crosses the value axis is automatically determined by the regression. |
| `name` | `text` | r/w | Returns or sets the name of this object. |
| `name is auto` | `boolean` | r/w | Returns or sets if Microsoft Excel automatically determines the name of the trendline. |
| `order` | `integer` | r/w | Returns or sets the trendline order, an integer greater than 1,  when the trendline type is polynomial. |
| `period` | `integer` | r/w | Returns or sets the period for the moving-average trendline. |
| `trendline type` | `XlTrendlineType` | r/w | Returns or sets the type of this trend line |

### `up bars`

Represents the up bars in a chart group. Up bars connect points on series one with higher values on the last series in the chart group, the lines go up from series one. Only 2-D line groups that contain at least two series can have up bars.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the up bars. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `name` | `text` | r/o | Returns the name of this object. |

### `walls`

Represents the walls of a 3-D chart.

*Plural:* `walls collection`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `border` | `border` | r/o | Returns a border object that represents the border of the object. |
| `chart fill format object` | `chart fill format` | r/o | Returns a chart fill format object for this object. |
| `chart format` | `chart format` | r/o | Returns a chart format object that contains chart formatting properties for the walls. |
| `interior object` | `interior` | r/o | Returns an interior object that represents the interior of the specified object. |
| `name` | `text` | r/o | Returns or sets the name of the object. |
| `picture type` | `XlChartPictureType` | r/w | Returns or sets the way pictures are displayed on a column or bar picture chart or on the walls and faces of a 3-D chart. |
| `picture unit` | `integer` | r/w | Returns or sets the unit for each picture on the chart if the picture type property is set to chart picture type stack scale, if not, this property is ignored. |
| `thickness` | `integer` | r/w | Returns or sets  the thickness of the floor. |

### `xy group`

*Plural:* `xy groups`

### Enumerations

### `chart patterned`

| Value | Description |
|-------|-------------|
| `chart fill format` |  |
| `chart title` |  |
| `axis title` |  |
| `series point` |  |
| `series` |  |
| `data label` |  |
| `legend key` |  |
| `down bars` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `legend` |  |
| `display unit label` |  |

### `clear`

| Value | Description |
|-------|-------------|
| `chart area` |  |
| `legend` |  |

### `clear formats`

| Value | Description |
|-------|-------------|
| `series point` |  |
| `series` |  |
| `legend key` |  |
| `trendline` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `error bars` |  |

### `apply data labels`

| Value | Description |
|-------|-------------|
| `chart` |  |
| `series point` |  |
| `series` |  |

### `paste`

| Value | Description |
|-------|-------------|
| `floor` |  |
| `walls` |  |

### `chart one color gradient`

| Value | Description |
|-------|-------------|
| `chart fill format` |  |
| `chart title` |  |
| `axis title` |  |
| `series point` |  |
| `series` |  |
| `data label` |  |
| `legend key` |  |
| `down bars` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `legend` |  |
| `display unit label` |  |

### `chart two color gradient`

| Value | Description |
|-------|-------------|
| `chart fill format` |  |
| `chart title` |  |
| `axis title` |  |
| `series point` |  |
| `series` |  |
| `data label` |  |
| `legend key` |  |
| `down bars` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `legend` |  |
| `display unit label` |  |

### `chart solid`

| Value | Description |
|-------|-------------|
| `chart fill format` |  |
| `chart title` |  |
| `axis title` |  |
| `series point` |  |
| `series` |  |
| `data label` |  |
| `legend key` |  |
| `down bars` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `legend` |  |
| `display unit label` |  |

### `copy object`

| Value | Description |
|-------|-------------|
| `chart object` |  |
| `series point` |  |
| `series` |  |
| `chart area` |  |

### `chart user picture`

| Value | Description |
|-------|-------------|
| `chart fill format` |  |
| `chart title` |  |
| `axis title` |  |
| `series point` |  |
| `series` |  |
| `data label` |  |
| `legend key` |  |
| `down bars` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `legend` |  |
| `display unit label` |  |

### `delete object`

| Value | Description |
|-------|-------------|
| `chart object` |  |
| `axis` |  |

### `preset chart gradient`

| Value | Description |
|-------|-------------|
| `chart fill format` |  |
| `chart title` |  |
| `axis title` |  |
| `series point` |  |
| `series` |  |
| `data label` |  |
| `legend key` |  |
| `down bars` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `legend` |  |
| `display unit label` |  |

### `preset chart textured`

| Value | Description |
|-------|-------------|
| `chart fill format` |  |
| `chart title` |  |
| `axis title` |  |
| `series point` |  |
| `series` |  |
| `data label` |  |
| `legend key` |  |
| `down bars` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `legend` |  |
| `display unit label` |  |

### `chart user textured`

| Value | Description |
|-------|-------------|
| `chart fill format` |  |
| `chart title` |  |
| `axis title` |  |
| `series point` |  |
| `series` |  |
| `data label` |  |
| `legend key` |  |
| `down bars` |  |
| `floor` |  |
| `walls` |  |
| `plot area` |  |
| `chart area` |  |
| `legend` |  |
| `display unit label` |  |

### `activate object`

| Value | Description |
|-------|-------------|
| `chart` |  |
| `chart object` |  |
