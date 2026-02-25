# Powerpoint AppleScript Dictionary

> Auto-generated from `PowerPoint.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Powerpoint"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Microsoft Office Suite](#microsoft-office-suite)
- [Microsoft PowerPoint Suite](#microsoft-powerpoint-suite)
- [Drawing Suite](#drawing-suite)
- [Text Suite](#text-suite)
- [Table Suite](#table-suite)

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

## Microsoft PowerPoint Suite

PowerPoint specific classes and commands

### Commands

### `add behavior`

add an animation behavior

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `effect` | no |  |
| `type` | `MsoAnimType` | no |  |

**Returns:** `animation behavior`

### `add design`

add a new design

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |
| `DesignName` | `text` | no |  |
| `index` | `integer` | yes |  |

**Returns:** `design`

### `add effect`

add an effect for a shape

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sequence` | no |  |
| `for` | `shape` | no |  |
| `fx` | `MsoAnimEffect` | no |  |
| `level` | `MsoAnimateByLevel` | yes |  |
| `trigger` | `MsoAnimTriggerType` | yes |  |
| `index` | `integer` | yes |  |

**Returns:** `effect`

### `add sequence`

add an interactive sequence

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `timeline` | no |  |
| `index` | `integer` | yes |  |

**Returns:** `sequence`

### `apply template`

Applies a theme or design template to the specified slide, master or presentation

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |
| `file name` | `text` | no |  |

### `apply theme`

Applies a theme or design template to the specified slide, master or presentation

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4002` | no |  |
| `file name` | `text` | no |  |

### `apply theme color scheme`

Applies a theme color to the specified slide.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide` | no |  |
| `file name` | `text` | no |  |

### `arrange windows`

Arrange Document Windows

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `EPPArrangeStyle` | no |  |

### `can check in`

Returns True if PowerPoint can check in a specified presentation to a server.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |

**Returns:** `boolean`

### `can check out`

Returns True if PowerPoint can check out a specified presentation from a server.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `file name` | `text` | no | The server path and name of the presentation. |

**Returns:** `boolean`

### `check in`

Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |
| `save changes` | `boolean` | yes | True saves the presentation to the server location. The default value is False. |
| `comments` | `text` | yes | Comments for the revision of the presentation being checked in. Only applies if SaveChanges equals True. |
| `make public` | `boolean` | yes | True allows the user to perform a publish on the presentation after being checked in. |

### `check in with version`

Returns a presentation from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |
| `save changes` | `boolean` | yes | True saves the presentation to the server location. The default value is False. |
| `comments` | `text` | yes | Comments for the revision of the presentation being checked in. Only applies if SaveChanges equals True. |
| `make public` | `boolean` | yes | True allows the user to perform a publish on the presentation after being checked in. |
| `version type` | `EPPCheckInVersionType` | yes | Version number of the presentation. |

### `check out`

Copies a specified presentation from a server to a local computer for editing. Returns a String that represents the local path and file name of the presentation checked out.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |
| `file name` | `text` | no | The server path and name of the presentation. |

### `collapse section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document window` | no |  |
| `at position` | `integer` | no |  |

### `convert to text unit effect`

convert an effect to a text unit effect

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sequence` | no |  |
| `Effect` | `effect` | no |  |
| `unit` | `MsoAnimTextUnitEffect` | no |  |

**Returns:** `effect`

### `copy object`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide` | no |  |

### `cut object`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide` | no |  |

### `delete section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |
| `at position` | `integer` | yes |  |
| `deleting slides` | `boolean` | no |  |

### `end show`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presenter tool` | no |  |

### `exit slide show`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide show view` | no |  |

### `expand section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document window` | no |  |
| `at position` | `integer` | no |  |

### `get color from`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `color scheme` | no |  |
| `at` | `EPPColorSchemeIndex` | no |  |

**Returns:** `integer`

### `get count of sections`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |

**Returns:** `integer`

### `get first slide of section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |
| `at position` | `integer` | yes |  |

**Returns:** `integer`

### `get id of section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |
| `at position` | `integer` | no |  |

**Returns:** `text`

### `get is expanded of section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document window` | no |  |
| `at position` | `integer` | no |  |

**Returns:** `boolean`

### `get name of section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |
| `at position` | `integer` | yes |  |

**Returns:** `text`

### `get player from`

get a player from a shape inside a slide show view

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4001` | no |  |
| `for` | `shape` | no |  |

**Returns:** `player`

### `get slide count of section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |
| `at position` | `integer` | yes |  |

**Returns:** `integer`

### `get text style from`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `master` | no |  |
| `at` | `PpTextStyleType` | no |  |

**Returns:** `text style`

### `go to first slide`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide show view` | no |  |

### `go to last slide`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide show view` | no |  |

### `go to next bookmark for player`

Advance the player to the next bookmark.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `player` | no |  |

### `go to next slide`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide show view` | no |  |

### `go to previous bookmark for player`

Rewind the player to the previous bookmark.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `player` | no |  |

### `go to previous slide`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide show view` | no |  |

### `go to slide`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `view` | no |  |
| `number` | `integer` | no |  |

### `import sound file`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sound effect` | no |  |
| `sound file name` | `text` | no |  |

### `insert`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `the text` | `text` | no |  |
| `at` | `location specifier` | no |  |

### `insert section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |
| `before section` | `integer` | yes |  |
| `before slide` | `integer` | yes |  |
| `titled` | `text` | no |  |

**Returns:** `integer`

### `launch speller on`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `document window` | no |  |

### `move section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |
| `at position` | `integer` | no |  |
| `to position` | `integer` | no |  |

### `move to start of section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide` | no |  |
| `at position` | `integer` | no |  |

### `next`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presenter tool` | no |  |

### `paste object`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4003` | no |  |

### `pause player`

Pause playback.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `player` | no |  |

### `pause timer`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presenter tool` | no |  |

### `play player`

Begin or resume playback.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `player` | no |  |

### `play sound effect`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `sound effect` | no |  |

### `previous`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presenter tool` | no |  |

### `print out`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |
| `from` | `integer` | yes |  |
| `to` | `integer` | yes |  |
| `print to file` | `text` | yes |  |
| `copies` | `integer` | yes |  |
| `collate` | `boolean` | yes |  |
| `showDialog` | `boolean` | yes |  |

### `quit`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no |  |

### `redo`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |
| `times` | `integer` | yes | The number of actions to be redone. |

### `register add in`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no |  |

**Returns:** `add in`

### `rename section`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `section properties` | no |  |
| `at position` | `integer` | no |  |
| `to` | `text` | no |  |

### `reset slide time`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide show view` | no |  |

### `reset timer`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presenter tool` | no |  |

### `run VB macro`

Runs a Visual Basic macro.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `macro name` | `text` | no |  |
| `list of parameters` | `text` | no |  |

**Returns:** `integer`

### `run slide show`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `slide show settings` | no |  |

**Returns:** `slide show window`

### `set bullet picture`

Sets the graphics file to be used for bullets in a bulleted list.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `bullet format` | no |  |
| `picture file` | `text` | no |  |

### `set color for`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `color scheme` | no |  |
| `at` | `EPPColorSchemeIndex` | no |  |
| `to color` | `integer` | no |  |

### `start timer`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presenter tool` | no |  |

### `stop player`

Stop playback.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `player` | no |  |

### `swap displays`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presenter tool` | no |  |

### `toggle slide miniature`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presenter tool` | no |  |

### `undo`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |
| `times` | `integer` | yes | The number of actions to be undone. |

### `unselect`

Cancels the current selection.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `selection` | no |  |

### `update links`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `presentation` | no |  |

### Classes

### `action setting`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `action` | `EPPActionType` | r/w |  |
| `action setting to run` | `text` | r/w |  |
| `action sound effect` | `sound effect` | r/o |  |
| `action verb` | `text` | r/w |  |
| `animate action` | `boolean` | r/w |  |
| `hyperlink` | `hyperlink` | r/o |  |
| `slide show name` | `text` | r/w |  |

### `add in`

*Plural:* `add ins`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto load` | `boolean` | r/w |  |
| `full name` | `text` | r/o |  |
| `loaded` | `boolean` | r/w |  |
| `name` | `text` | r/o |  |
| `path` | `text` | r/o |  |
| `registered` | `boolean` | r/w |  |
| `registered in HKLM` | `boolean` | r/o |  |

### `animation behavior`

*Plural:* `animation behaviors`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accumulate` | `MsoAnimAccumulate` | r/w |  |
| `additive` | `MsoAnimAdditive` | r/w |  |
| `animation behavior type` | `MsoAnimType` | r/w |  |
| `colors effect` | `colors effect` | r/o |  |
| `command effect` | `command effect` | r/o |  |
| `filter effect` | `filter effect` | r/o |  |
| `motion effect` | `motion effect` | r/o |  |
| `property effect` | `property effect` | r/o |  |
| `rotating effect` | `rotating effect` | r/o |  |
| `scale effect` | `scale effect` | r/o |  |
| `set effect` | `set effect` | r/o |  |
| `timing` | `timing` | r/o |  |

### `animation point`

*Plural:* `animation points`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `formula` | `text` | r/w |  |
| `time` | `real` | r/w |  |
| `value` | `location specifier` | r/w |  |

### `animation settings`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `advance time` | `real` | r/w |  |
| `after effect` | `EPPAfterEffect` | r/w |  |
| `animate` | `boolean` | r/w |  |
| `animate background` | `boolean` | r/w |  |
| `animate text in reverse` | `boolean` | r/w |  |
| `animation order` | `integer` | r/w |  |
| `animation play settings` | `play settings` | r/o |  |
| `animation sound effect` | `sound effect` | r/o |  |
| `chart unit effect` | `EPPChartUnitEffect` | r/w |  |
| `dim color` | `integer` | r/w |  |
| `dim color theme index` | `MsoThemeColorIndex` | r/w |  |
| `entry effect` | `EPPEntryEffect` | r/w |  |
| `text level effect` | `EPPTextLevelEffect` | r/w |  |
| `text unit effect` | `EPPTextUnitEffect` | r/w |  |

### `application`

An AppleScript object representing the Microsoft POWERPOINT application.

**Contains:** `presentation`, `document window`, `slide show window`, `command bar`, `add in`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `Version` | `text` | r/o |  |
| `active presentation` | `presentation` | r/o |  |
| `active printer` | `text` | r/o |  |
| `active window` | `document window` | r/o |  |
| `autocorrect object` | `autocorrect` | r/o | Returns the autocorrect object |
| `automation security` | `MsoAutomationSecurity` | r/w |  |
| `build` | `text` | r/o |  |
| `caption` | `text` | r/o |  |
| `default save format` | `EPPSaveAsFileType` | r/w |  |
| `default web options object` | `default web options` | r/o |  |
| `name` | `text` | r/o |  |
| `path` | `text` | r/o |  |
| `preference object` | `preferences` | r/o |  |
| `save as movie settings object` | `save as movie settings` | r/o |  |
| `start up dialog` | `boolean` | r/w |  |

### `autocorrect entry`

Represents a single autocorrect entry.

*Plural:* `autocorrect entries`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `autocorrect value` | `text` | r/o | Returns the value of the auto correct entry. |
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | Returns the name of the auto correct entry. |

### `autocorrect`

Represents the autocorrect functionality in PowerPoint.

**Contains:** `autocorrect entry`, `first letter exception`, `two initial caps exception`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `correct days` | `boolean` | r/w | Returns if PowerPoint automatically capitalizes the first letter of days of the week. |
| `correct initial caps` | `boolean` | r/w | Returns if PowerPoint automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, POwerPoint is corrected to PowerPoint. |
| `correct sentence caps` | `boolean` | r/w | Returns if PowerPoint automatically capitalizes the first letter in each sentence. |
| `replace text` | `boolean` | r/w | Returns if Microsoft PowerPoint automatically replaces specified text with entries from the autocorrect list. |

### `bullet format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bullet character` | `text` | r/w | Returns or sets the Unicode character value that is used for bullets in the specified text. |
| `bullet font` | `font` | r/o | Returns a font object that represents character formatting for a bullet format object. |
| `bullet number` | `integer` | r/o | Returns the bullet number of a paragraph. |
| `bullet start value` | `integer` | r/w |  |
| `bullet style` | `MsoNumberedBulletStyle` | r/w | Returns or sets a constant that represents the style of a bullet. |
| `bullet type` | `MsoBulletType` | r/w | Returns or sets a constant that represents the type of bullet. |
| `relative size` | `real` | r/w | Returns or sets the bullet size relative to the size of the first text character in the paragraph. |
| `use text color` | `boolean` | r/w | Determines whether the specified bullets are set to the color of the first text character in the paragraph. |
| `use text font` | `boolean` | r/w | Determines whether the specified bullets are set to the font of the first text character in the paragraph. |
| `visible` | `boolean` | r/w | Returns or sets a value that specifies whether the bullet is visible. |

### `color scheme`

*Plural:* `color schemes`

### `colors effect`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `by_color` | `integer` | r/w |  |
| `by_color theme index` | `MsoThemeColorIndex` | r/w | Returns an object that represents a change to the color of the object by the specified number, expressed in RGB format. |
| `from_color` | `integer` | r/w |  |
| `from_color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets an object that represents the starting RGB color value of an animation behavior. |
| `to_color` | `integer` | r/w |  |
| `to_color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets an object that represents the RGB color value of an animation behavior. |

### `command effect`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `command` | `text` | r/w |  |
| `type` | `MsoAnimCommandType` | r/w |  |

### `custom layout`

*Plural:* `custom layouts`

**Contains:** `shape`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `background` | `shape` | r/o |  |
| `design` | `design` | r/o |  |
| `display master shapes` | `boolean` | r/w |  |
| `follow master background` | `boolean` | r/w |  |
| `headers and footers` | `headers and footers` | r/o |  |
| `height` | `real` | r/o |  |
| `theme color scheme` | `theme color scheme` | r/o | Returns the color scheme of a custom layout. |
| `timeline` | `timeline` | r/o |  |
| `width` | `real` | r/o |  |

### `default web options`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow PNG` | `boolean` | r/w |  |
| `always save in default encoding` | `boolean` | r/w |  |
| `check if Office is HTML editor` | `boolean` | r/w |  |
| `encoding` | `MsoEncoding` | r/w |  |
| `update links on save` | `boolean` | r/w |  |

### `design`

*Plural:* `designs`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `slide master` | `master` | r/o |  |

### `document window`

*Plural:* `document windows`

**Contains:** `pane`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active` | `boolean` | r/o |  |
| `active pane` | `pane` | r/o |  |
| `black and white` | `boolean` | r/w |  |
| `caption` | `text` | r/o |  |
| `entry_index` | `integer` | r/o |  |
| `height` | `real` | r/w |  |
| `left position` | `real` | r/w |  |
| `presentation` | `presentation` | r/o |  |
| `selection` | `selection` | r/o | Represents the selection in the specified document window. |
| `split horizontal` | `integer` | r/w |  |
| `split vertical` | `integer` | r/w |  |
| `top` | `real` | r/w |  |
| `view` | `view` | r/o |  |
| `view type` | `EPPViewType` | r/w |  |
| `width` | `real` | r/w |  |

### `effect information`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `after effect information` | `MsoAnimAfterEffect` | r/o |  |
| `animate background information` | `boolean` | r/o |  |
| `animate text in reverse information` | `boolean` | r/o |  |
| `build by level` | `MsoAnimateByLevel` | r/o |  |
| `dim` | `integer` | r/o |  |
| `play settings information` | `play settings` | r/o |  |
| `sound effect information` | `sound effect` | r/o |  |
| `text unit effect information` | `MsoAnimTextUnitEffect` | r/o |  |

### `effect parameters`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `amount` | `real` | r/w |  |
| `color2` | `integer` | r/o |  |
| `color2 color theme index` | `MsoThemeColorIndex` | r/o | Returns an object that represents the color on which to end a color-cycle animation. |
| `direction` | `MsoAnimDirection` | r/w |  |
| `font2` | `text` | r/w |  |
| `relative` | `boolean` | r/w |  |
| `size2` | `real` | r/w |  |

### `effect`

*Plural:* `effects`

**Contains:** `animation behavior`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `animation effect type` | `MsoAnimEffect` | r/w |  |
| `effect information` | `effect information` | r/o |  |
| `effect parameters` | `effect parameters` | r/o |  |
| `exit animation` | `boolean` | r/w |  |
| `name` | `text` | r/o |  |
| `shape` | `shape` | r/w |  |
| `target paragraph` | `integer` | r/w |  |
| `text range length` | `integer` | r/o |  |
| `text range start` | `integer` | r/o |  |
| `timing` | `timing` | r/o |  |

### `filter effect`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `filter type` | `MsoAnimFilterEffectType` | r/w |  |
| `reveal` | `boolean` | r/w |  |
| `subtype` | `MsoAnimFilterEffectSubtype` | r/w |  |

### `first letter exception`

Represents an abbreviation excluded from automatic correction.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | Returns the name of the FirstLetterException. |

### `font`

Contains font attributes, such as font name, size, and color, for an object.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `ASCII name` | `text` | r/w | Returns or sets the font used for Latin text; characters with character codes from 0 through 127. |
| `auto rotate numbers` | `boolean` | r/w | Returns or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated. |
| `base line offset` | `real` | r/w | Returns or sets a value specifying the horizontaol offset of the selected font. |
| `bold` | `boolean` | r/w | Returns or sets a value specifying whether the font should be bold. |
| `caps type` | `MsoTextCaps` | r/w | Returns or sets a value specifying how the text should be capitalized. |
| `east asian name` | `text` | r/w | Returns or sets the font name used for Asian text. |
| `embedable` | `boolean` | r/o | Returns a value indicating whether the font can be embedded in a page. |
| `embedded` | `boolean` | r/o | Returns a value specifying whether the font is embedded in a page. |
| `emboss` | `boolean` | r/w |  |
| `equalize character height` | `boolean` | r/w | Returns or sets a value specifying whether the text should have the same horizontal height. |
| `fill format` | `fill format` | r/o | Returns a value specifying the fill formatting for text. |
| `font color` | `integer` | r/w |  |
| `font color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color for the specified font. |
| `font name` | `text` | r/w | Returns or sets a value specifying the font to use for a selection. |
| `font name other` | `text` | r/w | Returns or sets the font used for characters whose character set numbers are greater than 127. |
| `font size` | `real` | r/w |  |
| `glow format` | `glow format` | r/o | Returns a value specifying the glow formatting of the text. |
| `highlight color` | `integer` | r/w | Returns or sets the text highlight color for object. |
| `highlight color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the specified text highlight color for object. |
| `italic` | `boolean` | r/w |  |
| `kerning` | `real` | r/w | Returns or sets a value specifying the amount of spacing between text characters. |
| `line format` | `line format` | r/o | Returns a value specifiying the line formatting of the text. |
| `reflection format` | `reflection format` | r/o | Returns a value specifying the reflection formatting of the text. |
| `shadow format` | `shadow format` | r/o | Returns the value specifying the type of shadow effect for the selection of text. |
| `soft edge format` | `MsoSoftEdgeType` | r/w | Returns or sets a value specifying the soft edge fromatting of the text. |
| `spacing` | `real` | r/w | Returns or sets a value specifying the spacing between characters in a selection of text. |
| `strike type` | `MsoTextStrike` | r/w | Returns or sets a value specifying the strike format used for a selection of text. |
| `subscript` | `boolean` | r/w | Returns or sets a value specifying that the selected text should be displayed a subscript. |
| `superscript` | `boolean` | r/w | Returns or sets a value specifying that the selected text should be displayed a superscript. |
| `transparency` | `real` | r/w |  |
| `underline` | `boolean` | r/w |  |
| `underline color` | `integer` | r/w | Returns or sets the 24-bit color of the underline for the specified font object. |
| `underline color theme index` | `MsoThemeColorIndex` | r/w | Returns a value specifying the color of the underline for the selected text. |
| `underline style` | `MsoTextUnderlineType` | r/w | Returns or sets a value specifying the underline style for the selected text. |
| `word art styles format` | `MsoPresetTextEffect` | r/w | Returns or sets a value specifying the text effect for the selected text. |

### `header or footer`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `date format` | `EPPDateTimeFormat` | r/w |  |
| `header footer text` | `text` | r/w |  |
| `use date format` | `boolean` | r/w |  |
| `visible` | `boolean` | r/w |  |

### `headers and footers`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `date and time` | `header or footer` | r/o |  |
| `display headers footers on title slide` | `boolean` | r/w |  |
| `footer` | `header or footer` | r/o |  |
| `header` | `header or footer` | r/o |  |
| `slide number` | `header or footer` | r/o |  |

### `hyperlink`

*Plural:* `hyperlinks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `hyperlink address` | `text` | r/w |  |
| `hyperlink sub address` | `text` | r/w |  |
| `hyperlink type` | `mHyT` | r/o |  |

### `master`

**Contains:** `shape`, `hyperlink`, `custom layout`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `background` | `shape` | r/o |  |
| `color scheme` | `color scheme` | r/o |  |
| `design` | `design` | r/o |  |
| `headers and footers` | `headers and footers` | r/o |  |
| `height` | `real` | r/o |  |
| `theme` | `office theme` | r/o |  |
| `timeline` | `timeline` | r/o |  |
| `width` | `real` | r/o |  |

### `motion effect`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `by x` | `real` | r/w |  |
| `by y` | `real` | r/w |  |
| `from x` | `real` | r/w |  |
| `from y` | `real` | r/w |  |
| `path` | `text` | r/w |  |
| `to x` | `real` | r/w |  |
| `to y` | `real` | r/w |  |

### `named slide show`

*Plural:* `named slide shows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o |  |
| `number of slides` | `integer` | r/o |  |
| `slide IDs` | `integer` | r/o |  |

### `page setup`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `first slide number` | `integer` | r/w |  |
| `notes orientation` | `MsoOrientation` | r/w |  |
| `slide orientation` | `MsoOrientation` | r/w |  |
| `slide size` | `EPPSlideSizeType` | r/w |  |
| `slide width` | `real` | r/w |  |

### `pane`

*Plural:* `panes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active` | `boolean` | r/o |  |
| `pane view type` | `EPPViewType` | r/o |  |

### `paragraph format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `alignment` | `MsoParagraphAlignment` | r/w |  |
| `baseline alignment` | `MsoBaselineAlignment` | r/w | Returns or sets a constant that represents the vertical position of fonts in a paragraph. |
| `bullet format` | `bullet format` | r/o |  |
| `east asian line break control` | `boolean` | r/w |  |
| `hanging punctuation` | `boolean` | r/w | Determines whether hanging punctuation is enabled for the specified paragraphs. |
| `line rule after` | `boolean` | r/w | Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines. |
| `line rule before` | `boolean` | r/w | Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines. |
| `line rule within` | `boolean` | r/w | Determines whether line spacing between base lines is set to a specific number of points or lines. |
| `space after` | `real` | r/w | Returns or sets the spacing, in points, after the specified paragraphs. |
| `space before` | `real` | r/w | Returns or sets the spacing, in points, before the specified paragraphs. |
| `space within` | `real` | r/w | Returns or sets the amount of space between base lines in the specified paragraph, in points or lines. |
| `text direction` | `MsoTextDirection` | r/w | Returns or sets the text direction for the specified paragraph. |
| `word wrap` | `boolean` | r/w | Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs. |

### `play settings`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `hide while not playing` | `boolean` | r/w |  |
| `loop until stopped` | `boolean` | r/w |  |
| `pause animation` | `boolean` | r/w |  |
| `play on entry` | `boolean` | r/w |  |
| `rewind move` | `boolean` | r/w |  |
| `stop after slides` | `integer` | r/w |  |

### `player`

Represents an interface for playing movies

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `current position` | `integer` | r/w |  |
| `player state` | `EPPPlayerState` | r/o |  |

### `preferences`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `always suggest corrections` | `integer` | r/w |  |
| `append DOS extentions` | `integer` | r/w |  |
| `auto fit` | `integer` | r/w |  |
| `auto recovery frequency` | `integer` | r/w |  |
| `background spelling` | `integer` | r/w |  |
| `compress file` | `integer` | r/w |  |
| `default view` | `integer` | r/w |  |
| `drag drop` | `integer` | r/w |  |
| `end blank slide` | `integer` | r/w |  |
| `file properties prompt` | `integer` | r/w |  |
| `hide spelling errors` | `integer` | r/w |  |
| `ignore numbers in words` | `integer` | r/w |  |
| `ignore uppercase` | `integer` | r/w |  |
| `option bitmap` | `integer` | r/w |  |
| `ruler units` | `integer` | r/w |  |
| `save auto recovery` | `integer` | r/w |  |
| `show vertical ruler` | `integer` | r/w |  |
| `slide show control` | `integer` | r/w |  |
| `slide show right mouse` | `integer` | r/w |  |
| `smart cut paste` | `integer` | r/w |  |
| `smart quotes` | `integer` | r/w |  |
| `undo level count` | `integer` | r/w |  |
| `user initials` | `text` | r/w |  |
| `user name` | `text` | r/w |  |
| `word selection` | `integer` | r/w |  |

### `presentation`

*Plural:* `presentations`

**Contains:** `slide`, `color scheme`, `font`, `document window`, `document property`, `custom document property`, `design`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `default shape` | `shape` | r/o |  |
| `east asian line break level` | `EPPFarEastLineBreakLevel` | r/w | Returns or sets the East Asian line break control level for the specified paragraph. |
| `full name` | `text` | r/o |  |
| `handout master` | `master` | r/o |  |
| `has title master` | `boolean` | r/o |  |
| `is protected` | `boolean` | r/o | Returns true if the presentation is protected by Information Rights Management. |
| `layout direction` | `MsoTextDirection` | r/w |  |
| `name` | `text` | r/o |  |
| `no line break after` | `text` | r/w |  |
| `no line break before` | `text` | r/w |  |
| `notes master` | `master` | r/o |  |
| `page setup` | `page setup` | r/o |  |
| `password` | `text` | r/w | The password for opening the presentation |
| `path` | `text` | r/o |  |
| `print options` | `print options` | r/o |  |
| `read only` | `boolean` | r/o |  |
| `save as movie settings` | `save as movie settings` | r/o |  |
| `saved` | `boolean` | r/w |  |
| `section properties` | `section properties` | r/o |  |
| `slide master` | `master` | r/o |  |
| `slide show settings` | `slide show settings` | r/o |  |
| `slide show window` | `slide show window` | r/o |  |
| `template name` | `text` | r/o |  |
| `title master` | `master` | r/o |  |
| `web options` | `web options` | r/o |  |
| `write password` | `text` | r/w | The password for saving changes to the presentation |

### `presenter tool`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `current presenter slide` | `slide` | r/o |  |
| `current slide in show` | `integer` | r/o |  |
| `elapsed time` | `real` | r/o |  |
| `max slides in show` | `integer` | r/o |  |
| `notes text` | `text` | r/w |  |
| `notes zoom` | `integer` | r/w |  |
| `slide miniature` | `boolean` | r/o |  |

### `presenter view window`

*Plural:* `presenter view windows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active` | `boolean` | r/o |  |
| `height` | `real` | r/o |  |
| `presentation` | `presentation` | r/o |  |
| `presenter tool` | `presenter tool` | r/o |  |
| `width` | `real` | r/o |  |

### `print options`

**Contains:** `print range`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active printer` | `text` | r/o |  |
| `collate` | `boolean` | r/w |  |
| `fit to page` | `boolean` | r/w |  |
| `frame slides` | `boolean` | r/w |  |
| `number of copies` | `integer` | r/w |  |
| `output type` | `EPPPrintOutputType` | r/w |  |
| `print color type` | `EPPPrintColorType` | r/w |  |
| `print fonts as graphics` | `boolean` | r/w |  |
| `print hidden slides` | `boolean` | r/w |  |
| `range type` | `EPPPrintRangeType` | r/w |  |
| `section index` | `integer` | r/w |  |
| `slide show name` | `text` | r/w |  |

### `print range`

*Plural:* `print ranges`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `range end` | `integer` | r/o |  |
| `range start` | `integer` | r/o |  |

### `property effect`

**Contains:** `animation point`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `ending` | `null` | r/o |  |
| `property set effect` | `MsoAnimProperty` | r/w |  |
| `starting` | `null` | r/o |  |

### `rotating effect`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `rotating` | `real` | r/w |  |

### `ruler level`

*Plural:* `ruler levels`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `first margin` | `real` | r/w | Returns or sets the first-line indent for the specified outline level, in points. |
| `left margin` | `real` | r/w | Returns or sets the left indent for the specified outline level, in points. |

### `ruler`

Represents the ruler for the text in the specified shape or for all text in the specified text style. Contains tab stops and the indentation settings for text outline levels.

**Contains:** `tab stop`, `ruler level`

### `save as movie settings`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto loop enabled` | `boolean` | r/w |  |
| `background sound track file` | `text` | r/w |  |
| `background track segment end` | `integer` | r/w |  |
| `background track segment start` | `integer` | r/w |  |
| `background track start` | `integer` | r/w |  |
| `create movie preview` | `boolean` | r/w |  |
| `force all inline` | `boolean` | r/w |  |
| `include narration and sounds` | `boolean` | r/w |  |
| `movie actors` | `text` | r/w |  |
| `movie author` | `text` | r/w |  |
| `movie copyright` | `text` | r/w |  |
| `movie frame height` | `integer` | r/w |  |
| `movie frame width` | `integer` | r/w |  |
| `movie producer` | `text` | r/w |  |
| `optimization` | `EPPMovieOptimization` | r/w |  |
| `show movie controller` | `boolean` | r/w |  |
| `transition description` | `text` | r/w |  |
| `use single transition` | `boolean` | r/w |  |

### `scale effect`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `by x` | `real` | r/w |  |
| `by y` | `real` | r/w |  |
| `from x` | `real` | r/w |  |
| `from y` | `real` | r/w |  |
| `to x` | `real` | r/w |  |
| `to y` | `real` | r/w |  |

### `section properties`

### `selection`

Represents the selection in the specified document window

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `child shape range` | `shape range` | r/o | Returns the child shapes of a selection. |
| `has child shape range` | `boolean` | r/o | Returns whether the selection contains child shapes. |
| `selection type` | `EPPSelectionType` | r/o | Represents the type of objects in a selection. |
| `shape range` | `shape range` | r/o | Returns a collection of shapes selected on the specified slide. |
| `slide range` | `slide range` | r/o | Returns a collection of selected slides. |
| `text range` | `text range` | r/o | Returns the text and text properties of the selected text. |

### `sequence`

*Plural:* `sequences`

**Contains:** `effect`

### `set effect`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `ending` | `null` | r/w |  |
| `property set effect` | `MsoAnimProperty` | r/w |  |

### `slide range`

A collection that represents a notes page or a slide range, which is a set of slides that can contain a single slide to as many as all slides in a presentation.

**Contains:** `slide`, `shape`, `hyperlink`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color scheme` | `color scheme` | r/w | Returns or sets the color scheme for the specified slide, slide range, or slide master. |
| `custom layout` | `custom layout` | r/w | Returns the custom layout associated with the specified range of slides. |
| `design` | `design` | r/w |  |
| `display master shapes` | `boolean` | r/w | Determines whether the specified range of slides displays the background objects on the slide master. |
| `follow master background` | `boolean` | r/w | Determines whether the range of slides follows the slide master background. |
| `headers and footers` | `headers and footers` | r/o | Returns a collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides. |
| `layout` | `EPPSlideLayout` | r/w | Returns or sets the slide layout. |
| `notes page` | `slide range` | r/o | Returns a slide range object that represents the notes pages for the specified slide or range of slides. |
| `print steps` | `integer` | r/o |  |
| `slide ID` | `integer` | r/o |  |
| `slide index` | `integer` | r/o |  |
| `slide master` | `master` | r/o | Returns the slide master. |
| `slide number` | `integer` | r/o | Returns the slide number. |
| `slide show transitions` | `slide show transition` | r/o | Returns the special effects for the specified slide transition. |
| `timeline` | `timeline` | r/o | Returns the animation timeline for the slide. |

### `slide show settings`

**Contains:** `named slide show`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `advance mode` | `EPPSlideShowAdvanceMode` | r/w |  |
| `ending slide` | `integer` | r/w |  |
| `loop until stopped` | `boolean` | r/w |  |
| `pointer color` | `integer` | r/w |  |
| `pointer color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color of  default pointer color for a presentation. |
| `range type` | `EPPSlideShowRangeType` | r/w |  |
| `show type` | `EPPSlideShowType` | r/w |  |
| `show with animation` | `boolean` | r/o |  |
| `show with narration` | `boolean` | r/w |  |
| `show with presenter` | `boolean` | r/w |  |
| `slide show name` | `text` | r/w |  |
| `starting slide` | `integer` | r/w |  |

### `slide show transition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `advance on click` | `boolean` | r/w |  |
| `advance on time` | `boolean` | r/w |  |
| `advance time` | `real` | r/w |  |
| `entry effect` | `EPPEntryEffect` | r/w |  |
| `hidden` | `boolean` | r/w |  |
| `loop sound until next` | `boolean` | r/w |  |
| `sound effect transition` | `sound effect` | r/o |  |
| `transition duration` | `real` | r/w |  |

### `slide show view`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accelerations enabled` | `boolean` | r/w |  |
| `current show position` | `integer` | r/o |  |
| `is named show` | `boolean` | r/o |  |
| `last slide viewed` | `slide` | r/o |  |
| `pointer color` | `integer` | r/w |  |
| `pointer color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the color of temporary pointer color for a view of a slide show. |
| `pointer type` | `EPPSlideShowPointerType` | r/w |  |
| `presentation elapsed time` | `real` | r/o |  |
| `slide` | `slide` | r/o |  |
| `slide elapsed time` | `real` | r/w |  |
| `slide show name` | `text` | r/o |  |
| `slide state` | `EPPSlideShowState` | r/w |  |
| `zoom` | `integer` | r/o |  |

### `slide show window`

*Plural:* `slide show windows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active` | `boolean` | r/o |  |
| `bounds` | `rectangle` | r/w |  |
| `height` | `real` | r/w |  |
| `is full screen` | `boolean` | r/o |  |
| `left position` | `real` | r/w |  |
| `presentation` | `presentation` | r/o |  |
| `slideshow view` | `slide show view` | r/o |  |
| `top` | `real` | r/w |  |
| `width` | `real` | r/w |  |

### `slide`

*Plural:* `slides`

**Contains:** `shape`, `hyperlink`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `background` | `shape` | r/o |  |
| `color scheme` | `color scheme` | r/w |  |
| `custom layout` | `custom layout` | r/w |  |
| `design` | `design` | r/w |  |
| `display master shapes` | `boolean` | r/w |  |
| `follow master background` | `boolean` | r/w |  |
| `headers and footers` | `headers and footers` | r/o |  |
| `layout` | `EPPSlideLayout` | r/w |  |
| `notes page` | `slide` | r/o |  |
| `print steps` | `integer` | r/o |  |
| `section index` | `integer` | r/o |  |
| `section number` | `integer` | r/o |  |
| `slide ID` | `integer` | r/o |  |
| `slide index` | `integer` | r/o |  |
| `slide master` | `master` | r/o |  |
| `slide number` | `integer` | r/o |  |
| `slide show transition` | `slide show transition` | r/o |  |
| `timeline` | `timeline` | r/o |  |

### `sound effect`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o |  |
| `sound type` | `EPPSoundEffectType` | r/w |  |

### `tab stop`

Represents a single tab stop.

*Plural:* `tab stops`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `tab position` | `real` | r/w | Returns or sets the position of a tab stop relative to the left margin. |
| `tab stop type` | `MsoTabStopType` | r/w | Returns or sets the type of the tab stop object. |

### `text style level`

*Plural:* `text style levels`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `font` | `font` | r/o |  |
| `paragraph format` | `paragraph format` | r/o |  |

### `text style`

**Contains:** `text style level`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `ruler` | `ruler` | r/o |  |
| `text frame` | `text frame` | r/o |  |

### `timeline`

**Contains:** `sequence`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `main sequence` | `sequence` | r/o |  |

### `timing`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `acceleration` | `real` | r/w |  |
| `autoreverse` | `boolean` | r/w |  |
| `deceleration` | `real` | r/w |  |
| `duration` | `real` | r/w |  |
| `repeat count` | `integer` | r/w |  |
| `repeat duration` | `real` | r/w |  |
| `restart` | `MsoAnimEffectRestart` | r/w |  |
| `rewind` | `boolean` | r/w |  |
| `smooth end` | `boolean` | r/w |  |
| `smooth start` | `boolean` | r/w |  |
| `speed` | `real` | r/w |  |

### `two initial caps exception`

Represents a single initial-capital autocorrect exception.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entry_index` | `integer` | r/o | Returns the index for the position of the object in its container element list. |
| `name` | `text` | r/o | Returns the name of the TwoInitialCapsException. |

### `view`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `display slide miniature` | `boolean` | r/w |  |
| `slide` | `slide` | r/w |  |
| `view type` | `EPPViewType` | r/o |  |
| `zoom` | `integer` | r/w |  |
| `zoom to fit` | `boolean` | r/w |  |

### `web options`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `allow PNG` | `boolean` | r/w |  |
| `encoding` | `MsoEncoding` | r/w |  |

### Enumerations

### `EPPWindowState`

| Value | Description |
|-------|-------------|
| `window normal` |  |
| `window minimized` |  |

### `EPPArrangeStyle`

| Value | Description |
|-------|-------------|
| `arrange tiled` |  |
| `arrange cascade` |  |

### `EPPViewType`

| Value | Description |
|-------|-------------|
| `slide view` |  |
| `master view` |  |
| `page view` |  |
| `handout master view` |  |
| `notes master view` |  |
| `outline view` |  |
| `slide sorter view` |  |
| `title master view` |  |
| `normal view` |  |
| `print preview` |  |
| `thumbnail view` |  |

### `EPPColorSchemeIndex`

| Value | Description |
|-------|-------------|
| `scheme color unset` |  |
| `not a scheme color` |  |
| `background scheme` |  |
| `foreground scheme` |  |
| `shadow scheme` |  |
| `title scheme` |  |
| `fill scheme` |  |
| `accent1 scheme` |  |
| `accent2 scheme` |  |
| `accent3 scheme` |  |

### `EPPSlideSizeType`

| Value | Description |
|-------|-------------|
| `slide size on screen` |  |
| `slide size letter paper` |  |
| `slide size A4 paper` |  |
| `slide size 35 MM` |  |
| `slide size overhead` |  |
| `slide size banner` |  |
| `slide size custom` |  |
| `slide size ledger paper` |  |
| `slide size A3 paper` |  |
| `slide size B4 ISO paper` |  |
| `slide size B5 ISO paper` |  |
| `slide size B4 JIS paper` |  |
| `slide size B5 JIS paper` |  |
| `slide size hagaki card` |  |
| `slide size on screen 16x9` |  |
| `slide size on screen 16x10` |  |

### `EPPSaveAsFileType`

| Value | Description |
|-------|-------------|
| `save as presentation` |  |
| `save as template` |  |
| `save as RTF` |  |
| `save as show` |  |
| `save as addIn` |  |
| `save as default` |  |
| `save as HTML` |  |
| `save as movie` |  |
| `save as package` |  |
| `save as PDF` |  |
| `save as Open XML presentation` |  |
| `save as Open XML presentation macro enabled` |  |
| `save as Open XML show` |  |
| `save as Open XML show macro enabled` |  |
| `save as Open XML template` |  |
| `save as Open XML template macro enabled` |  |
| `save as Open XML theme` |  |
| `save as GIF` |  |
| `save as JPG` |  |
| `save as PNG` |  |
| `save as BMP` |  |
| `save as TIF` |  |

### `PpTextStyleType`

| Value | Description |
|-------|-------------|
| `text style default` |  |
| `text style title` |  |
| `text style body` |  |

### `EPPSlideLayout`

| Value | Description |
|-------|-------------|
| `slide layout unset` |  |
| `slide layout title slide` |  |
| `slide layout text slide` |  |
| `slide layout two column text` |  |
| `slide layout table` |  |
| `slide layout text and chart` |  |
| `slide layout chart and text` |  |
| `slide layout orgchart` |  |
| `slide layout chart` |  |
| `slide layout text and clipart` |  |
| `slide layout clipart and text` |  |
| `slide layout title only` |  |
| `slide layout blank` |  |
| `slide layout text and object` |  |
| `slide layout object and text` |  |
| `slide layout large object` |  |
| `slide layout object` |  |
| `slide layout media clip` |  |
| `slide layout media clip and text` |  |
| `slide layout object over text` |  |
| `slide layout text over object` |  |
| `slide layout text and two objects` |  |
| `slide layout two objects and text` |  |
| `slide layout two objects over text` |  |
| `slide layout four objects` |  |
| `slide layout vertical text` |  |
| `slide layout clipart and vertical text` |  |
| `slide layout vertical title and text` |  |
| `slide layout vertical title and text over chart` |  |
| `slide layout two objects` |  |
| `slide layout object and two objects` |  |
| `slide layout two objects and object` |  |
| `slide layout custom` |  |
| `slide layout section header` |  |
| `slide layout comparison` |  |
| `slide layout content with caption` |  |
| `slide layout picture with caption` |  |

### `EPPEntryEffect`

| Value | Description |
|-------|-------------|
| `entry effect airplane left` |  |
| `entry effect airplane right` |  |
| `entry effect appear` |  |
| `entry effect blinds horizontal` |  |
| `entry effect blinds vertical` |  |
| `entry effect box down` |  |
| `entry effect box in` |  |
| `entry effect box left` |  |
| `entry effect box out` |  |
| `entry effect box right` |  |
| `entry effect box up` |  |
| `entry effect checkerboard across` |  |
| `entry effect checkerboard down` |  |
| `entry effect circle` |  |
| `entry effect collapse across` |  |
| `entry effect collapse bottom` |  |
| `entry effect collapse left` |  |
| `entry effect collapse right` |  |
| `entry effect collapse up` |  |
| `entry effect crush` |  |
| `entry effect comb horizontal` |  |
| `entry effect comb vertical` |  |
| `entry effect conveyor left` |  |
| `entry effect conveyor right` |  |
| `entry effect cover down` |  |
| `entry effect cover left down` |  |
| `entry effect cover left up` |  |
| `entry effect cover left` |  |
| `entry effect cover right down` |  |
| `entry effect cover right up` |  |
| `entry effect cover right` |  |
| `entry effect cover up` |  |
| `entry effect crawl from down` |  |
| `entry effect crawl from left` |  |
| `entry effect crawl from right` |  |
| `entry effect crawl from up` |  |
| `entry effect cube down` |  |
| `entry effect cube left` |  |
| `entry effect cube right` |  |
| `entry effect cube up` |  |
| `entry effect curtains` |  |
| `entry effect cut black` |  |
| `entry effect cut` |  |
| `entry effect diamond` |  |
| `entry effect dissolve` |  |
| `entry effect doors horizontal` |  |
| `entry effect doors vertical` |  |
| `entry effect drape left` |  |
| `entry effect drape right` |  |
| `entry effect fade black` |  |
| `entry effect fade fly from bottom left` |  |
| `entry effect fade fly from bottom right` |  |
| `entry effect fade fly from bottom` |  |
| `entry effect fade fly from left` |  |
| `entry effect fade fly from right` |  |
| `entry effect fade fly from top left` |  |
| `entry effect fade fly from top right` |  |
| `entry effect fade fly from top` |  |
| `entry effect fade smoothly` |  |
| `entry effect fade` |  |
| `entry effect fall over left` |  |
| `entry effect fall over right` |  |
| `entry effect ferris wheel left` |  |
| `entry effect ferris wheel right` |  |
| `entry effect flash once fast` |  |
| `entry effect flash once medium` |  |
| `entry effect flash once slow` |  |
| `entry effect flashbulb` |  |
| `entry effect flip down` |  |
| `entry effect flip left` |  |
| `entry effect flip right` |  |
| `entry effect flip up` |  |
| `entry effect fly from bottom left` |  |
| `entry effect fly from bottom right` |  |
| `entry effect fly from bottom` |  |
| `entry effect fly from left` |  |
| `entry effect fly from right` |  |
| `entry effect fly from top left` |  |
| `entry effect fly from top right` |  |
| `entry effect fly from top` |  |
| `entry effect fly through in bounce` |  |
| `entry effect fly through in` |  |
| `entry effect fly through out bounce` |  |
| `entry effect fly through out` |  |
| `entry effect fracture` |  |
| `entry effect gallery left` |  |
| `entry effect gallery right` |  |
| `entry effect glitter diamond down` |  |
| `entry effect glitter diamond left` |  |
| `entry effect glitter diamond right` |  |
| `entry effect glitter diamond up` |  |
| `entry effect glitter hexagon down` |  |
| `entry effect glitter hexagon left` |  |
| `entry effect glitter hexagon right` |  |
| `entry effect glitter hexagon up` |  |
| `entry effect honeycomb` |  |
| `entry effect morph by object` |  |
| `entry effect morph by word` |  |
| `entry effect morph by char` |  |
| `entry effect news flash` |  |
| `entry effect none` |  |
| `entry effect orbit down` |  |
| `entry effect orbit left` |  |
| `entry effect orbit right` |  |
| `entry effect orbit up` |  |
| `entry effect origami left` |  |
| `entry effect origami right` |  |
| `entry effect page curl double left` |  |
| `entry effect page curl double right` |  |
| `entry effect page curl single left` |  |
| `entry effect page curl single right` |  |
| `entry effect pan down` |  |
| `entry effect pan left` |  |
| `entry effect pan right` |  |
| `entry effect pan up` |  |
| `entry effect peek from down` |  |
| `entry effect peek from left` |  |
| `entry effect peek from right` |  |
| `entry effect peek from up` |  |
| `entry effect peel off left` |  |
| `entry effect peel off right` |  |
| `entry effect plus` |  |
| `entry effect prestige` |  |
| `entry effect push down` |  |
| `entry effect push left` |  |
| `entry effect push right` |  |
| `entry effect push up` |  |
| `entry effect random bars horizontal` |  |
| `entry effect random bars vertical` |  |
| `entry effect random` |  |
| `entry effect reveal black left` |  |
| `entry effect reveal black right` |  |
| `entry effect reveal smooth left` |  |
| `entry effect reveal smooth right` |  |
| `entry effect ripple center` |  |
| `entry effect ripple left down` |  |
| `entry effect ripple left up` |  |
| `entry effect ripple right down` |  |
| `entry effect ripple right up` |  |
| `entry effect rotate down` |  |
| `entry effect rotate left` |  |
| `entry effect rotate right` |  |
| `entry effect rotate up` |  |
| `entry effect shred rectangle in` |  |
| `entry effect shred rectangle out` |  |
| `entry effect shred strips in` |  |
| `entry effect shred strips out` |  |
| `entry effect spinner` |  |
| `entry effect spiral` |  |
| `entry effect split horizontal in` |  |
| `entry effect split horizontal out` |  |
| `entry effect split vertical in` |  |
| `entry effect split vertical out` |  |
| `entry effect stretch across` |  |
| `entry effect stretch down` |  |
| `entry effect stretch left` |  |
| `entry effect stretch right` |  |
| `entry effect stretch up` |  |
| `entry effect strips down left` |  |
| `entry effect strips down right` |  |
| `entry effect strips left down` |  |
| `entry effect strips left up` |  |
| `entry effect strips right down` |  |
| `entry effect strips right up` |  |
| `entry effect strips up left` |  |
| `entry effect strips up right` |  |
| `entry effect switch down` |  |
| `entry effect switch left` |  |
| `entry effect switch right` |  |
| `entry effect switch up` |  |
| `entry effect swivel` |  |
| `entry effect uncover down` |  |
| `entry effect uncover left down` |  |
| `entry effect uncover left up` |  |
| `entry effect uncover left` |  |
| `entry effect uncover right down` |  |
| `entry effect uncover right up` |  |
| `entry effect uncover right` |  |
| `entry effect uncover up` |  |
| `entry effect unset` |  |
| `entry effect vortex down` |  |
| `entry effect vortex left` |  |
| `entry effect vortex right` |  |
| `entry effect vortex up` |  |
| `entry effect warp in` |  |
| `entry effect warp out` |  |
| `entry effect wedge` |  |
| `entry effect wheel 1 spoke` |  |
| `entry effect wheel 2 spokes` |  |
| `entry effect wheel 3 spokes` |  |
| `entry effect wheel 4 spokes` |  |
| `entry effect wheel 8 spokes` |  |
| `entry effect wheel reverse 1 spoke` |  |
| `entry effect wind left` |  |
| `entry effect wind right` |  |
| `entry effect window horizontal` |  |
| `entry effect window vertical` |  |
| `entry effect wipe down` |  |
| `entry effect wipe left` |  |
| `entry effect wipe right` |  |
| `entry effect wipe up` |  |
| `entry effect zoom bottom` |  |
| `entry effect zoom center` |  |
| `entry effect zoom in slightly` |  |
| `entry effect zoom in` |  |
| `entry effect zoom out slightly` |  |
| `entry effect zoom out` |  |

### `EPPTextLevelEffect`

| Value | Description |
|-------|-------------|
| `animation level unset` |  |
| `animate level none` |  |
| `animate level first level` |  |
| `animate level second level` |  |
| `animate level third level` |  |
| `animate level fourth level` |  |
| `animate level fifth level` |  |
| `animate level all levels` |  |

### `EPPTextUnitEffect`

| Value | Description |
|-------|-------------|
| `animation unit unset` |  |
| `text unit effect by paragraph` |  |
| `text unit effect by word` |  |
| `text unit effect by character` |  |

### `EPPChartUnitEffect`

| Value | Description |
|-------|-------------|
| `animation chart unset` |  |
| `chart unit effect by series` |  |
| `chart unit effect by category` |  |
| `chart unit effect by series element` |  |

### `EPPAfterEffect`

| Value | Description |
|-------|-------------|
| `after effect unset` |  |
| `after effect none` |  |
| `after effect hide` |  |
| `after effect dim` |  |

### `EPPAdvanceMode`

| Value | Description |
|-------|-------------|
| `advance mode unset` |  |
| `advance mode on click` |  |

### `EPPSoundEffectType`

| Value | Description |
|-------|-------------|
| `sound effect unset` |  |
| `sound effect none` |  |
| `sound effect stop previous` |  |
| `sound effect file` |  |

### `EPPUpdateOption`

| Value | Description |
|-------|-------------|
| `update option unset` |  |
| `update option manual` |  |

### `EPPChangeCase`

| Value | Description |
|-------|-------------|
| `ppCaseSentence` |  |
| `ppCaseLower` |  |
| `ppCaseUpper` |  |
| `ppCaseTitle` |  |

### `EPPDialogMode`

| Value | Description |
|-------|-------------|
| `dialog mode unset` |  |
| `dialog mode modless` |  |
| `dialog mode modal` |  |

### `EPPDialogStyle`

| Value | Description |
|-------|-------------|
| `dialog style unset` |  |
| `dialog style standard` |  |

### `EPPSlideShowPointerType`

| Value | Description |
|-------|-------------|
| `slide show pointer none` |  |
| `slide show pointer arrow` |  |
| `slide show pointer pen` |  |
| `slide show pointer always hidden` |  |

### `EPPSlideShowState`

| Value | Description |
|-------|-------------|
| `slide show state running` |  |
| `slide show state paused` |  |
| `slide show state black screen` |  |
| `slide show state white screen` |  |
| `slide show state done` |  |

### `EPPPlayerState`

| Value | Description |
|-------|-------------|
| `player state playing` |  |
| `player state paused` |  |
| `player state stopped` |  |
| `player state not ready` |  |

### `EPPSlideShowAdvanceMode`

| Value | Description |
|-------|-------------|
| `slide show advance manual advance` |  |
| `slide show advance use slide timings` |  |

### `EPPPrintOutputType`

| Value | Description |
|-------|-------------|
| `print slides` |  |
| `print two slide handouts` |  |
| `print three slide handouts` |  |
| `print six slide handouts` |  |
| `print notes pages` |  |
| `print outline` |  |
| `print four slide handouts` |  |
| `print nine slide handouts` |  |

### `EPPPrintColorType`

| Value | Description |
|-------|-------------|
| `print color` |  |
| `print black and white` |  |

### `EPPSelectionType`

| Value | Description |
|-------|-------------|
| `selection type none` |  |
| `selection type slides` |  |
| `selection type shapes` |  |
| `selection type text` |  |

### `EPPDirection`

| Value | Description |
|-------|-------------|
| `direction unset` |  |
| `left to right` |  |

### `EPPDateTimeFormat`

| Value | Description |
|-------|-------------|
| `unset` |  |
| `Mdyy` |  |
| `ddddMMMMddyyyy` |  |
| `MMMMyyyy` |  |
| `MMMMdyyyy` |  |
| `MMMyy` |  |
| `MMMMyy` |  |
| `MMyy` |  |
| `MMddyyHmm` |  |
| `MddyyhmmAMPM` |  |
| `Hmm` |  |
| `Hmmss` |  |
| `hmmAMPM` |  |
| `hmmssAMPM` |  |

### `EPPTransitionSpeed`

| Value | Description |
|-------|-------------|
| `transition speed unset` |  |
| `transistion speed slow` |  |
| `transistion speed medium` |  |

### `EPPMouseActivation`

| Value | Description |
|-------|-------------|
| `mouse activation mouse click` |  |
| `mouse activation mouse over` |  |

### `EPPActionType`

| Value | Description |
|-------|-------------|
| `action type unset` |  |
| `action type none` |  |
| `action type next slide` |  |
| `action type previous slide` |  |
| `action type first slide` |  |
| `action type last slide` |  |
| `action type last slide viewed` |  |
| `action type end show` |  |
| `action type hyperlink action` |  |
| `action type run macro` |  |
| `action type run program` |  |
| `action type named slideshow action` |  |
| `action type OLE verb` |  |

### `EPPPlaceholderType`

| Value | Description |
|-------|-------------|
| `placeholder type unset` |  |
| `placeholder type title placeholder` |  |
| `placeholder type body placeholder` |  |
| `placeholder type center title placeholder` |  |
| `placeholder type subtitle placeholder` |  |
| `placeholder type vertical title placeholder` |  |
| `placeholder type vertical body placeholder` |  |
| `placeholder type object placeholder` |  |
| `placeholder type chart placeholder` |  |
| `placeholder type bitmap placeholder` |  |
| `placeholder type media clip placeholder` |  |
| `placeholder type org chart placeholder` |  |
| `placeholder type table placeholder` |  |
| `placeholder type slide number placeholder` |  |
| `placeholder type header placeholder` |  |
| `placeholder type footer placeholder` |  |
| `placeholder type date placeholder` |  |
| `placeholder type vertical object placeholder` |  |
| `placeholder type picture placeholder` |  |
| `placeholder type cameo placeholder` |  |

### `EPPSlideShowType`

| Value | Description |
|-------|-------------|
| `slide show type speaker` |  |
| `slide show type window` |  |
| `slide show type kiosk` |  |
| `slide show type presenter` |  |

### `EPPPrintRangeType`

| Value | Description |
|-------|-------------|
| `print range all` |  |
| `print range selection` |  |
| `print range current` |  |
| `print range slide range` |  |
| `print section` |  |

### `EPPAutoSize`

| Value | Description |
|-------|-------------|
| `ppAutoSizemixed` |  |
| `ppAutoSizeNone` |  |
| `ppAutoSizeShapeToFitText` |  |
| `ppAutoSizeTextToFitShape` |  |

### `EPPMediaType`

| Value | Description |
|-------|-------------|
| `media type unset` |  |
| `media type other` |  |
| `media type sound` |  |
| `media type movie` |  |

### `EPPSoundFormatType`

| Value | Description |
|-------|-------------|
| `sound format unset` |  |
| `sound format none` |  |
| `sound format WAV` |  |
| `sound format MIDI` |  |

### `EPPFarEastLineBreakLevel`

| Value | Description |
|-------|-------------|
| `east asian line break normal` |  |
| `east asian line break strict` |  |
| `east asian line break custom` |  |

### `EPPSlideShowRangeType`

| Value | Description |
|-------|-------------|
| `slide show range show all` |  |
| `slide show range` |  |
| `slide show range named slideshow` |  |

### `EPPFrameColors`

| Value | Description |
|-------|-------------|
| `frame colors browser colors` |  |
| `frame colors presentation scheme text color` |  |
| `frame colors presentation scheme accent color` |  |
| `frame colors white text on black` |  |
| `frame colors black text on white` |  |

### `EPPMovieOptimization`

| Value | Description |
|-------|-------------|
| `movie optimization normal` |  |
| `movie optimization size` |  |
| `movie optimization speed` |  |
| `movie optimization quality` |  |

### `EPPBulletType`

| Value | Description |
|-------|-------------|
| `ppBulletmixed` |  |
| `ppBulletNone` |  |
| `ppBulletUnnumbered` |  |
| `ppBulletNumbered` |  |
| `ppBulletPicture` |  |

### `EPPNumberedBulletStyle`

| Value | Description |
|-------|-------------|
| `ppBulletStylemixed` |  |
| `ppBulletAlphaLCPeriod` |  |
| `ppBulletAlphaUCPeriod` |  |
| `ppBulletArabicParenRight` |  |
| `ppBulletArabicPeriod` |  |
| `ppBulletRomanLCParenBoth` |  |
| `ppBulletRomanLCParenRight` |  |
| `ppBulletRomanLCPeriod` |  |
| `ppBulletRomanUCPeriod` |  |
| `ppBulletAlphaLCParenBoth` |  |
| `ppBulletAlphaLCParenRight` |  |
| `ppBulletAlphaUCParenBoth` |  |
| `ppBulletAlphaUCParenRight` |  |
| `ppBulletArabicParenBoth` |  |
| `ppBulletArabicPlain` |  |
| `ppBulletRomanUCParenBoth` |  |
| `ppBulletRomanUCParenRight` |  |
| `ppBulletSimpChinPlain` |  |
| `ppBulletSimpChinPeriod` |  |
| `ppBulletCircleNumDBPlain` |  |
| `ppBulletCircleNumWDWhitePlain` |  |
| `ppBulletCircleNumWDBlackPlain` |  |
| `ppBulletTradChinPlain` |  |
| `ppBulletTradChinPeriod` |  |
| `ppBulletArabicAlphaDash` |  |
| `ppBulletArabicAbjadDash` |  |
| `ppBulletHebrewAlphaDash` |  |
| `ppBulletKanjiKoreanPlain` |  |
| `ppBulletKanjiKoreanPeriod` |  |
| `ppBulletArabicDBPlain` |  |

### `EPPShapeFormat`

| Value | Description |
|-------|-------------|
| `shape format GIF` |  |
| `shape format JPEG` |  |
| `shape format PNG` |  |
| `shape format PICT` |  |

### `EPPExportMode`

| Value | Description |
|-------|-------------|
| `export mode relative to slide` |  |
| `export mode clip relative to slide` |  |
| `export mode scale to fit` |  |
| `export mode scale XY` |  |

### `PpBorderType`

| Value | Description |
|-------|-------------|
| `top border` |  |
| `left border` |  |
| `bottom border` |  |
| `right border` |  |
| `diagonal down border` |  |
| `diagonal up border` |  |

### `EPPCheckInVersionType`

| Value | Description |
|-------|-------------|
| `minor version` |  |
| `major version` |  |
| `overwrite current version` |  |

### `IppPageLayout`

| Value | Description |
|-------|-------------|
| `page layout normal` |  |
| `page layout full screen` |  |

### `IppButtonsType`

| Value | Description |
|-------|-------------|
| `regular` |  |
| `fancy` |  |
| `text only` |  |

### `IppNavBarPlacement`

| Value | Description |
|-------|-------------|
| `bar placement top` |  |
| `bar placement bottom` |  |

### `MsoAnimEffect`

| Value | Description |
|-------|-------------|
| `animation type custom` |  |
| `animation type appear` |  |
| `animation type fly` |  |
| `animation type blinds` |  |
| `animation type box` |  |
| `animation type checkerboard` |  |
| `animation type circle` |  |
| `animation type crawl` |  |
| `animation type diamond` |  |
| `animation type dissolve` |  |
| `animation type fade` |  |
| `animation type flash once` |  |
| `animation type peek` |  |
| `animation type plus` |  |
| `animation type random bars` |  |
| `animation type spiral` |  |
| `animation type split` |  |
| `animation type stretch` |  |
| `animation type strips` |  |
| `animation type swivel` |  |
| `animation type wedge` |  |
| `animation type wheel` |  |
| `animation type wipe` |  |
| `animation type zoom` |  |
| `animation type random effect` |  |
| `animation type boomerang` |  |
| `animation type bounce` |  |
| `animation type color reveal` |  |
| `animation type credits` |  |
| `animation type ease in` |  |
| `animation type float` |  |
| `animation type grow and turn` |  |
| `animation type light speed` |  |
| `animation type pinwheel` |  |
| `animation type rise up` |  |
| `animation type swish` |  |
| `animation type thin line` |  |
| `animation type unfold` |  |
| `animation type whip` |  |
| `animation type ascend` |  |
| `animation type center revolve` |  |
| `animation type faded swivel` |  |
| `animation type descend` |  |
| `animation type sling` |  |
| `animation type spinner` |  |
| `animation type stretchy` |  |
| `animation type zip` |  |
| `animation type arc up` |  |
| `animation type fade zoom` |  |
| `animation type glide` |  |
| `animation type expand` |  |
| `animation type flip` |  |
| `animation type shimmer` |  |
| `animation type fold` |  |
| `animation type change fill color` |  |
| `animation type change font` |  |
| `animation type change font color` |  |
| `animation type change font size` |  |
| `animation type change font style` |  |
| `animation type grow shrink` |  |
| `animation type change line color` |  |
| `animation type spin` |  |
| `animation type transparency` |  |
| `animation type bold flash` |  |
| `animation type blast` |  |
| `animation type bold reveal` |  |
| `animation type brush on color` |  |
| `animation type brush on underline` |  |
| `animation type color blend` |  |
| `animation type color wave` |  |
| `animation type complementary color` |  |
| `animation type complementary color 2` |  |
| `animation type constrasting color` |  |
| `animation type darken` |  |
| `animation type desaturate` |  |
| `animation type flash bulb` |  |
| `animation type flicker` |  |
| `animation type grow with color` |  |
| `animation type lighten` |  |
| `animation type style emphasis` |  |
| `animation type teeter` |  |
| `animation type vertical grow` |  |
| `animation type wave` |  |
| `animation type media play` |  |
| `animation type media pause` |  |
| `animation type media stop` |  |
| `animation type circle path` |  |
| `animation type right triangle path` |  |
| `animation type diamond path` |  |
| `animation type hexagon path` |  |
| `animation type 5 point star path` |  |
| `animation type crescent moon path` |  |
| `animation type square path` |  |
| `animation type trapezoid path` |  |
| `animation type heart path` |  |
| `animation type octagon path` |  |
| `animation type 6 point star path` |  |
| `animation type football path` |  |
| `animation type equal triangle path` |  |
| `animation type parallelogram path` |  |
| `animation type pentagon path` |  |
| `animation type 4 point star path` |  |
| `animation type 8 point star path` |  |
| `animation type teardrop path` |  |
| `animation type pointy star path` |  |
| `animation type curved square path` |  |
| `animation type curved x path` |  |
| `animation type vertical figure 8 path` |  |
| `animation type curvy star path` |  |
| `animation type loop de loop path` |  |
| `animation type buzzsaw path` |  |
| `animation type horizontal figure 8 path` |  |
| `animation type peanut path` |  |
| `animation type figure 8 four path` |  |
| `animation type neutron path` |  |
| `animation type swoosh path` |  |
| `animation type bean path` |  |
| `animation type plus path` |  |
| `animation type inverted triangle path` |  |
| `animation type inverted square path` |  |
| `animation type left path` |  |
| `animation type turn right path` |  |
| `animation type arc down path` |  |
| `animation type zigzag path` |  |
| `animation type s curve 2 path` |  |
| `animation type sine wave path` |  |
| `animation type bounce left path` |  |
| `animation type down path` |  |
| `animation type turn up path` |  |
| `animation type arc up path` |  |
| `animation type heartbeat path` |  |
| `animation type spiral right path` |  |
| `animation type wave path` |  |
| `animation type curvy left path` |  |
| `animation type diagonal down right path` |  |
| `animation type turn down path` |  |
| `animation type arc left path` |  |
| `animation type funnel path` |  |
| `animation type spring path` |  |
| `animation type bounce right path` |  |
| `animation type spiral left path` |  |
| `animation type diagonal up right path` |  |
| `animation type turn up right path` |  |
| `animation type arc right path` |  |
| `animation type s curve 1 path` |  |
| `animation type decaying wave path` |  |
| `animation type curvy right path` |  |
| `animation type stairs down path` |  |
| `animation type up path` |  |
| `animation type right path` |  |

### `MsoAnimateByLevel`

| Value | Description |
|-------|-------------|
| `text by no levels` |  |
| `text by all levels` |  |
| `text by first level` |  |
| `text by second level` |  |
| `text by third level` |  |
| `text by fourth level` |  |
| `text by fifth level` |  |
| `chart all at once` |  |
| `chart by category` |  |
| `chart by ctageory elements` |  |
| `chart by series` |  |
| `chart by series elements` |  |

### `MsoAnimTriggerType`

| Value | Description |
|-------|-------------|
| `no trigger` |  |
| `on page click` |  |
| `with previous` |  |
| `after previous` |  |
| `on shape click` |  |

### `MsoAnimAfterEffect`

| Value | Description |
|-------|-------------|
| `no after effect` |  |
| `dim` |  |
| `hide` |  |
| `hide on next click` |  |

### `MsoAnimTextUnitEffect`

| Value | Description |
|-------|-------------|
| `by paragraph` |  |
| `by character` |  |
| `by word` |  |

### `MsoAnimEffectRestart`

| Value | Description |
|-------|-------------|
| `restart always` |  |
| `restart when off` |  |
| `never restart` |  |

### `MsoAnimEffectAfter`

| Value | Description |
|-------|-------------|
| `after freeze` |  |
| `after remove` |  |
| `after hold` |  |
| `after transition` |  |

### `MsoAnimDirection`

| Value | Description |
|-------|-------------|
| `no direction` |  |
| `up` |  |
| `right` |  |
| `down` |  |
| `left` |  |
| `ordinal mask` |  |
| `up left` |  |
| `up right` |  |
| `down right` |  |
| `down left` |  |
| `top` |  |
| `bottom` |  |
| `top left` |  |
| `top right` |  |
| `bottom right` |  |
| `bottom left` |  |
| `horizontal` |  |
| `vertical` |  |
| `across` |  |
| `inward` |  |
| `out` |  |
| `clockwise` |  |
| `counterclockwise` |  |
| `horizontal in` |  |
| `horizontal out` |  |
| `vertical in` |  |
| `vertical out` |  |
| `slightly` |  |
| `center` |  |
| `in slightly` |  |
| `in center` |  |
| `in bottom` |  |
| `out slightly` |  |
| `out center` |  |
| `out bottom` |  |
| `font bold` |  |
| `font italic` |  |
| `font underline` |  |
| `font strikethrough` |  |
| `font shadow` |  |
| `font all caps` |  |
| `instant` |  |
| `gradual` |  |
| `cycle clockwise` |  |
| `cycle counterclockwise` |  |

### `MsoAnimType`

| Value | Description |
|-------|-------------|
| `animation type none` |  |
| `animation type motion` |  |
| `animation type color` |  |
| `animation type scale` |  |
| `animation type rotation` |  |
| `animation type property` |  |
| `animation type command` |  |
| `animation type filter` |  |
| `animation type set` |  |

### `MsoAnimAdditive`

| Value | Description |
|-------|-------------|
| `no additive` |  |
| `motion` |  |

### `MsoAnimAccumulate`

| Value | Description |
|-------|-------------|
| `no accumulate` |  |
| `always` |  |

### `MsoAnimProperty`

| Value | Description |
|-------|-------------|
| `no property` |  |
| `x` |  |
| `y` |  |
| `width` |  |
| `height` |  |
| `opacity` |  |
| `rotation` |  |
| `colors` |  |
| `visibility` |  |
| `text font bold` |  |
| `text font color` |  |
| `text font emboss` |  |
| `text font italic` |  |
| `text font name` |  |
| `text font shadow` |  |
| `text font size` |  |
| `text font subscript` |  |
| `text font superscript` |  |
| `text font underline` |  |
| `text font strikethrough` |  |
| `text bullet character` |  |
| `text bullet fontName` |  |
| `text bullet number` |  |
| `text bullet color` |  |
| `text bullet relative size` |  |
| `text bullet style` |  |
| `text bullet type` |  |
| `shape picture contrast` |  |
| `shape picture brightness` |  |
| `shape picture gamma` |  |
| `shape picture grayscale` |  |
| `shape fill on` |  |
| `shape fill color` |  |
| `shape fill opacity` |  |
| `shape fill back color` |  |
| `shape line on` |  |
| `shape line color` |  |
| `shape shadow on` |  |
| `shape shadow type` |  |
| `shape shadow color` |  |
| `shape shadow opacity` |  |
| `shape shadow offset X` |  |
| `shape shadow offset Y` |  |

### `MsoAnimCommandType`

| Value | Description |
|-------|-------------|
| `event` |  |
| `call` |  |
| `verb` |  |

### `MsoAnimFilterEffectType`

| Value | Description |
|-------|-------------|
| `no filter effect type` |  |
| `filter effect type barn` |  |
| `filter effect type blinds` |  |
| `filter effect type box` |  |
| `filter effect type checkerboard` |  |
| `filter effect type circle` |  |
| `filter effect type diamond` |  |
| `filter effect type dissolve` |  |
| `filter effect type fade` |  |
| `filter effect type image` |  |
| `filter effect type pixelate` |  |
| `filter effect type plus` |  |
| `filter effect type random bar` |  |
| `filter effect type slide` |  |
| `filter effect type stretch` |  |
| `filter effect type strips` |  |
| `filter effect type wedge` |  |
| `filter effect type wheel` |  |
| `filter effect type wipe` |  |

### `MsoAnimFilterEffectSubtype`

| Value | Description |
|-------|-------------|
| `no effect subtype` |  |
| `filter effect subtype in vertical` |  |
| `filter effect subtype out vertical` |  |
| `filter effect subtype in horizontal` |  |
| `filter effect subtype out horizontal` |  |
| `filter effect subtype horizontal` |  |
| `filter effect subtype vertical` |  |
| `filter effect subtype inward` |  |
| `filter effect subtype out` |  |
| `filter effect subtype across` |  |
| `filter effect subtype from left` |  |
| `filter effect subtype from right` |  |
| `filter effect subtype from top` |  |
| `filter effect subtype from bottom` |  |
| `filter effect subtype down left` |  |
| `filter effect subtype up left` |  |
| `filter effect subtype down right` |  |
| `filter effect subtype up right` |  |
| `filter effect subtype spoke 1` |  |
| `filter effect subtype spokes 2` |  |
| `filter effect subtype spokes 3` |  |
| `filter effect subtype spokes 4` |  |
| `filter effect subtype spokes 8` |  |
| `filter effect subtype left` |  |
| `filter effect subtype right` |  |
| `filter effect subtype down` |  |
| `filter effect subtype up` |  |

### `get player from`

| Value | Description |
|-------|-------------|
| `view` |  |
| `slide show view` |  |

### `paste object`

| Value | Description |
|-------|-------------|
| `view` |  |
| `presentation` |  |

### `apply theme`

| Value | Description |
|-------|-------------|
| `slide` |  |
| `master` |  |
| `presentation` |  |

---

## Drawing Suite

Classes and Methods used for Graphic Objects

### Commands

### `align`

Aligns the shapes in the specified range of shapes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape range` | no |  |
| `align type` | `MsoAlignCmd` | no |  |
| `relative to slide` | `boolean` | yes |  |

### `apply`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4004` | no |  |

### `automatic length`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4024` | no |  |

### `begin connect`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4019` | no |  |
| `connected shape` | `shape` | no |  |
| `connection site` | `integer` | no |  |

### `begin disconnect`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4020` | no |  |

### `copy shape`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `copy shape range`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape range` | no |  |

### `custom drop`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4025` | no |  |
| `drop amount` | `real` | no |  |

### `custom length`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4026` | no |  |
| `length` | `real` | no |  |

### `cut shape`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape` | no |  |

### `cut shape range`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape range` | no |  |

### `delete gradient stop`

Removes a gradient stop.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `fill format` | no |  |
| `stop index` | `integer` | no | The index number of the stop. |

### `distribute`

Evenly distributes the shapes in the specified range of shapes. You can specify whether you want to distribute the shapes horizontally or vertically and whether you want to distribute them over the entire slide or over the space they originally occupy.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape range` | no |  |
| `distribute type` | `MsoDistributeCmd` | no |  |
| `relative to slide` | `boolean` | yes |  |

### `end connect`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4021` | no |  |
| `connected shape` | `shape` | no |  |
| `connection site` | `integer` | no |  |

### `end disconnect`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4022` | no |  |

### `flip`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4008` | no |  |
| `direction` | `MsoFlipCmd` | no |  |

### `get action setting for`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4010` | no |  |
| `event` | `EPPMouseActivation` | no |  |

**Returns:** `action setting`

### `get custom color`

Returns the custom color for the specified Microsoft Office theme.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme color scheme` | no |  |
| `name` | `text` | no |  |

**Returns:** `integer`

### `group`

Groups the shapes in the specified shape range.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape range` | no |  |

**Returns:** `shape` — The grouped shapes as a single shape object.

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

Loads the color scheme of a Microsoft Office theme from a file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `theme color scheme` | no |  |
| `file name` | `text` | no | The name of the color theme file. |

### `load theme effect scheme`

Loads the effects scheme of a Microsoft Office theme from a file.

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

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4011` | no |  |
| `style` | `MsoGradientStyle` | no |  |
| `variant` | `integer` | no |  |
| `degree` | `real` | no |  |

### `patterned`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4012` | no |  |
| `pattern` | `MsoPatternType` | no |  |

### `pick up`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4005` | no |  |

### `preset drop`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `callout format` | no |  |
| `drop type` | `MsoCalloutDropType` | no |  |

### `preset gradient`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4013` | no |  |
| `style` | `MsoGradientStyle` | no |  |
| `variant` | `integer` | no |  |
| `gradient type` | `MsoPresetGradientType` | no |  |

### `preset textured`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4014` | no |  |
| `texture` | `MsoPresetTexture` | no |  |

### `regroup`

Regroups the group that the specified shape range belonged to previously.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape range` | no |  |

**Returns:** `shape` — The regrouped shapes as a single shape object.

### `reroute connections`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4006` | no |  |

### `reset rotation`

Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4023` | no |  |

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

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `picture` | no |  |
| `factor` | `real` | no |  |
| `relative to original size` | `boolean` | no |  |
| `scale` | `MsoScaleFrom` | no |  |

### `scale width`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `picture` | no |  |
| `factor` | `real` | no |  |
| `relative to original size` | `boolean` | no |  |
| `scale` | `MsoScaleFrom` | no |  |

### `set shapes default properties`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4007` | no |  |

### `solid`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4015` | no |  |

### `toggle vertical text`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `word art format` | no |  |

### `two color gradient`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4016` | no |  |
| `style` | `MsoGradientStyle` | no |  |
| `variant` | `integer` | no |  |

### `ungroup`

Ungroups any grouped shapes in the specified shape or range of shapes.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `shape range` | no |  |

**Returns:** `shape range` — The ungrouped shapes as a single shape range object.

### `user picture`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4017` | no |  |
| `picture file` | `text` | no |  |

### `user textured`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4018` | no |  |
| `texture file` | `text` | no |  |

### `z order`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `4009` | no |  |
| `z order position` | `MsoZOrderCmd` | no |  |

### Classes

### `adjustment`

*Plural:* `adjustments`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `adjustment_value` | `real` | r/w |  |

### `callout format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accent` | `boolean` | r/w |  |
| `angle` | `MsoCalloutAngleType` | r/w |  |
| `auto attach` | `boolean` | r/w |  |
| `auto length` | `boolean` | r/o |  |
| `border` | `boolean` | r/w |  |
| `callout format length` | `real` | r/o |  |
| `callout type` | `MsoCalloutType` | r/w |  |
| `drop` | `real` | r/o |  |
| `drop type` | `MsoCalloutDropType` | r/o |  |
| `gap` | `real` | r/w |  |

### `callout`

*Plural:* `callouts`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `callout Type` | `MsoCalloutType` | r/o |  |
| `callout format` | `callout format` | r/o |  |

### `comment`

*Plural:* `comments`

### `connector format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `begin connected` | `boolean` | r/o |  |
| `begin connected shape` | `shape` | r/o |  |
| `begin connection site` | `integer` | r/o |  |
| `connector type` | `MsoConnectorType` | r/w |  |
| `end connected` | `boolean` | r/o |  |
| `end connected shape` | `shape` | r/o |  |
| `end connection site` | `integer` | r/o |  |

### `connector`

*Plural:* `connectors`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `connector format` | `connector format` | r/o |  |
| `connector type` | `MsoConnectorType` | r/o |  |

### `fill format`

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.

**Contains:** `gradient stop`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `back color` | `integer` | r/w |  |
| `back color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the specified fill background color. |
| `fill format type` | `MsoFillType` | r/o |  |
| `fore color` | `integer` | r/w |  |
| `fore color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the specified foreground fill or solid color. |
| `gradient color type` | `MsoGradientColorType` | r/o |  |
| `gradient degree` | `real` | r/o |  |
| `gradient style` | `MsoGradientStyle` | r/o |  |
| `gradient variant` | `integer` | r/o |  |
| `pattern` | `MsoPatternType` | r/o |  |
| `preset gradient type` | `MsoPresetGradientType` | r/o |  |
| `preset texture` | `MsoPresetTexture` | r/o |  |
| `rotate with object` | `boolean` | r/w | Returns or sets whether the fill rotates with the specified shape. |
| `texture alignment` | `MsoTextureAlignment` | r/w | Returns or sets the texture alignment for the specified object. |
| `texture horizontal scale` | `real` | r/w | Returns or sets the texture alignment for the specified object. |
| `texture name` | `text` | r/o |  |
| `texture offset X` | `real` | r/w | Returns or sets the texture alignment for the specified object. |
| `texture offset Y` | `real` | r/w | Returns or sets the texture alignment for the specified object. |
| `texture tile` | `boolean` | r/w | Returns the texture tile style for the specified fill. |
| `texture vertical scale` | `real` | r/w | Returns or sets the texture alignment for the specified object. |
| `transparency` | `real` | r/w |  |
| `visible` | `boolean` | r/w |  |

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

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `back color` | `integer` | r/w |  |
| `back color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the background color for a patterned line. |
| `begin arrow head length` | `MsoArrowheadLength` | r/w |  |
| `begin arrowhead style` | `MsoArrowheadStyle` | r/w |  |
| `begin arrowhead width` | `MsoArrowheadWidth` | r/w |  |
| `dash style` | `MsoLineDashStyle` | r/w |  |
| `end arrowhead length` | `MsoArrowheadLength` | r/w |  |
| `end arrowhead style` | `MsoArrowheadStyle` | r/w |  |
| `end arrowhead width` | `MsoArrowheadWidth` | r/w |  |
| `fore color` | `integer` | r/w |  |
| `fore color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the foreground color for the line. |
| `line format patterned` | `MsoPatternType` | r/w |  |
| `line style` | `MsoLineStyle` | r/w |  |
| `line weight` | `real` | r/w |  |
| `transparency` | `real` | r/w |  |

### `line shape`

The line shape uses begin line X, begin line Y, end line X, and end line Y when created

*Plural:* `line shapes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `begin line X` | `real` | r/w |  |
| `begin line Y` | `real` | r/w |  |
| `end line X` | `real` | r/w |  |
| `end line Y` | `real` | r/w |  |

### `link format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto update` | `EPPUpdateOption` | r/w |  |
| `source full name` | `text` | r/w |  |

### `major theme font`

Represents a container for the font schemes of a Microsoft Office theme.

*Plural:* `major theme fonts`

### `media object`

*Plural:* `media objects`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `text` | r/o |  |

### `media2 object`

*Plural:* `media2 objects`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `text` | r/o |  |
| `link to file` | `boolean` | r/o |  |
| `save with document` | `boolean` | r/o |  |

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

### `picture format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `brightness` | `real` | r/w | Returns or sets the brightness of the specified picture. The value for this property must be a number from 0.0, dimmest to 1.0, brightest. |
| `color type` | `MsoPictureColorType` | r/w | Returns or sets the type of color transformation applied to the specified picture. |
| `contrast` | `real` | r/w | Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast. |
| `crop bottom` | `real` | r/w | Returns or sets the number of points that are cropped off the bottom of the specified picture. |
| `crop left` | `real` | r/w | Returns or sets the number of points that are cropped off the left side of the specified picture. |
| `crop right` | `real` | r/w | Returns or sets the number of points that are cropped off the right side of the specified picture. |
| `crop top` | `real` | r/w | Returns or sets the number of points that are cropped off the top of the specified picture. |
| `transparency color` | `integer` | r/w | Returns or sets the transparent color for the specified picture as aRGB color. For this property to take effect, the transparent background property must be set to true. |
| `transparent background` | `boolean` | r/w | Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent. |

### `picture`

*Plural:* `pictures`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `text` | r/o |  |
| `link to file` | `boolean` | r/o |  |
| `picture format` | `picture format` | r/o |  |
| `save with document` | `boolean` | r/o |  |

### `place holder`

*Plural:* `place holders`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `place holder format` | `placeholder format` | r/o |  |
| `placeholder type` | `EPPPlaceholderType` | r/o |  |

### `placeholder format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `contained type` | `MsoShapeType` | r/o |  |
| `placeholder name` | `text` | r/w |  |
| `placeholder type` | `EPPPlaceholderType` | r/o |  |

### `reflection format`

Represents the reflection effect in Office graphics.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `reflection type` | `MsoReflectionType` | r/w | Returns or sets the type of the reflection format object. |

### `shadow format`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `blur` | `real` | r/w |  |
| `fore color` | `integer` | r/w |  |
| `fore color theme index` | `MsoThemeColorIndex` | r/w | Returns or sets the foreground color for the shadow format. |
| `obscured` | `boolean` | r/w |  |
| `offset X` | `real` | r/w |  |
| `offset Y` | `real` | r/w |  |
| `rotate with shape` | `boolean` | r/w | Returns or sets whether to rotate the shadow when rotating the shape. |
| `shadow style` | `MsoShadowStyle` | r/w | Returns or sets the style of shadow formatting to apply to a shape. |
| `shadow type` | `MsoShadowType` | r/w |  |
| `size` | `real` | r/w | Returns or sets the width of the shadow. |
| `transparency` | `real` | r/w |  |
| `visible` | `boolean` | r/w |  |

### `shape range`

A collection that represents a set of shapes on a document.

**Contains:** `adjustment`, `shape`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `animation settings` | `animation settings` | r/o | Returns all the special effects that you can apply to the animation of the specified shape. |
| `auto shape type` | `MsoAutoShapeType` | r/w | Returns or sets the shape type for AutoShapes other than a line, freeform drawing, or connector. |
| `background style` | `MsoBackgroundStyleIndex` | r/w | Returns or sets the background style for the specified object. |
| `black and white mode` | `MsoBlackWhiteMode` | r/w | Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. |
| `callout format` | `callout format` | r/o | Returns callout formatting properties for the specified line callout shapes. |
| `child` | `boolean` | r/o | Returns whether all shapes in a shape range are child shapes of the same parent. |
| `connection site count` | `integer` | r/o | Returns the number of connection sites on the specified shape. |
| `fill format` | `fill format` | r/o | Returns the fill format properties for the specified object. |
| `glow format` | `glow format` | r/o | Returns the glow format for the specified range of shapes. |
| `has child` | `boolean` | r/o |  |
| `has table` | `boolean` | r/o | Returns whether the specified shape is a table. |
| `has text frame` | `boolean` | r/o | Returns whether the specified shape has a text frame. |
| `height` | `real` | r/w | Returns or sets the height of the specified object. |
| `horizontal flip` | `boolean` | r/o | Returns whether the specified shape is flipped around the horizontal axis. |
| `is connector` | `boolean` | r/o | Determines whether the specified shape is a connector. |
| `left position` | `real` | r/w | Returns or sets a value equal to the distance from the left edge of the leftmost shape in the shape range to the left edge of the slide. |
| `line format` | `line format` | r/o | Returns line format properties for the specified line or shape border. |
| `link format` | `link format` | r/o | Returns the properties for linked OLE objects. |
| `lock aspect ratio` | `boolean` | r/w | Determines whether the specified shape retains its original proportions when you resize it. |
| `media type` | `EPPMediaType` | r/o | Returns the OLE media type. |
| `name` | `text` | r/w | Identifies the shape on a slide. |
| `reflection format` | `reflection format` | r/o | Returns the reflection format for the specified range of shapes. |
| `rotation` | `real` | r/w | Returns or sets the number of degrees that the specified shape is rotated around the z-axis. |
| `shadow format` | `shadow format` | r/o | Returns shadow formatting for specified shapes. |
| `shape style` | `MsoShapeStyleIndex` | r/w | Returns or sets the shape style index for the specified object. |
| `shape type` | `MsoShapeType` | r/o | Represents the type of shape or shapes in a range of shapes. |
| `soft edge format` | `soft edge format` | r/o | Returns the soft edge format for the specified range of shapes. |
| `text frame` | `text frame` | r/o | Returns the alignment and anchoring properties for the specified shape or master text style. |
| `threeD format` | `threeD format` | r/o | Returns 3-D effect formatting properties for the specified shape. |
| `top` | `real` | r/w | Returns or sets a value equal to the distance from the top edge of the topmost shape in the shape range to the top edge of the document. |
| `vertical flip` | `boolean` | r/o | Determines whether the specified shape is flipped around the vertical axis. |
| `visible` | `boolean` | r/w | Returns or sets the visibility of the specified object or the formatting applied to the specified object. |
| `width` | `real` | r/w | Returns or sets the width of the specified object. |
| `z order position` | `integer` | r/o | Returns the position of the specified shape in the z-order. |

### `shape table`

*Plural:* `shape tables`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `number of columns` | `integer` | r/o |  |
| `number of rows` | `integer` | r/o |  |
| `table object` | `table` | r/o |  |

### `shape`

*Plural:* `shapes`

**Contains:** `shape`, `callout`, `connector`, `picture`, `line shape`, `place holder`, `text box`, `comment`, `shape table`, `adjustment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `animation settings` | `animation settings` | r/o |  |
| `auto shape type` | `MsoAutoShapeType` | r/w |  |
| `background style` | `MsoBackgroundStyleIndex` | r/w | Returns or sets the background style. |
| `black and white mode` | `MsoBlackWhiteMode` | r/w |  |
| `child` | `boolean` | r/o | True if the shape is a child shape. |
| `connection site count` | `integer` | r/o |  |
| `fill format` | `fill format` | r/o | Returns a fill format object that contains fill formatting properties for the specified shape. |
| `glow format` | `glow format` | r/o | Returns the formatting properties for a glow effect. |
| `has table` | `boolean` | r/o |  |
| `has text frame` | `boolean` | r/o |  |
| `height` | `real` | r/w |  |
| `horizontal flip` | `boolean` | r/o |  |
| `is connector` | `boolean` | r/o |  |
| `left position` | `real` | r/w |  |
| `line format` | `line format` | r/o |  |
| `link format` | `link format` | r/o |  |
| `lock aspect ratio` | `boolean` | r/w |  |
| `media type` | `EPPMediaType` | r/o |  |
| `name` | `text` | r/w |  |
| `reflection format` | `reflection format` | r/o | Returns the formatting properties for a reflection effect. |
| `rotation` | `real` | r/w |  |
| `shadow format` | `shadow format` | r/o |  |
| `shape style` | `MsoShapeStyleIndex` | r/w | Returns or sets the shape style corresponding to the Quick Styles. |
| `shape type` | `MsoShapeType` | r/o |  |
| `soft edge format` | `soft edge format` | r/o | Returns the formatting properties for a soft edge effect. |
| `text frame` | `text frame` | r/o |  |
| `threeD format` | `threeD format` | r/o | Returns a threeD format object that contains 3-D-effect formatting properties for the specified shape. |
| `top` | `real` | r/w |  |
| `vertical flip` | `boolean` | r/o |  |
| `visible` | `boolean` | r/w |  |
| `width` | `real` | r/w |  |
| `word art format` | `word art format` | r/o | Returns the formatting properties for a word art effect. |
| `z order position` | `integer` | r/o |  |

### `soft edge format`

Represents the soft edge formatting for a shape or range of shapes

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `soft edge type` | `MsoSoftEdgeType` | r/w | Returns or sets the type soft edge format object. |

### `text box`

*Plural:* `text boxes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `text orientation` | `MsoTextOrientation` | r/o |  |

### `text column`

Represents a single text column.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `column number` | `integer` | r/w | Returns or sets the index of the text column object. |
| `spacing` | `real` | r/w | Returns or sets the spacing between text columns in a text column object. |
| `text direction` | `MsoTextDirection` | r/w | Returns or sets the direction of text in the text column object. |

### `text frame`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `auto size` | `MsoAutoSize` | r/w |  |
| `has text` | `boolean` | r/o | Returns whether the specified text frame has text. |
| `horizontal anchor` | `MsoHorizontalAnchor` | r/w |  |
| `margin bottom` | `real` | r/w |  |
| `margin left` | `real` | r/w |  |
| `margin right` | `real` | r/w |  |
| `margin top` | `real` | r/w |  |
| `orientation` | `MsoTextOrientation` | r/o |  |
| `path format` | `MsoPathFormat` | r/w | Returns or sets the path type for the specified text frame. |
| `ruler` | `ruler` | r/o |  |
| `text column` | `text column` | r/o | Returns the text column object that represents the columns within the text frame. |
| `text orientation` | `MsoTextOrientation` | r/w |  |
| `text range` | `text range` | r/o |  |
| `threeD format` | `threeD format` | r/o | Returns the 3-D effect formatting properties for the specified text. |
| `vertical anchor` | `MsoVerticalAnchor` | r/w |  |
| `warp format` | `MsoWarpFormat` | r/w | Returns or sets the warp type for the specified text frame. |
| `word art styles format` | `MsoPresetTextEffect` | r/o | Returns or sets a value specifying the text effect for the selected text. |
| `word wrap` | `boolean` | r/w | Returns or sets text break lines within or past the boundaries of the shape. |
| `wordart format` | `MsoPresetTextEffect` | r/w | Returns or sets the WordArt type for the specified text frame. |

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

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `font bold` | `boolean` | r/w |  |
| `font italic` | `boolean` | r/w |  |
| `font name` | `text` | r/w |  |
| `font size` | `real` | r/w |  |
| `kerned pairs` | `boolean` | r/w |  |
| `normalized height` | `boolean` | r/w |  |
| `preset shape` | `MsoPresetTextEffectShape` | r/w |  |
| `rotated chars` | `boolean` | r/w |  |
| `text alignment` | `MsoTextEffectAlignment` | r/w |  |
| `tracking` | `real` | r/w |  |
| `word art styles format` | `MsoPresetTextEffect` | r/w | Returns or sets a value specifying the text effect for the selected text. |
| `word art text` | `text` | r/w |  |

### Enumerations

### `one color gradient`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `two color gradient`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `automatic length`

| Value | Description |
|-------|-------------|
| `callout` |  |
| `callout format` |  |

### `begin connect`

| Value | Description |
|-------|-------------|
| `connector` |  |
| `connector format` |  |

### `begin disconnect`

| Value | Description |
|-------|-------------|
| `connector` |  |
| `connector format` |  |

### `custom length`

| Value | Description |
|-------|-------------|
| `callout` |  |
| `callout format` |  |

### `custom drop`

| Value | Description |
|-------|-------------|
| `callout` |  |
| `callout format` |  |

### `end connect`

| Value | Description |
|-------|-------------|
| `connector` |  |
| `connector format` |  |

### `end disconnect`

| Value | Description |
|-------|-------------|
| `connector` |  |
| `connector format` |  |

### `patterned`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `get action setting for`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `shape range` |  |

### `solid`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `reset rotation`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `threeD format` |  |

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

### `z order`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `shape range` |  |

### `preset textured`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `preset gradient`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `fill format` |  |

### `apply`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `shape range` |  |

### `flip`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `shape range` |  |

### `pick up`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `shape range` |  |

### `reroute connections`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `shape range` |  |

### `set shapes default properties`

| Value | Description |
|-------|-------------|
| `shape` |  |
| `shape range` |  |
| `shape range` |  |

---

## Text Suite

Classes and types for scripting text manipulations

### Commands

### `add periods to`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `change case`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `to` | `MsoTextChangeCase` | no |  |

### `copy text range`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `cut text range`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `get rotated text bounds`

Returns back a list containing the four point of the text bounds as follows  {x1, y1}, {x2, y2}, {x3, y3}, {x4, y4} }

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

**Returns:** `real`

### `get text action setting`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `result` | `EPPMouseActivation` | no |  |

**Returns:** `action setting`

### `insert text text range`

Adds text to existing text.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |
| `insert where` | `MsoTextRangeInsertPosition` | no | Determines where text will be inerted. |
| `new text` | `text` | no | Contains the text to be inserted. |

### `paste text range`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### `remove periods from`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text range` | no |  |

### Classes

### `character`

*Plural:* `characters`

### `line`

*Plural:* `lines`

### `paragraph`

*Plural:* `paragraphs`

### `sentence`

*Plural:* `sentences`

### `text flow`

*Plural:* `text flows`

### `text range`

**Contains:** `character`, `word`, `sentence`, `line`, `paragraph`, `text flow`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bounds height` | `real` | r/o |  |
| `bounds width` | `real` | r/o |  |
| `content` | `text` | r/w |  |
| `font` | `font` | r/o |  |
| `indent level` | `integer` | r/w |  |
| `left bounds` | `real` | r/o |  |
| `offset` | `integer` | r/o |  |
| `paragraph format` | `paragraph format` | r/o |  |
| `text length` | `integer` | r/o |  |
| `top bounds` | `real` | r/o |  |

### `word`

*Plural:* `words`

---

## Table Suite

Classes and types for scripting table manipulations

### Commands

### `get border`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cell` | no |  |
| `edge` | `PpBorderType` | no |  |

**Returns:** `line format`

### `get cell from`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `table` | no |  |
| `row` | `integer` | no |  |
| `column` | `integer` | no |  |

**Returns:** `cell`

### `merge`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cell` | no |  |
| `merge with` | `cell` | no |  |

### `split`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `cell` | no |  |
| `number of rows` | `integer` | no |  |
| `number of columns` | `integer` | no |  |

### Classes

### `cell`

*Plural:* `cells`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `selected` | `boolean` | r/o |  |
| `shape` | `shape` | r/o |  |

### `column`

*Plural:* `columns`

**Contains:** `cell`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `width` | `real` | r/w |  |

### `row`

*Plural:* `rows`

**Contains:** `cell`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `height` | `real` | r/w |  |

### `table`

*Plural:* `tables`

**Contains:** `column`, `row`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `table direction` | `MsoTextDirection` | r/w |  |
