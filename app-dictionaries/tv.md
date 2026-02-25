# Com.apple.tv AppleScript Dictionary

> Auto-generated from `com.apple.TV.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Com.apple.tv"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [TV Suite](#tv-suite)
- [Internet suite](#internet-suite)

---

## Standard Suite

Common terms for most applications

### Commands

### `close`

Close an object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object to close |

### `count`

Return the number of elements of a particular class within an object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object whose elements are to be counted |
| `each` | `type` | no | the class of the elements to be counted. Keyword 'each' is optional in AppleScript |

**Returns:** `integer` — the number of elements

### `delete`

Delete an element from an object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the element to delete |

### `duplicate`

Duplicate one or more object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object(s) to duplicate |
| `to` | `location specifier` | yes | the new location for the object(s) |

**Returns:** `specifier` — to the duplicated object(s)

### `exists`

Verify if an object exists

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object in question |

**Returns:** `boolean` — true if it exists, false if not

### `make`

Make a new element

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `new` | `type` | no | the class of the new element. Keyword 'new' is optional in AppleScript |
| `at` | `location specifier` | yes | the location at which to insert the element |
| `with properties` | `record` | yes | the initial values for the properties of the element |

**Returns:** `specifier` — to the new object(s)

### `move`

Move playlist(s) to a new location

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `playlist` | no | the playlist(s) to move |
| `to` | `location specifier` | no | the new location for the playlist(s) |

### `open`

Open the specified object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | list of objects to open |

### `run`

Run the application

### `quit`

Quit the application

### `save`

Save the specified object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object(s) to save |

---

## TV Suite

The event suite specific to TV

### Commands

### `add`

add one or more files to a playlist

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file` | no | the file(s) to add |
| `to` | `location specifier` | yes | the location of the added file(s) |

**Returns:** `track` — reference to added track(s)

### `back track`

reposition to beginning of current track or go to previous track if already at start of current track

### `convert`

convert one or more files or tracks

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the file(s)/tracks(s) to convert |

**Returns:** `track` — reference to converted track(s)

### `download`

download a cloud track or playlist

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `item` | no | the shared track, URL track or playlist to download |

### `fast forward`

skip forward in a playing track

### `next track`

advance to the next track in the current playlist

### `pause`

pause playback

### `play`

play the current track or the specified track or file.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | yes | item to play |
| `once` | `boolean` | yes | If true, play this track once and then stop. |

### `playpause`

toggle the playing/paused state of the current track

### `previous track`

return to the previous track in the current playlist

### `refresh`

update file track information from the current information in the track’s file

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file track` | no | the file track to update |

### `resume`

disable fast forward/rewind and resume playback, if playing.

### `reveal`

reveal and select a track or playlist

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `item` | no | the item to reveal |

### `rewind`

skip backwards in a playing track

### `search`

search a playlist for tracks matching the search string. Identical to entering search text in the Search field.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `playlist` | no | the playlist to search |
| `for` | `text` | no | the search text |
| `only` | `eSrA` | yes | area to search (default is all) |

**Returns:** `track` — reference to found track(s)

### `select`

select the specified object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object(s) to select |

### `stop`

stop playback

### Classes

### `application`

The application program

**Contains:** `browser window`, `playlist`, `playlist window`, `source`, `track`, `video window`, `window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `current playlist` | `playlist` | r/o | the playlist containing the currently targeted track |
| `current stream title` | `text` | r/o | the name of the current track in the playing stream (provided by streaming server) |
| `current stream URL` | `text` | r/o | the URL of the playing stream or streaming web site (provided by streaming server) |
| `current track` | `track` | r/o | the current targeted track |
| `fixed indexing` | `boolean` | r/w | true if all AppleScript track indices should be independent of the play order of the owning playlist. |
| `frontmost` | `boolean` | r/w | is this the active application? |
| `full screen` | `boolean` | r/w | is the application using the entire screen? |
| `name` | `text` | r/o | the name of the application |
| `mute` | `boolean` | r/w | has the sound output been muted? |
| `player position` | `real` | r/w | the player’s position within the currently playing track in seconds. |
| `player state` | `ePlS` | r/o | is the player stopped, paused, or playing? |
| `selection` | `specifier` | r/o | the selection visible to the user |
| `sound volume` | `integer` | r/w | the sound output volume (0 = minimum, 100 = maximum) |
| `version` | `text` | r/o | the version of the application |

### `artwork`

a piece of art within a track or playlist

*Plural:* `artworks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `data` | `picture` | r/w | data for this artwork, in the form of a picture |
| `description` | `text` | r/w | description of artwork as a string |
| `downloaded` | `boolean` | r/o | was this artwork downloaded by iTunes? |
| `format` | `type` | r/o | the data format for this piece of artwork |
| `kind` | `integer` | r/w | kind or purpose of this piece of artwork |
| `raw data` | `any` | r/w | data for this artwork, in original format |

### `browser window`

the main window

*Plural:* `browser windows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `selection` | `specifier` | r/o | the selected tracks |
| `view` | `playlist` | r/w | the playlist currently displayed in the window |

### `file track`

a track representing a video file

*Plural:* `file tracks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `location` | `file` | r/w | the location of the file represented by this track |

### `folder playlist`

a folder that contains other playlists

*Plural:* `folder playlists`

### `item`

an item

*Plural:* `items`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `class` | `type` | r/o | the class of the item |
| `container` | `specifier` | r/o | the container of the item |
| `id` | `integer` | r/o | the id of the item |
| `index` | `integer` | r/o | the index of the item in internal application order |
| `name` | `text` | r/w | the name of the item |
| `persistent ID` | `text` | r/o | the id of the item as a hexadecimal string. This id does not change over time. |
| `properties` | `record` | r/w | every property of the item |

### `library playlist`

the main library playlist

*Plural:* `library playlists`

**Contains:** `file track`, `URL track`, `shared track`

### `playlist`

a list of tracks/streams

*Plural:* `playlists`

**Contains:** `track`, `artwork`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `description` | `text` | r/w | the description of the playlist |
| `duration` | `integer` | r/o | the total length of all tracks (in seconds) |
| `name` | `text` | r/w | the name of the playlist |
| `parent` | `playlist` | r/o | folder which contains this playlist (if any) |
| `size` | `integer` | r/o | the total size of all tracks (in bytes) |
| `special kind` | `eSpK` | r/o | special playlist kind |
| `time` | `text` | r/o | the length of all tracks in MM:SS format |
| `visible` | `boolean` | r/o | is this playlist visible in the Source list? |

### `playlist window`

a sub-window showing a single playlist

*Plural:* `playlist windows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `selection` | `specifier` | r/o | the selected tracks |
| `view` | `playlist` | r/o | the playlist displayed in the window |

### `shared track`

a track residing in a shared library

*Plural:* `shared tracks`

### `source`

a media source (library, CD, device, etc.)

*Plural:* `sources`

**Contains:** `library playlist`, `playlist`, `user playlist`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `capacity` | `double integer` | r/o | the total size of the source if it has a fixed size |
| `free space` | `double integer` | r/o | the free space on the source if it has a fixed size |
| `kind` | `eSrc` | r/o |  |

### `track`

playable video source

*Plural:* `tracks`

**Contains:** `artwork`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `album` | `text` | r/w | the album name of the track |
| `album rating` | `integer` | r/w | the rating of the album for this track (0 to 100) |
| `album rating kind` | `eRtK` | r/o | the rating kind of the album rating for this track |
| `bit rate` | `integer` | r/o | the bit rate of the track (in kbps) |
| `bookmark` | `real` | r/w | the bookmark time of the track in seconds |
| `bookmarkable` | `boolean` | r/w | is the playback position for this track remembered? |
| `category` | `text` | r/w | the category of the track |
| `comment` | `text` | r/w | freeform notes about the track |
| `database ID` | `integer` | r/o | the common, unique ID for this track. If two tracks in different playlists have the same database ID, they are sharing the same data. |
| `date added` | `date` | r/o | the date the track was added to the playlist |
| `description` | `text` | r/w | the description of the track |
| `director` | `text` | r/w | the artist/source of the track |
| `disc count` | `integer` | r/w | the total number of discs in the source album |
| `disc number` | `integer` | r/w | the index of the disc containing this track on the source album |
| `downloader account` | `text` | r/o | the account of the person who downloaded this track |
| `downloader name` | `text` | r/o | the name of the person who downloaded this track |
| `duration` | `real` | r/o | the length of the track in seconds |
| `enabled` | `boolean` | r/w | is this track checked for playback? |
| `episode ID` | `text` | r/w | the episode ID of the track |
| `episode number` | `integer` | r/w | the episode number of the track |
| `finish` | `real` | r/w | the stop time of the track in seconds |
| `genre` | `text` | r/w | the genre (category) of the track |
| `grouping` | `text` | r/w | the grouping (piece) of the track. Generally used to denote movements within a classical work. |
| `kind` | `text` | r/o | a text description of the track |
| `long description` | `text` | r/w | the long description of the track |
| `media kind` | `eMdK` | r/w | the media kind of the track |
| `modification date` | `date` | r/o | the modification date of the content of this track |
| `played count` | `integer` | r/w | number of times this track has been played |
| `played date` | `date` | r/w | the date and time this track was last played |
| `purchaser account` | `text` | r/o | the account of the person who purchased this track |
| `purchaser name` | `text` | r/o | the name of the person who purchased this track |
| `rating` | `integer` | r/w | the rating of this track (0 to 100) |
| `rating kind` | `eRtK` | r/o | the rating kind of this track |
| `release date` | `date` | r/o | the release date of this track |
| `sample rate` | `integer` | r/o | the sample rate of the track (in Hz) |
| `season number` | `integer` | r/w | the season number of the track |
| `skipped count` | `integer` | r/w | number of times this track has been skipped |
| `skipped date` | `date` | r/w | the date and time this track was last skipped |
| `show` | `text` | r/w | the show name of the track |
| `sort album` | `text` | r/w | override string to use for the track when sorting by album |
| `sort director` | `text` | r/w | override string to use for the track when sorting by artist |
| `sort name` | `text` | r/w | override string to use for the track when sorting by name |
| `sort show` | `text` | r/w | override string to use for the track when sorting by show name |
| `size` | `double integer` | r/o | the size of the track (in bytes) |
| `start` | `real` | r/w | the start time of the track in seconds |
| `time` | `text` | r/o | the length of the track in MM:SS format |
| `track count` | `integer` | r/w | the total number of tracks on the source album |
| `track number` | `integer` | r/w | the index of the track on the source album |
| `unplayed` | `boolean` | r/w | is this track unplayed? |
| `volume adjustment` | `integer` | r/w | relative volume adjustment of the track (-100% to 100%) |
| `year` | `integer` | r/w | the year the track was recorded/released |

### `URL track`

a track representing a network stream

*Plural:* `URL tracks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `address` | `text` | r/w | the URL for this track |

### `user playlist`

custom playlists created by the user

*Plural:* `user playlists`

**Contains:** `file track`, `URL track`, `shared track`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `shared` | `boolean` | r/w | is this playlist shared? |
| `smart` | `boolean` | r/o | is this a Smart Playlist? |

### `video window`

the video window

*Plural:* `video windows`

### `window`

any window

*Plural:* `windows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bounds` | `rectangle` | r/w | the boundary rectangle for the window |
| `closeable` | `boolean` | r/o | does the window have a close button? |
| `collapseable` | `boolean` | r/o | does the window have a collapse button? |
| `collapsed` | `boolean` | r/w | is the window collapsed? |
| `full screen` | `boolean` | r/w | is the window full screen? |
| `position` | `point` | r/w | the upper left position of the window |
| `resizable` | `boolean` | r/o | is the window resizable? |
| `visible` | `boolean` | r/w | is the window visible? |
| `zoomable` | `boolean` | r/o | is the window zoomable? |
| `zoomed` | `boolean` | r/w | is the window zoomed? |

### Enumerations

### `ePlS`

| Value | Description |
|-------|-------------|
| `stopped` |  |
| `playing` |  |
| `paused` |  |
| `fast forwarding` |  |
| `rewinding` |  |

### `eSrc`

| Value | Description |
|-------|-------------|
| `library` |  |
| `shared library` |  |
| `iTunes Store` |  |
| `unknown` |  |

### `eSrA`

| Value | Description |
|-------|-------------|
| `albums` | albums only |
| `all` | all text fields |
| `artists` | artists only |
| `displayed` | visible text fields |
| `names` | track names only |

### `eSpK`

| Value | Description |
|-------|-------------|
| `none` |  |
| `folder` |  |
| `Library` |  |
| `Movies` |  |
| `TV Shows` |  |

### `eMdK`

| Value | Description |
|-------|-------------|
| `home video` | home video track |
| `movie` | movie track |
| `TV show` | TV show track |
| `unknown` |  |

### `eRtK`

| Value | Description |
|-------|-------------|
| `user` | user-specified rating |
| `computed` | computed rating |

---

## Internet suite

Standard terms for Internet scripting

### Commands

### `open location`

Opens an iTunes Store or stream URL

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | yes | the URL to open |
