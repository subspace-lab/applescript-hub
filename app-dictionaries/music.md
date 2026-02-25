# Com.apple.music AppleScript Dictionary

> Auto-generated from `com.apple.Music.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Com.apple.music"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Music Suite](#music-suite)
- [Internet suite](#internet-suite)

---

## Standard Suite

Common terms for most applications

### Commands

### `print`

Print the specified object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | yes | list of objects to print |
| `print dialog` | `boolean` | yes | Should the application show the print dialog |
| `with properties` | `print settings` | yes | the print settings |
| `kind` | `eKnd` | yes | the kind of printout desired |
| `theme` | `text` | yes | name of theme to use for formatting the printout |

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

### Enumerations

### `eKnd`

| Value | Description |
|-------|-------------|
| `track listing` | a basic listing of tracks within a playlist |
| `album listing` | a listing of a playlist grouped by album |
| `cd insert` | a printout of the playlist for jewel case inserts |

### `enum`

| Value | Description |
|-------|-------------|
| `standard` | Standard PostScript error handling |
| `detailed` | print a detailed report of PostScript errors |

---

## Music Suite

The event suite specific to Music

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

### `export`

export a source or playlist

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `item` | no | the source or playlist to export |
| `as` | `eExF` | yes | the format to export for a playlist |
| `to` | `file` | yes | the destination of the export |

**Returns:** `text` — the exported data for the source or playlist, if not written to a file

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

### `AirPlay device`

an AirPlay device

*Plural:* `AirPlay devices`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active` | `boolean` | r/o | is the device currently being played to? |
| `available` | `boolean` | r/o | is the device currently available? |
| `kind` | `eAPD` | r/o | the kind of the device |
| `network address` | `text` | r/o | the network (MAC) address of the device |
| `protected` | `boolean` | r/o | is the device password- or passcode-protected? |
| `selected` | `boolean` | r/w | is the device currently selected? |
| `supports audio` | `boolean` | r/o | does the device support audio playback? |
| `supports video` | `boolean` | r/o | does the device support video playback? |
| `sound volume` | `integer` | r/w | the output volume for the device (0 = minimum, 100 = maximum) |

### `application`

The application program

**Contains:** `AirPlay device`, `browser window`, `encoder`, `EQ preset`, `EQ window`, `miniplayer window`, `playlist`, `playlist window`, `source`, `track`, `video window`, `visual`, `window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `AirPlay enabled` | `boolean` | r/o | is AirPlay currently enabled? |
| `converting` | `boolean` | r/o | is a track currently being converted? |
| `current AirPlay devices` | `AirPlay device` | r/w | the currently selected AirPlay device(s) |
| `current encoder` | `encoder` | r/w | the currently selected encoder (MP3, AIFF, WAV, etc.) |
| `current EQ preset` | `EQ preset` | r/w | the currently selected equalizer preset |
| `current playlist` | `playlist` | r/o | the playlist containing the currently targeted track |
| `current stream title` | `text` | r/o | the name of the current track in the playing stream (provided by streaming server) |
| `current stream URL` | `text` | r/o | the URL of the playing stream or streaming web site (provided by streaming server) |
| `current track` | `track` | r/o | the current targeted track |
| `current visual` | `visual` | r/w | the currently selected visual plug-in |
| `EQ enabled` | `boolean` | r/w | is the equalizer enabled? |
| `fixed indexing` | `boolean` | r/w | true if all AppleScript track indices should be independent of the play order of the owning playlist. |
| `frontmost` | `boolean` | r/w | is this the active application? |
| `full screen` | `boolean` | r/w | is the application using the entire screen? |
| `name` | `text` | r/o | the name of the application |
| `mute` | `boolean` | r/w | has the sound output been muted? |
| `player position` | `real` | r/w | the player’s position within the currently playing track in seconds. |
| `player state` | `ePlS` | r/o | is the player stopped, paused, or playing? |
| `selection` | `specifier` | r/o | the selection visible to the user |
| `shuffle enabled` | `boolean` | r/w | are songs played in random order? |
| `shuffle mode` | `eShM` | r/w | the playback shuffle mode |
| `song repeat` | `eRpt` | r/w | the playback repeat mode |
| `sound volume` | `integer` | r/w | the sound output volume (0 = minimum, 100 = maximum) |
| `version` | `text` | r/o | the version of the application |
| `visuals enabled` | `boolean` | r/w | are visuals currently being displayed? |

### `artwork`

a piece of art within a track or playlist

*Plural:* `artworks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `data` | `picture` | r/w | data for this artwork, in the form of a picture |
| `description` | `text` | r/w | description of artwork as a string |
| `downloaded` | `boolean` | r/o | was this artwork downloaded by Music? |
| `format` | `type` | r/o | the data format for this piece of artwork |
| `kind` | `integer` | r/w | kind or purpose of this piece of artwork |
| `raw data` | `any` | r/w | data for this artwork, in original format |

### `audio CD playlist`

a playlist representing an audio CD

*Plural:* `audio CD playlists`

**Contains:** `audio CD track`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `artist` | `text` | r/w | the artist of the CD |
| `compilation` | `boolean` | r/w | is this CD a compilation album? |
| `composer` | `text` | r/w | the composer of the CD |
| `disc count` | `integer` | r/w | the total number of discs in this CD’s album |
| `disc number` | `integer` | r/w | the index of this CD disc in the source album |
| `genre` | `text` | r/w | the genre of the CD |
| `year` | `integer` | r/w | the year the album was recorded/released |

### `audio CD track`

a track on an audio CD

*Plural:* `audio CD tracks`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `location` | `file` | r/o | the location of the file represented by this track |

### `browser window`

the main window

*Plural:* `browser windows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `selection` | `specifier` | r/o | the selected tracks |
| `view` | `playlist` | r/w | the playlist currently displayed in the window |

### `encoder`

converts a track to a specific file format

*Plural:* `encoders`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `format` | `text` | r/o | the data format created by the encoder |

### `EQ preset`

equalizer preset configuration

*Plural:* `EQ presets`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `band 1` | `real` | r/w | the equalizer 32 Hz band level (-12.0 dB to +12.0 dB) |
| `band 2` | `real` | r/w | the equalizer 64 Hz band level (-12.0 dB to +12.0 dB) |
| `band 3` | `real` | r/w | the equalizer 125 Hz band level (-12.0 dB to +12.0 dB) |
| `band 4` | `real` | r/w | the equalizer 250 Hz band level (-12.0 dB to +12.0 dB) |
| `band 5` | `real` | r/w | the equalizer 500 Hz band level (-12.0 dB to +12.0 dB) |
| `band 6` | `real` | r/w | the equalizer 1 kHz band level (-12.0 dB to +12.0 dB) |
| `band 7` | `real` | r/w | the equalizer 2 kHz band level (-12.0 dB to +12.0 dB) |
| `band 8` | `real` | r/w | the equalizer 4 kHz band level (-12.0 dB to +12.0 dB) |
| `band 9` | `real` | r/w | the equalizer 8 kHz band level (-12.0 dB to +12.0 dB) |
| `band 10` | `real` | r/w | the equalizer 16 kHz band level (-12.0 dB to +12.0 dB) |
| `modifiable` | `boolean` | r/o | can this preset be modified? |
| `preamp` | `real` | r/w | the equalizer preamp level (-12.0 dB to +12.0 dB) |
| `update tracks` | `boolean` | r/w | should tracks which refer to this preset be updated when the preset is renamed or deleted? |

### `EQ window`

the equalizer window

*Plural:* `EQ windows`

### `file track`

a track representing an audio file (MP3, AIFF, etc.)

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

### `miniplayer window`

the miniplayer window

*Plural:* `miniplayer windows`

### `playlist`

a list of tracks/streams

*Plural:* `playlists`

**Contains:** `track`, `artwork`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `description` | `text` | r/w | the description of the playlist |
| `disliked` | `boolean` | r/w | is this playlist disliked? |
| `duration` | `integer` | r/o | the total length of all tracks (in seconds) |
| `name` | `text` | r/w | the name of the playlist |
| `favorited` | `boolean` | r/w | is this playlist favorited? |
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

### `radio tuner playlist`

the radio tuner playlist

*Plural:* `radio tuner playlists`

**Contains:** `URL track`

### `shared track`

a track residing in a shared library

*Plural:* `shared tracks`

### `source`

a media source (library, CD, device, etc.)

*Plural:* `sources`

**Contains:** `audio CD playlist`, `library playlist`, `playlist`, `radio tuner playlist`, `subscription playlist`, `user playlist`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `capacity` | `double integer` | r/o | the total size of the source if it has a fixed size |
| `free space` | `double integer` | r/o | the free space on the source if it has a fixed size |
| `kind` | `eSrc` | r/o |  |

### `subscription playlist`

a subscription playlist from Apple Music

*Plural:* `subscription playlists`

**Contains:** `file track`, `URL track`

### `track`

playable audio source

*Plural:* `tracks`

**Contains:** `artwork`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `album` | `text` | r/w | the album name of the track |
| `album artist` | `text` | r/w | the album artist of the track |
| `album disliked` | `boolean` | r/w | is the album for this track disliked? |
| `album favorited` | `boolean` | r/w | is the album for this track favorited? |
| `album rating` | `integer` | r/w | the rating of the album for this track (0 to 100) |
| `album rating kind` | `eRtK` | r/o | the rating kind of the album rating for this track |
| `artist` | `text` | r/w | the artist/source of the track |
| `bit rate` | `integer` | r/o | the bit rate of the track (in kbps) |
| `bookmark` | `real` | r/w | the bookmark time of the track in seconds |
| `bookmarkable` | `boolean` | r/w | is the playback position for this track remembered? |
| `bpm` | `integer` | r/w | the tempo of this track in beats per minute |
| `category` | `text` | r/w | the category of the track |
| `cloud status` | `eClS` | r/o | the iCloud status of the track |
| `comment` | `text` | r/w | freeform notes about the track |
| `compilation` | `boolean` | r/w | is this track from a compilation album? |
| `composer` | `text` | r/w | the composer of the track |
| `database ID` | `integer` | r/o | the common, unique ID for this track. If two tracks in different playlists have the same database ID, they are sharing the same data. |
| `date added` | `date` | r/o | the date the track was added to the playlist |
| `description` | `text` | r/w | the description of the track |
| `disc count` | `integer` | r/w | the total number of discs in the source album |
| `disc number` | `integer` | r/w | the index of the disc containing this track on the source album |
| `disliked` | `boolean` | r/w | is this track disliked? |
| `downloader account` | `text` | r/o | the account of the person who downloaded this track |
| `downloader name` | `text` | r/o | the name of the person who downloaded this track |
| `duration` | `real` | r/o | the length of the track in seconds |
| `enabled` | `boolean` | r/w | is this track checked for playback? |
| `episode ID` | `text` | r/w | the episode ID of the track |
| `episode number` | `integer` | r/w | the episode number of the track |
| `EQ` | `text` | r/w | the name of the EQ preset of the track |
| `finish` | `real` | r/w | the stop time of the track in seconds |
| `gapless` | `boolean` | r/w | is this track from a gapless album? |
| `genre` | `text` | r/w | the music/audio genre (category) of the track |
| `grouping` | `text` | r/w | the grouping (piece) of the track. Generally used to denote movements within a classical work. |
| `kind` | `text` | r/o | a text description of the track |
| `long description` | `text` | r/w | the long description of the track |
| `favorited` | `boolean` | r/w | is this track favorited? |
| `lyrics` | `text` | r/w | the lyrics of the track |
| `media kind` | `eMdK` | r/w | the media kind of the track |
| `modification date` | `date` | r/o | the modification date of the content of this track |
| `movement` | `text` | r/w | the movement name of the track |
| `movement count` | `integer` | r/w | the total number of movements in the work |
| `movement number` | `integer` | r/w | the index of the movement in the work |
| `played count` | `integer` | r/w | number of times this track has been played |
| `played date` | `date` | r/w | the date and time this track was last played |
| `purchaser account` | `text` | r/o | the account of the person who purchased this track |
| `purchaser name` | `text` | r/o | the name of the person who purchased this track |
| `rating` | `integer` | r/w | the rating of this track (0 to 100) |
| `rating kind` | `eRtK` | r/o | the rating kind of this track |
| `release date` | `date` | r/o | the release date of this track |
| `sample rate` | `integer` | r/o | the sample rate of the track (in Hz) |
| `season number` | `integer` | r/w | the season number of the track |
| `shufflable` | `boolean` | r/w | is this track included when shuffling? |
| `skipped count` | `integer` | r/w | number of times this track has been skipped |
| `skipped date` | `date` | r/w | the date and time this track was last skipped |
| `show` | `text` | r/w | the show name of the track |
| `sort album` | `text` | r/w | override string to use for the track when sorting by album |
| `sort artist` | `text` | r/w | override string to use for the track when sorting by artist |
| `sort album artist` | `text` | r/w | override string to use for the track when sorting by album artist |
| `sort name` | `text` | r/w | override string to use for the track when sorting by name |
| `sort composer` | `text` | r/w | override string to use for the track when sorting by composer |
| `sort show` | `text` | r/w | override string to use for the track when sorting by show name |
| `size` | `double integer` | r/o | the size of the track (in bytes) |
| `start` | `real` | r/w | the start time of the track in seconds |
| `time` | `text` | r/o | the length of the track in MM:SS format |
| `track count` | `integer` | r/w | the total number of tracks on the source album |
| `track number` | `integer` | r/w | the index of the track on the source album |
| `unplayed` | `boolean` | r/w | is this track unplayed? |
| `volume adjustment` | `integer` | r/w | relative volume adjustment of the track (-100% to 100%) |
| `work` | `text` | r/w | the work name of the track |
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
| `genius` | `boolean` | r/o | is this a Genius Playlist? |

### `video window`

the video window

*Plural:* `video windows`

### `visual`

a visual plug-in

*Plural:* `visuals`

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

### `eRpt`

| Value | Description |
|-------|-------------|
| `off` |  |
| `one` |  |
| `all` |  |

### `eShM`

| Value | Description |
|-------|-------------|
| `songs` |  |
| `albums` |  |
| `groupings` |  |

### `eSrc`

| Value | Description |
|-------|-------------|
| `library` |  |
| `audio CD` |  |
| `MP3 CD` |  |
| `radio tuner` |  |
| `shared library` |  |
| `iTunes Store` |  |
| `unknown` |  |

### `eSrA`

| Value | Description |
|-------|-------------|
| `albums` | albums only |
| `all` | all text fields |
| `artists` | artists only |
| `composers` | composers only |
| `displayed` | visible text fields |
| `names` | track names only |

### `eSpK`

| Value | Description |
|-------|-------------|
| `none` |  |
| `folder` |  |
| `Genius` |  |
| `Library` |  |
| `Music` |  |
| `Purchased Music` |  |

### `eMdK`

| Value | Description |
|-------|-------------|
| `song` | music track |
| `music video` | music video track |
| `movie` | movie track |
| `TV show` | TV show track |
| `unknown` |  |

### `eRtK`

| Value | Description |
|-------|-------------|
| `user` | user-specified rating |
| `computed` | computed rating |

### `eAPD`

| Value | Description |
|-------|-------------|
| `computer` |  |
| `AirPort Express` |  |
| `Apple TV` |  |
| `AirPlay device` |  |
| `Bluetooth device` |  |
| `HomePod` |  |
| `TV` |  |
| `unknown` |  |

### `eClS`

| Value | Description |
|-------|-------------|
| `unknown` |  |
| `purchased` |  |
| `matched` |  |
| `uploaded` |  |
| `ineligible` |  |
| `removed` |  |
| `error` |  |
| `duplicate` |  |
| `subscription` |  |
| `prerelease` |  |
| `no longer available` |  |
| `not uploaded` |  |

### `eExF`

| Value | Description |
|-------|-------------|
| `plain text` |  |
| `Unicode text` |  |
| `XML` |  |
| `M3U` |  |
| `M3U8` |  |

---

## Internet suite

Standard terms for Internet scripting

### Commands

### `open location`

Opens an iTunes Store or audio stream URL

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | yes | the URL to open |
