-- title: Search Music library
-- verified: no
-- macos: 13+
-- app: Music
-- description: Search for tracks by name in the Music library
-- source: app dictionary

set searchTerm to "Bohemian"

tell application "Music"
	set libraryPlaylist to first library playlist
	set results to search libraryPlaylist for searchTerm only songs

	set trackList to {}
	repeat with t in results
		set end of trackList to {name of t, artist of t, album of t}
	end repeat
	return trackList
end tell
