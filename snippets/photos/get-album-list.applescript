-- title: List all photo albums
-- verified: no
-- macos: 13+
-- app: Photos
-- description: List all albums in the Photos library with their item counts
-- source: app dictionary

tell application "Photos"
	set albumList to {}
	repeat with a in albums
		set albumName to name of a
		set itemCount to count of media items of a
		set end of albumList to {albumName, itemCount}
	end repeat
	return albumList
end tell
