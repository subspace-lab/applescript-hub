-- title: Get now playing info
-- verified: no
-- macos: 13+
-- app: Music
-- description: Get the name, artist, album, and duration of the currently playing track
-- source: app dictionary

tell application "Music"
	if player state is playing then
		set trackName to name of current track
		set trackArtist to artist of current track
		set trackAlbum to album of current track
		set trackDuration to duration of current track
		set trackPosition to player position
		return {trackName, trackArtist, trackAlbum, trackDuration, trackPosition}
	else
		return "Nothing is playing."
	end if
end tell
