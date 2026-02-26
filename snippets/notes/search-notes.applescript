-- title: Search notes by text
-- verified: no
-- macos: 13+
-- app: Notes
-- description: Search for notes containing specific text and return their names
-- source: app dictionary

set searchTerm to "meeting"

tell application "Notes"
	set matchingNotes to every note whose plaintext contains searchTerm
	set results to {}
	repeat with n in matchingNotes
		set end of results to {name of n, creation date of n, name of container of n}
	end repeat
	return results
end tell
