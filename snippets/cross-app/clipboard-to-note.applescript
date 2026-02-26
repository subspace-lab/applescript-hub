-- title: Save clipboard contents to a new note
-- verified: no
-- macos: 13+
-- app: Notes
-- description: Get the current clipboard text and create a new note with it
-- source: original

set clipText to the clipboard as text

if clipText is "" then
	display dialog "Clipboard is empty." buttons {"OK"} default button "OK"
	return
end if

-- Format with timestamp
set noteTitle to "Clipboard — " & (current date) as text
set noteBody to "<h1>" & noteTitle & "</h1><p>" & clipText & "</p>"

tell application "Notes"
	make new note at folder "Notes" of default account with properties {¬
		name:noteTitle, ¬
		body:noteBody ¬
	}
end tell

return "Saved clipboard to Notes: " & noteTitle
