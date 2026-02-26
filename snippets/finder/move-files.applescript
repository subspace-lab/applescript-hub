-- title: Move files to a folder
-- verified: no
-- macos: 13+
-- app: Finder
-- description: Move selected Finder items to a specified destination folder
-- source: app dictionary

set destPath to (path to desktop folder as text) & "Archive:"

tell application "Finder"
	-- Ensure destination exists
	if not (exists folder destPath) then
		make new folder at desktop with properties {name:"Archive"}
	end if

	set selectedItems to selection
	if (count of selectedItems) is 0 then
		return "No files selected."
	end if

	move selectedItems to folder destPath
	return "Moved " & (count of selectedItems) & " items."
end tell
