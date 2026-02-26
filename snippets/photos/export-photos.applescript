-- title: Export selected photos
-- verified: no
-- macos: 13+
-- app: Photos
-- description: Export the currently selected photos to a folder on the Desktop
-- source: app dictionary
-- notes: Photos must be selected in the Photos app before running

set exportFolder to (POSIX path of (path to desktop folder)) & "Exported Photos/"

-- Create export folder if needed
do shell script "mkdir -p " & quoted form of exportFolder

tell application "Photos"
	set selectedItems to selection
	if (count of selectedItems) is 0 then
		return "No photos selected."
	end if
	export selectedItems to (POSIX file exportFolder as alias)
	return "Exported " & (count of selectedItems) & " photos."
end tell
