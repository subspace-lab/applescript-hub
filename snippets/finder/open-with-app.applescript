-- title: Open file with specific application
-- verified: no
-- macos: 13+
-- app: Finder
-- description: Open a file using a specified application instead of the default
-- source: app dictionary

set targetFile to (path to desktop folder as text) & "example.txt"
set targetApp to "TextEdit"

tell application "Finder"
	open file targetFile using application file id "com.apple.TextEdit"
end tell

-- Alternative: use the app name directly
-- tell application "Finder"
-- 	open file targetFile using application targetApp
-- end tell
