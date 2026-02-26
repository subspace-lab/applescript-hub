-- title: Execute a shell command in Terminal
-- verified: no
-- macos: 13+
-- app: Terminal
-- description: Run a shell command in a new or existing Terminal tab
-- source: app dictionary

set shellCommand to "ls -la ~/Desktop"

tell application "Terminal"
	activate
	-- Run in a new window
	do script shellCommand

	-- To run in the frontmost tab instead, use:
	-- do script shellCommand in front window
end tell
