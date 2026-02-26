-- title: Get output from current Terminal session
-- verified: no
-- macos: 13+
-- app: Terminal
-- description: Get the text contents (output) of the frontmost Terminal tab
-- source: app dictionary
-- notes: Returns all visible text in the terminal buffer

tell application "Terminal"
	-- Get contents of the selected tab of the front window
	set termOutput to contents of selected tab of front window
	return termOutput
end tell
