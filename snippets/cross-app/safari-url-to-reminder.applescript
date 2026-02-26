-- title: Save Safari URL as a reminder
-- verified: no
-- macos: 13+
-- app: Safari, Reminders
-- description: Grab the current Safari tab URL and title, create a reminder with it
-- source: original

-- Get the URL and title from Safari
tell application "Safari"
	set pageURL to URL of current tab of front window
	set pageTitle to name of current tab of front window
end tell

-- Create a reminder with the link
tell application "Reminders"
	set targetList to default list
	make new reminder in targetList with properties {¬
		name:"Read: " & pageTitle, ¬
		body:pageURL ¬
	}
end tell

return "Saved reminder: " & pageTitle
