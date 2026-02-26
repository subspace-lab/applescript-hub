-- title: Click a button in a window by name
-- verified: no
-- macos: 13+
-- app: System Events
-- description: Click a named button in the frontmost window of an application
-- source: macosxautomation.com
-- notes: Requires assistive access enabled in System Settings > Privacy & Security

-- Click a button by name in an app's front window
on click_button(app_name, button_name)
	tell application app_name to activate
	delay 0.5
	tell application "System Events"
		tell process app_name
			click button button_name of front window
		end tell
	end tell
end click_button

-- Example: click OK button in a dialog
click_button("System Settings", "OK")
