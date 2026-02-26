-- title: Introspect UI hierarchy of an app
-- verified: no
-- macos: 13+
-- app: System Events
-- description: Discover UI elements (buttons, menus, text fields) in an app window for RPA automation
-- source: macosxautomation.com
-- notes: Requires assistive access. Essential for discovering element names before automating.

-- List all UI elements of the front window of an app
on get_ui_elements(app_name)
	tell application app_name to activate
	delay 0.5
	tell application "System Events"
		tell process app_name
			-- Get top-level UI elements of the front window
			set uiList to {}
			tell front window
				set allElements to entire contents
				repeat with elem in allElements
					set elemClass to class of elem
					set elemDesc to description of elem
					try
						set elemName to name of elem
					on error
						set elemName to "(no name)"
					end try
					set end of uiList to (elemClass as text) & ": " & elemName & " — " & elemDesc
				end repeat
			end tell
			return uiList
		end tell
	end tell
end get_ui_elements

-- Example: inspect TextEdit's UI
get_ui_elements("TextEdit")
