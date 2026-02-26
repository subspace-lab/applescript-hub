-- title: Click menu item via UI scripting
-- verified: no
-- macos: 13+
-- app: System Events
-- description: Reusable handler to click a menu item in any app via accessibility UI
-- source: macosxautomation.com
-- notes: Requires assistive access enabled in System Settings > Privacy & Security

-- Click a menu item in any application
-- Usage: do_menu("TextEdit", "Format", "Make Plain Text")
on do_menu(app_name, menu_name, menu_item)
	tell application app_name to activate
	delay 0.5
	tell application "System Events"
		tell process app_name
			click menu item menu_item of menu menu_name of menu bar 1
		end tell
	end tell
end do_menu

-- Example: toggle word wrap in TextEdit
do_menu("TextEdit", "Format", "Wrap to Page")
