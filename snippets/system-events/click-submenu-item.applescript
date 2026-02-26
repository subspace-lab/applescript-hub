-- title: Click submenu item via UI scripting
-- verified: no
-- macos: 13+
-- app: System Events
-- description: Reusable handler to navigate and click a submenu item in any app
-- source: macosxautomation.com
-- notes: Requires assistive access enabled in System Settings > Privacy & Security

-- Click a submenu item in any application
-- Usage: do_submenu("Safari", "View", "Toolbars", "Show Tab Bar")
on do_submenu(app_name, menu_name, submenu_name, menu_item)
	tell application app_name to activate
	delay 0.5
	tell application "System Events"
		tell process app_name
			click menu item menu_item of menu submenu_name of menu item submenu_name of menu menu_name of menu bar 1
		end tell
	end tell
end do_submenu

-- Example: open a specific submenu item
do_submenu("Finder", "View", "Arrange By", "Name")
