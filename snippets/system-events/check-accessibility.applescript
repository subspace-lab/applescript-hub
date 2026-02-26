-- title: Check if assistive access is enabled
-- verified: no
-- macos: 13+
-- app: System Events
-- description: Check whether UI scripting (assistive access) is enabled, prompt user if not
-- source: original

tell application "System Events"
	if not UI elements enabled then
		display dialog "Assistive access is not enabled." & return & return & ¬
			"Go to System Settings > Privacy & Security > Accessibility" & return & ¬
			"and add this app to enable UI scripting." buttons {"Open Settings", "OK"} default button "OK"
		if button returned of result is "Open Settings" then
			tell application "System Settings"
				activate
				-- Open Privacy & Security > Accessibility
				open location "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility"
			end tell
		end if
		return false
	else
		return true
	end if
end tell
