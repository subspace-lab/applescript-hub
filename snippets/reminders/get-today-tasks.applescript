-- title: List today's reminders
-- verified: no
-- macos: 13+
-- app: Reminders
-- description: Get all incomplete reminders due today or overdue
-- source: app dictionary

tell application "Reminders"
	set todayStart to current date
	set time of todayStart to 0
	set todayEnd to todayStart + (1 * days)

	set todayTasks to {}
	repeat with r in (every reminder whose completed is false)
		try
			set dueD to due date of r
			if dueD ≤ todayEnd then
				set end of todayTasks to {name of r, due date of r}
			end if
		end try
	end repeat
	return todayTasks
end tell
