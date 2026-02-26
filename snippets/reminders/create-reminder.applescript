-- title: Create a reminder with due date
-- verified: no
-- macos: 13+
-- app: Reminders
-- description: Create a new reminder in the default list with a due date and optional notes
-- source: app dictionary

tell application "Reminders"
	set targetList to default list
	set dueDate to (current date) + (1 * days) -- tomorrow

	make new reminder in targetList with properties {¬
		name:"Buy groceries", ¬
		body:"Milk, eggs, bread", ¬
		due date:dueDate, ¬
		priority:1 ¬
	}
end tell
