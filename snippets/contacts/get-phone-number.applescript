-- title: Get phone number for a contact
-- verified: no
-- macos: 13+
-- app: Contacts
-- description: Look up a contact by name and return their phone numbers with labels
-- source: app dictionary

set contactName to "Jane Doe"

tell application "Contacts"
	set matchingPeople to every person whose name is contactName
	if (count of matchingPeople) is 0 then
		return "No contact found with name: " & contactName
	end if

	set p to item 1 of matchingPeople
	set phoneList to {}
	repeat with ph in phones of p
		set end of phoneList to {label of ph, value of ph}
	end repeat
	return phoneList
end tell
