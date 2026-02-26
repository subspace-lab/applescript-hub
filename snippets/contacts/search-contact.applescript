-- title: Find contact by name
-- verified: no
-- macos: 13+
-- app: Contacts
-- description: Search for a contact by name and return their info
-- source: app dictionary

set searchName to "John"

tell application "Contacts"
	set matchingPeople to every person whose name contains searchName
	set results to {}
	repeat with p in matchingPeople
		set pName to name of p
		set pOrg to organization of p
		set pEmails to value of every email of p
		set pPhones to value of every phone of p
		set end of results to {pName, pOrg, pEmails, pPhones}
	end repeat
	return results
end tell
