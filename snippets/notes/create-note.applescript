-- title: Create a new note
-- verified: no
-- macos: 13+
-- app: Notes
-- description: Create a new note with a title and body in the default account
-- source: app dictionary

tell application "Notes"
	set defaultAcct to first account whose name is (name of default account)
	set targetFolder to folder "Notes" of defaultAcct

	make new note at targetFolder with properties {¬
		name:"My New Note", ¬
		body:"<h1>My New Note</h1><p>This is the note body.</p>" ¬
	}
end tell
