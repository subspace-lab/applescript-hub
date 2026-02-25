# App Dictionaries

Pre-extracted and human-readable documentation of scriptable app dictionaries.

## Generating a Dictionary

```bash
# Raw SDEF (XML)
sdef /Applications/AppName.app > app-dictionaries/appname.sdef

# For browsing in Script Editor
# File → Open Dictionary → select .sdef file
```

## Extraction Method

SDEF files are copied directly from app bundles (no Xcode required):

```bash
cp /System/Library/CoreServices/Finder.app/Contents/Resources/Finder.sdef app-dictionaries/finder.sdef
cp /System/Applications/Mail.app/Contents/Resources/Mail.sdef app-dictionaries/mail.sdef
cp /System/Applications/Calendar.app/Contents/Resources/iCal.sdef app-dictionaries/calendar.sdef
cp "/System/Volumes/Preboot/Cryptexes/App/System/Applications/Safari.app/Contents/Resources/Safari.sdef" app-dictionaries/safari.sdef
cp /System/Applications/Notes.app/Contents/Resources/Notes.sdef app-dictionaries/notes.sdef
```

Find the SDEF for any app:
```bash
find /Applications/AppName.app -name "*.sdef"
```

## Coverage

| App | File | Lines | macOS version |
|-----|------|-------|---------------|
| Finder | `finder.sdef` | 545 | 15.x |
| Mail | `mail.sdef` | 961 | 15.x |
| Calendar | `calendar.sdef` | 288 | 15.x |
| Safari | `safari.sdef` | 164 | 15.x |
| Notes | `notes.sdef` | 201 | 15.x |

Contributions welcome — open a PR with SDEFs from other apps or macOS versions.
