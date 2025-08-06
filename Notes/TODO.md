# TODO
## General
- [x] ~~BUG FIX Partial Pull, Partial Push and Highlight Changes not working properly.~~

## Keys
- Synonym resolving for Keys (sym-links).
- Pull Keys-only action
  - i.e., Pull all existing Keys from Remote and Append to bottom of table (even when no fields mapped)
- KeyGroups (aka Collections, Key Sets)
- Key management, including soft-deletion/deprecation.

## Fields
- Remote Field Manager UI (read/write)
- Rename existing Fields. 
- Delete existing Fields 
 - (Probably created in error)
 - Handle cases where code expects Remote Field to still exist

## Values
- GUI to browse Cell (KeyxField) history. i.e., Timeline of changes.

## Commits
- Pull by specific Commit
- Rollback Commits.
- Backdating Commits (i.e., storing Historical SCD data)

## Maps
- [x] ~~BUG Pull/Push not using the most recent Map for a table when it has more than one map.~~

## GUI
- Remote Manager UI to select TreeView/ListView items by Key, not Index.
- Filter remote fields to only fields that already have values for this KeyPath
- Handle cases where mapping/fields are empty (uninit)
- Handle cases where mapping is present but no fields are mapped yet (only table & key)
- Toggle to filter Remote Fields to only show fields that already have Values for this keys in the selected KeyPath.
- Captions for fields for eventual GUI.

## Unsorted
- Remote could keep a watermark to cache naive Pulls.
- Check if Working(LO) is protected (or opened in read only mode) to prevent errors.

# Notes
- Surrogate key in Values is SHA1[0,8] of Latest (KeyIDxFieldID) and CreationTime.
- Mapped Fields in TableMap should remain in TableMap Range stored on Worksheet, even if most recent push or pull doesn't include them (e.g., Push on subset of ListObject)
- Unique ListObject ID looks to be BuiltinDocumentProperties 3 Author 11 Creation Date. Most of the other fields are Empty or errors (WinWord specific?)
- Bug with ListViews in a MultiPage control. If a ListView is on a Page that is not the default Page when the UserForm is opened, if the user changes to the page with the ListView control for the first time, the ListView will be at the incorrect position of 0,0. If the user changes to a different page and back, the ListView will then show in the correct position. A workaround is to set the Visible property to False then back to True on the MultiPage Change event.

# Rambling
- As per Zawinski's Law of Software Development, by the time we're done, the GUI will look like Outlook 1997. The navigator tree will be Keys/Fields/Values/Commits. The "emails" ListView will be a list of Commits, with each Commit being an "email". The contents of each "email" will be a `<TABLE>` of the new values in KeyxFieldxValue format

---
⏏️ [Back to README](../README.md)