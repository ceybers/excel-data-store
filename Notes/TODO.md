# TODO
## Next
- [x] Pull action should check if one or more cells contain a Formula. Ask user before overwriting them.
- [x] BUG Tables with more than one local column mapped to the same Remote Field results in error 457 key already associated.
- [ ] BUG Case where Numbers formatted as text get stored in the Remote data store as number Variant types while the original cell is a String Variant type. Both Push and Pull actions will detect these values as different and continuously try to update.
- [ ] BUG Trying to manually enter a new Key Path changes the focus to Auto Map button on each key press.
- [ ] Ribbon > Remote > Save/Close to only proceed if Remote store is already open.

## Nice to Have
### Keys
- Handle cases where the same Key appears multiple times in a table.
  - Pull action should update all instances, not just the first.
  - Push action should ignore them with a warning.
    - Technically it could Push new Values if all the rows have the same new Value, but this seems like a bad pattern to accomodate.
- Compound Keys/Keys made out of multiple columns.
  - Could either use a delimiter behind the scenes, or prompt the user for one.
  - Would need to completely redo the UI for selecting Key column.
- Multiple Keys in a single table.
  - Fields would need to associated with a specific Key Column to know which to use when doing a lookup.
  - Could potentially use the Key Path prefix to match the Field Paths. (Wasn't this the entire point of them?)
  - Lots of edge cases to handle here. Would require major overhaul/rewriting from scratch. 
- Synonym resolving for Keys (sym-links).
- KeyGroups (aka Collections, Key Sets) (Tags?)
- Key lifecycle management (e.g., soft-deletion, deprecation)
- Pull Keys-only action
  - i.e., Pull all existing Keys from Remote and Append to bottom of table (even when no fields mapped)

### Fields
- UI for maintaining Fields (CRUD)
- Field lifecycle management (e.g., soft-deletion, deprecation)
- Data Types for Fields 
  - String, Integer, Float, Currency, Date, Boolean
- Default Number Formatting for Fields

## Values
- When updating Columns, shrink the Range being updated to the sub range of FirstCellAffected to LastCellAffected. This is to minimise the amount of cells written to with `.Value2`.
- UI for browsing Value history for selected Field×Key. (i.e., Timeline of changes.)
  - [x] Single Value read-only history
  - [x] Formating using the NumberFormat of selected cell. 
  - [x] Modeless UI that traps `Worskheet_SelectionChange`.
  - [ ] Roll back to previous value.
  - [ ] Multiple Key(s) and Field(s) modeless dialog.

## Push/Pull Actions
- [ ] Conditional Push/Pull
  - Permit/Deny changes for Push/Pull of:
  - `0→A` New Values
  - `A→B` Existing Value Changed
  - `A→0` Existing Value Deleted
  - e.g., only allow worksheet to Push to Remote if there are no existing Values in the Remote.
  - e.g., only allow worksheet to update empty cells from Remote (don't overwrite existing Values, even if they are out-of-date).
  - i.e., Only Push values from worksheet to Remote if there is no existing value in the Remote.
- [ ] Different color highlighting for `0→A`, `A→B`, and `A→0` changes.
- [ ] Pull All/Partial by a specific Commit.
- [ ] Bulk Ingest data
  - i.e., Local table has columns for Key, Field, Value _and_ Date. 
  - Optionally also a column for Commit Caption.
  - Backdating Commits (e.g., storing Historical SCD data)
  
## Complete
- [x] ~~BUG Issue with RemoteKeys if there is only 1 key for a Path.~~
- [x] ~~BUG Pull/Push not using the most recent Map for a table when it has more than one map.~~
- [x] ~~BUG Pulling or Partial Pulling with Filters applied results in unexpected behaviour.~~
- [x] ~~BUG FIX Partial Pull, Partial Push and Highlight Changes not working properly.~~
- [x] ~~Check if Working(LO) is protected (or opened in read only mode) to prevent errors.~~

# Notes
- Surrogate key in Values is SHA1[0,8] of Latest (`KeyID×FieldID`) and `CreationTime`.
- Mapped Fields in TableMap should remain in TableMap Range stored on Worksheet, even if most recent push or pull doesn't include them (e.g., Push on subset of ListObject)
- Unique ListObject ID looks to be BuiltinDocumentProperties 3 Author 11 Creation Date. 
  - Most of the other fields are Empty or errors (WinWord specific?)
- Bug with ListViews in a MultiPage control. 
  - If a ListView is on a Page that is not the default Page when the UserForm is opened, if the user changes to the page with the ListView control for the first time, the ListView will be at the incorrect position of 0,0. If the user changes to a different page and back, the ListView will then show in the correct position. 
  - A workaround is to set the Visible property to False then back to True on the MultiPage Change event.

# Rambling
- As per [Zawinski's *Law of Software Development*](https://www.laws-of-software.com/laws/zawinski/), by the time we're done, the GUI will look like Outlook 1997. The navigator tree will be Keys/Fields/Values/Commits. The "email's" `ListView` will be a list of Commits, with each Commit being an "email". The contents of each "email" will be a `<TABLE>` of the new values in Key×Field×Value format.

---
⏏️ [Back to README](../README.md)