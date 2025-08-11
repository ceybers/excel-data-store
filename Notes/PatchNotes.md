# Patch Notes
## v0.6.4-beta (2025/08/11)
- [x] Check if Worksheet is Protected and exit gracefully.
- [x] Check if Workbook is opened in Protected Viewing and exit gracefully.
- [x] Blank cells in Key column will be ignored instead of throwing errors.
- [x] Error cells (#N/A) in Key column will be ignored instead of throwing errors.
- [x] Error cells (#N/A) in Value cells will be treated as variant type vbEmpty.
- [x] Remote UI now has buttons to Save or Save & Close the Repo worksheet.  
 
## v0.6.3-beta (2025/08/10)
- [x] New View History UI for a single selected Value (Key x Field)

## v0.6.2-beta (2025/08/10)
- [x] New TableMapGUI:
  - TreeView for browsing Fields in their paths.
  - Search for Table Column Names and Field Names.
  - AutoMap all.
  - Reset all Column-Field mappings.

## v0.6.1-beta (2025/08/06)
- [x] Pull and Push (Full and Partial) now count changes and asks for confirmation before committing changes.
- [x] Highlight can now highlight Key and Mapped Fields 
- [x] Remove all Highlights EntryPoint on ribbon.
- [x] New icons for ribbon and UserForms.

## v0.6.0-beta (2025/08/06)
- [x] The best fitting Table Map is now picked when opening the TableMap UI or running Pull/Push.
  - Scored based on table name, worksheet name, workbook name, workbook path, whether the Key column is present in the table, and how many of the Fields in the Table Map are present in the selected table.
- [x] UI for browsing Table Maps in the Data Store, including results of how well they match the selected table.   

## v0.5.6-beta (2025/08/05)
- [x] Remote can now rebuild missing IDs on Keys and Fields.
- [x] Remote UI now loads from Remote on demand instead of upfront.
- [x] The most updated Table Map should be used after saving an existing Table Map.
- [x] BUG FIX Highlight Changes would crash if there was a Duplicate Values conditional format present.

## v0.5.5-beta (2025/08/04)
- [x] Implemented DebugEx library for logging.
- [x] Changed Remote* LoadFromArray to ByRef instead of copying the array for each call for each row. 
- [x] Fixed performance issue with matching Keys to KeyIDs to rows in worksheet.

## v0.5.4-beta (2025/08/04)
- [x] Hash Lengths increased from 4 characters to full 40-character length.
- [x] RemoteValues now only loads Latest KeyxField Value.
 
## v0.5.3-beta (2025/08/04)
- [x] Fixed bug with not being able to Add New Fields via Remote Manager UI.
- [x] Temporarily disabled Column Sorting in Remote Manager UI for Fields.
 
## v0.5.2-beta (2025/08/04)
- [x] Added icons to UI forms.
- [x] Moved MsgBox strings to Constants resource file.
 
## v0.5.1-beta (2025/08/03)
- [x] Fixed bug when deserialising a table map that no longer matches the table (e.g., columns renamed).
- [x] Added a README.md, some documentation and some screenshots.

## v0.5.0-beta (2025/08/03)
- [x] Pull Dry-Run (Highlight only, don't replace existing)
- [x] Fixed caching issue with RemoteValues Latest lookup
- [x] Fixed RemoteValue casting .Value variants to String. This broke Dry Run highlighting (Change Mask) for number and date values. 
- [x] Entrypoints for Excel Ribbon
- [x] Ribbon .exportedUI file (i.e., choice of icons)

## v0.4.3-alpha (2025/08/03)
- [x] Public git repo
- [x] Housekeeping and Rubberducking

## v0.4.2-alpha (2025/07/30)
- [x] RemoteCommits and RemoteCommit working OK

## v0.4.1-alpha (2025/07/27)
- [x] Pull/Push/Partial Pull/Partial Push 4 working OK
- [x] Abandoned ACE ADODB because it cannot DELETE rows.
- [x] Implemented TableMapMVVM
- [x] Removeed LocalRepoMVVM
- [x] Remote-side saved Serialized TableMap and Serialized ListObject references.
- [x] "Add New Field" in Remote MVVM
  - Equivalent for Keys would be to type in a new KeyPath in TableMapMVVM's KeyPath ComboBox. 
- ~~Refactor RepoManager to use CustomXML instead of VeryHidden Worksheets~~

## v0.4.0-alpha (2025/07/26)
- [x] Refactored version 4 with different model (TableMap -> MappedTable)

## v0.3.0-alpha (2025/07/24)
- [x] RepoManager should find existing Local Repos properly now.
- [x] LocalRepo2MVVM should handle draft Local Repos properly now.
- [x] LocalRepo2 can serialize and deserialize to a given Worksheet.
  - Still can't create new Worksheet if it needs to.
    
## v0.2.5-alpha (2025/07/22)
- [x] RemoteValue and RemoteValues (replacing `Values`)
- [x] Remote2 GetKeyIDs(from Key values)
- [x] Remote2 Pull, Compare and Push using KeyIDs and FieldIDs
  
## v0.2.4-alpha (2025/07/20)
- [x] Starting MVVM-down refactoring
- [x] Remote GUI shows (read-only) Keys: Path TreeView and Key ListView
- [x] Remote GUI shows (read-only) Fields (ListView). Sortable by column.
- [x] LocalRepo GUI from ListObject and Remote (no existing LocalRepo deserialized)
- [x] Changed LocalRepo serialization format/layout.  
  - Actual Serialize/Deserialize NYI (WIP Serialize working OK)
- [x] RepoManager can now open a LocalRepo2 GUI
- [x] Cleaned and separated concerns:
  - Repository Manager (view ListObjects)
  - LocalRepo GUI (map keys and fields)
  - Remote Data Store GUI (view Keys and Fields)
  - Renamed everything to make it easier to deprecate v1
  - Cleaned and standardised @EntryPoints 
  
## v0.2.3-alpha (2025/07/19)
- [x] Draft GUI/MVVM for `RepoManager`
- [x] Housekeeping, renaming, commenting, refactoring 

## v0.2.2-alpha (2025/07/13)
- [x] DataStore automatically loads a `datastore.xlsx` file if not already open.
- [x] Data is now correctly saved to a separate file and not the `.xlsm` file. 
- [x] Persistent User-level storage remembers where `datastore.xlsx` is saved.
- [x] GUI for Mapping Local Key Column to Remote KeyPath
- [x] GUI for Mapping Local ListColumns to Remote FieldPath & FieldNames
- [x] GUI loads and saves RemoteMap in Staging Worksheet  
- [x] Tested and works with two different Local files. 

## v0.2.1-alpha (2025/07/12)
- [x] Rewrote Query and LocalRepo.
- [x] Working (partial) -> Staging -> Remote Push
- [x] Working (partial) <- Staging <- Remote Pull
- [x] Pull no longer fails if there are not Latest Values for any KeyxField pairs in a query.
- [x] Keys and Fields should no longer write Hashes to their table as numbers when they can be parsed as scientific/numeric.
- [x] Last Push/Last Pull timestamp updates. 

## v0.2.0-alpha (2025/07/10)
- [x] Rewrote TableMap (Working <- -> Staging) to replace Query class.
  - Shape of data structure changed. Fields are no longer duplicated in header and content.

## v0.1.0-alpha (2025/07/07)
- [x] Initial commit 🎂
- [x] Staging -> Remote Push works OK
- [x] Staging <- Remote Pull works OK
- [x] Keys are added if not present. IDs are SHA1 hash.
- [x] Fields are added if not present. IDs are SHA1 hash.
- [x] Values insert first new row if none-present
- [x] Values does not if existing Latest matches Pushed value
- [x] Values inserts new row and updates existing Latest to blank if Pushed value differs (SCD).
- [x] Working -> Staging Push
- [x] Working <- Staging Pull
- [x] Working -> Staging Init
  - i.e., Create Unmapped Query
  - With a given LO, single LC for Key, and collection/array of LC for fields.
  - Should populate everything _except_ the green cells (Remote Paths & Names)

---
⏏️ [Back to README](../README.md)