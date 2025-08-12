Attribute VB_Name = "StringConstants"
'@Folder "Version4.Resources.Constants"
Option Explicit

Public Const APP_TITLE As String = "Excel Data Store Tool"
Public Const APP_VERSION As String = "Version 0.6.5-beta"
Public Const APP_COPYRIGHT As String = "2025 Craig Eybers" & vbCrLf & "All rights reserved."

Public Const ASCII_US As Long = 164 '31
Public Const ASCII_RS As Long = 165 '30
Public Const ASCII_GS As Long = 166 '29
Public Const ASCII_FS As Long = 167 '28

Public Const MSG_CANT_OPEN_DATASTORE_WB As String = "Could not open Data Store Repository workbook."
Public Const MSG_DATASTORE_WB_INVALID As String = "Data Store workbook is not a valid repository."
Public Const MSG_MAP_NO_TABLE As String = "Please select a table first before running the Map command."
Public Const MSG_PULL_NO_TABLE As String = "Please select a table first before running the Pull command."
Public Const MSG_PUSH_NO_TABLE As String = "Please select a table first before running the Push command."
Public Const MSG_IS_PROTECTED As String = "Worksheet is Protected. Unprotect it and try again."

Public Const MSG_REMOTE_REBUILD_KEYS As String = "Remote Keys Table rebuild OK!"
Public Const MSG_REMOTE_REBUILD_FIELDS As String = "Remote Fields Table rebuild OK!"

Public Const MSG_PULL_CONFIRM As String = "Update table with {0} changes from Data Store?"
Public Const MSG_PUSH_CONFIRM As String = "Update Data Store with {0} changes from table?"
Public Const MSG_PULL_NOCHANGES As String = "No new values found in Data Store to update this table with."
Public Const MSG_PUSH_NOCHANGES As String = "No new values found in this table to update the Data Store with."
Public Const MSG_AUTOMAP As String = "Are you sure you want to Auto Map all fields?"
Public Const MSG_RESETALL As String = "Are you sure you want to reset mapping?"

Public Const NODE_PREFIX_PATH As String = "Path:"
Public Const NODE_KEY_UNMAPPED As String = "#UNMAPPED"
Public Const NODE_KEY_ROOT As String = "#ROOT"

Public Const TV_CAP_UNMAPPED As String = "(Unmapped)"
Public Const TV_CAP_ROOT As String = "Fields"

Public Const LV_COL_LCN As String = "Column Name"
Public Const LV_COL_MAP As String = "Mapped To"
Public Const LV_CAP_UNMAPPED As String = "(Key Column)"

Public Const IMG_BLANK As String = "imgBlank"
Public Const IMG_MAPPED As String = "imgMapped"
Public Const IMG_KEY As String = "imgKey"
Public Const IMG_FIELDS As String = "imgFields"
Public Const IMG_FIELD As String = "imgField"
Public Const IMG_FOLDEROPEN As String = "imgFolderOpen"
Public Const IMG_FOLDERCLSD As String = "imgFolderClosed"
