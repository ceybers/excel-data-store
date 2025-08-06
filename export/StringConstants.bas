Attribute VB_Name = "StringConstants"
'@Folder "Version4.Resources.Constants"
Option Explicit

Public Const APP_TITLE As String = "Excel Data Store Tool"
Public Const APP_VERSION As String = "Version 0.6.1-beta"
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

Public Const MSG_REMOTE_REBUILD_KEYS As String = "Remote Keys Table rebuild OK!"
Public Const MSG_REMOTE_REBUILD_FIELDS As String = "Remote Fields Table rebuild OK!"

Public Const MSG_PULL_CONFIRM As String = "Update table with {0} changes from Data Store?"
Public Const MSG_PUSH_CONFIRM As String = "Update Data Store with {0} changes from table?"
Public Const MSG_PULL_NOCHANGES As String = "No new values found in Data Store to update this table with."
Public Const MSG_PUSH_NOCHANGES As String = "No new values found in this table to update the Data Store with."
