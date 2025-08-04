VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Remote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RemoteDataStore.Models"
Option Explicit

Private Const DO_DEBUG As Boolean = True
Private Const WS_KEYS_NAME As String = "Keys"
Private Const WS_FIELDS_NAME As String = "Fields"
Private Const WS_VALUES_NAME  As String = "Values"
Private Const WS_COMMITS_NAME  As String = "Commits"
Private Const WS_MAPS_NAME  As String = "Maps"
Private Const WS_TABLES_NAME  As String = "Tables"

Private Type TState
    Workbook As Workbook
    RemoteKeys As RemoteKeys
    RemoteFields As RemoteFields
    RemoteValues As RemoteValues
    RemoteCommits As RemoteCommits
    RemoteMaps As RemoteMaps
    RemoteTables As RemoteTables
End Type
Private This As TState

Public Property Get Keys() As RemoteKeys
    If DO_DEBUG Then Debug.Print "Remote.Keys Property Get"
    Set Keys = This.RemoteKeys
End Property

Public Property Get Fields() As RemoteFields
    If DO_DEBUG Then Debug.Print "Remote.Fields Property Get"
    If This.RemoteFields Is Nothing Then
            Set This.RemoteFields = New RemoteFields
            This.RemoteFields.Load This.Workbook.Worksheets.Item(WS_FIELDS_NAME)
    End If
    Set Fields = This.RemoteFields
End Property

Public Property Get Values() As RemoteValues
    If DO_DEBUG Then Debug.Print "Remote.Values Property Get"
    If This.RemoteValues Is Nothing Then
        Set This.RemoteValues = New RemoteValues
        This.RemoteValues.Load This.Workbook.Worksheets.Item(WS_VALUES_NAME)
    End If
    Set Values = This.RemoteValues
End Property

Public Property Get Commits() As RemoteCommits
    If DO_DEBUG Then Debug.Print "Remote.Commits Property Get"
    Set Commits = This.RemoteCommits
End Property

Public Property Get Maps() As RemoteMaps
    If DO_DEBUG Then Debug.Print "Remote.Maps Property Get"
    Set Maps = This.RemoteMaps
End Property

Public Property Get Tables() As RemoteTables
    If DO_DEBUG Then Debug.Print "Remote.Tables Property Get"
    Set Tables = This.RemoteTables
End Property

Public Sub Load(ByVal Workbook As Workbook)
    If DO_DEBUG Then Debug.Print "Remote.Load()"
    Set This.Workbook = Workbook
    Reload
End Sub

Public Sub Reload()
    With This
        Set .RemoteKeys = New RemoteKeys
        .RemoteKeys.Load This.Workbook.Worksheets.Item(WS_KEYS_NAME)
        
        If Not .RemoteFields Is Nothing Then
            Set .RemoteFields = New RemoteFields
            .RemoteFields.Load This.Workbook.Worksheets.Item(WS_FIELDS_NAME)
        End If
        
        If Not .RemoteValues Is Nothing Then
            Set .RemoteValues = New RemoteValues
            .RemoteValues.Load This.Workbook.Worksheets.Item(WS_VALUES_NAME)
        End If
        
        Set .RemoteMaps = New RemoteMaps
        .RemoteMaps.Load This.Workbook.Worksheets.Item(WS_MAPS_NAME)
        
        Set .RemoteTables = New RemoteTables
        .RemoteTables.Load This.Workbook.Worksheets.Item(WS_TABLES_NAME)
        
        Set .RemoteCommits = New RemoteCommits
        .RemoteCommits.Load This.Workbook.Worksheets.Item(WS_COMMITS_NAME)
    End With
End Sub
