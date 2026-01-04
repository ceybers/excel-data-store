VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueTimelineVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Version4.ViewModels")
Option Explicit

Public Event SelectionChanged()

'@Ignore HungarianNotation
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mWorksheet As Worksheet
Attribute mWorksheet.VB_VarHelpID = -1

Private Type TValueTimelineVM
    Query As ValueTimelineQuery
    Values As RemoteValuesVM
    NumberFormat As String
End Type
Private This As TValueTimelineVM

Public Property Get Values() As RemoteValuesVM
    Set Values = This.Values
End Property

Public Property Get NumberFormat() As String
    NumberFormat = This.NumberFormat
End Property

Public Property Let NumberFormat(ByVal RHS As String)
    If RHS = "General" Then
        This.NumberFormat = "@"
    Else
        This.NumberFormat = RHS
    End If
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = mWorksheet
End Property

Public Property Set Worksheet(ByVal RHS As Worksheet)
    Set mWorksheet = RHS
End Property

Private Sub Class_Initialize()
    This.NumberFormat = "@"
End Sub

Public Sub Load(ByVal Query As ValueTimelineQuery)
    Set This.Query = Query
    
    UpdateSelectedCell
End Sub

'@Ignore HungarianNotation
Private Sub mWorksheet_SelectionChange(ByVal Target As Range)
    If GuardSelectionSingleCell Then Exit Sub
    UpdateSelectedCell
End Sub

Private Sub UpdateSelectedCell()
    Set This.Values = New RemoteValuesVM
    This.Query.MappedTable.SelectKeysAndFields PartialSelection:=True
    This.Query.Run
    This.Values.Load This.Query.Results, This.Query.Remote
    RaiseEvent SelectionChanged
End Sub
