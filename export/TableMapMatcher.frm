VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TableMapMatcher 
   Caption         =   "Data Store Table Map Matcher"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12555
   OleObjectBlob   =   "TableMapMatcher.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TableMapMatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ArgumentWithIncompatibleObjectType, HungarianNotation
'@Folder("TableMapMatcher.Views")
Option Explicit

Implements IView

Private Type TState
    ViewModel As TableMapMatcherVM
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdCancel_Click()
    OnCancel
End Sub

Private Sub txtTableImage_Click()
    frmAbout.Show
End Sub

Private Sub cmdNew_Click()
    This.ViewModel.CreateNew = True
    Me.Hide
End Sub

Private Sub cmdOpen_Click()
    Me.Hide
End Sub

Private Sub lblTableImage_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmAbout.Show
End Sub

Private Sub lvTableMaps_DblClick()
    If Not This.ViewModel.GetSelectedMatch Is Nothing Then
        Me.Hide
    End If
End Sub

Private Sub lvTableMaps_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.SelectTableMapByIndex Item.Index
    UpdateControls
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set This.ViewModel = ViewModel
    This.IsCancelled = False
    
    InitializeControls
    UpdateControls
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeControls()
    Me.txtTableName.Value = This.ViewModel.ListObjectName
    LoadListViewFromMatches Me.lvTableMaps, This.ViewModel.Matches
    UpdateControls
End Sub

Private Sub UpdateControls()
    LoadListViewFromMatch Me.lvMatchResults, This.ViewModel.GetSelectedMatch
End Sub

Private Sub LoadListViewFromMatches(ByVal ListView As ListView, ByVal Matches As TableMapMatches)
    With ListView
        .ListItems.Clear
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .View = lvwReport
    End With
    With ListView.ColumnHeaders
        .Clear
        .Add Text:="Caption"
        .Add Text:="Score"
        .Add Text:="Timestamp"
        .Item(1).Width = 142
        .Item(2).Width = 32
        .Item(3).Width = 82
    End With
    
    Dim i As Long
    For i = 1 To Matches.Count
        Dim TableMapMatch As TableMapMatch
        Set TableMapMatch = Matches.Item(i)
        
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Text:=TableMapMatch.Caption)
        With ListItem
            .ListSubItems.Add Text:=TableMapMatch.Score
            .ListSubItems.Add Text:=Format$(TableMapMatch.Timestamp, "yyyy/mm/dd hh:MM")
            If TableMapMatch.Score = 0 Then
                .ForeColor = RGB(128, 128, 128)
                .ListSubItems.Item(1).ForeColor = RGB(128, 128, 128)
                .ListSubItems.Item(2).ForeColor = RGB(128, 128, 128)
            End If
        End With
    Next i
End Sub

Private Sub LoadListViewFromMatch(ByVal ListView As ListView, ByVal Match As TableMapMatch)
    With ListView
        .ListItems.Clear
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .View = lvwReport
    End With
    With ListView.ColumnHeaders
        .Clear
        .Add Text:="Item"
        .Add Text:="Value"
        .Add Text:="Match?"
        .Item(1).Width = 64
        .Item(2).Width = 140
        .Item(3).Width = 64
    End With
    
    Dim i As Long
    For i = 1 To UBound(Match.Results)
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Text:=Match.Results(i, 1))
        With ListItem
            .ListSubItems.Add Text:=Match.Results(i, 2)
            If VarType(Match.Results(i, 3)) = vbBoolean Then
                .ListSubItems.Add Text:=IIf(Match.Results(i, 3), "OK", "No Match")
            Else
                .ListSubItems.Add Text:=Match.Results(i, 3)
            End If
        End With
    Next i
End Sub
