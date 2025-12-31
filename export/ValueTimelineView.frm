VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ValueTimelineView 
   Caption         =   "ValueTimelineView"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8175
   OleObjectBlob   =   "ValueTimelineView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ValueTimelineView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ImplicitDefaultMemberAccess, ArgumentWithIncompatibleObjectType, HungarianNotation
'@Folder "Version4.Views"
Option Explicit

Implements IView

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mViewModel As ValueTimelineVM
Attribute mViewModel.VB_VarHelpID = -1

Private Type TState
    'ViewModel As ValueTimelineVM
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdCancel_Click()
    OnCancel
End Sub

Private Sub cmdClose_Click()
    OnCancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set mViewModel = ViewModel
    This.IsCancelled = False
    
    InitializeControls
    UpdateControls
    
    Me.Show vbModeless
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub InitializeControls()
    InitValuesListView Me.lvValues
    ' This UserForm only gets a reference to the VM and can only catch the
    ' SelectionChanged event _after_ it has already fired for the first time.
    LoadValuesToListView mViewModel.Values, Me.lvValues
    UpdateControls
End Sub

Private Sub InitValuesListView(ByVal ListView As ListView)
    With ListView
        .BorderStyle = ccNone
        .LabelEdit = tvwManual
        .FullRowSelect = True
        .HideSelection = False
        .View = lvwReport
    End With
    
    With ListView.ColumnHeaders
        .Add Text:="Value", Width:=(ListView.Width - 92 - 160) - 16
        .Add Text:="Timestamp", Width:=92 '(ListView.Width / 3) - 16
        .Add Text:="Commit", Width:=160 '(ListView.Width / 3) - 16
    End With
End Sub

Private Sub LoadValuesToListView(ByVal RemoteValues As RemoteValuesVM, ByVal ListView As ListView)
    ListView.ListItems.Clear
    
    Dim i As Long
    For i = 1 To RemoteValues.Count
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Text:=Format$(RemoteValues.Item(i).Value, mViewModel.NumberFormat))
        ListItem.ListSubItems.Add Text:=RemoteValues.Item(i).Timestamp
        ListItem.ListSubItems.Add Text:=RemoteValues.Item(i).Commit
    Next i
End Sub

Private Sub UpdateControls()
    Exit Sub
End Sub

Private Sub mViewModel_SelectionChanged()
    LoadValuesToListView mViewModel.Values, Me.lvValues
    UpdateControls
End Sub
