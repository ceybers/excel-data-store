VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SerializedTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder("Version4.Model")
Option Explicit

Public Enum MatchType
    NO_MATCH
    LISTOBJECT_NAME
    WORKSHEET_NAME
    WORKBOOK_NAME
End Enum

Private Type TSerializedTable
    ListObject As ListObject
    
    WorkbookFullName As String
    WorkbookPath As String
    WorkbookName As String
    WorksheetName As String
    ListObjectName As String
End Type
Private This As TSerializedTable

Public Property Get WorkbookFullName() As String
    WorkbookFullName = This.WorkbookFullName
End Property

Public Property Get WorkbookPath() As String
    WorkbookPath = This.WorkbookPath
End Property

Public Property Get WorkbookName() As String
    WorkbookName = This.WorkbookName
End Property

Public Property Get WorksheetName() As String
    WorksheetName = This.WorksheetName
End Property

Public Property Get ListObjectName() As String
    ListObjectName = This.ListObjectName
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Set ListObject(ByVal RHS As ListObject)
    Set This.ListObject = RHS
End Property

Public Function Serialize() As String
    Dim Tokens(0 To 4) As String
    Tokens(0) = This.ListObject.Parent.Parent.FullName
    Tokens(1) = This.ListObject.Parent.Parent.Path ' WB Path only
    Tokens(2) = This.ListObject.Parent.Parent.Name 'WB
    Tokens(3) = This.ListObject.Parent.Name ' WS
    Tokens(4) = This.ListObject.Name ' LO
    
    Serialize = Join$(Tokens, Chr$(ASCII_FS))
End Function

Public Sub Deserialize(ByVal SerializedString As String)
    Dim Tokens As Variant
    Tokens = Split(SerializedString, Chr$(ASCII_FS))
    Debug.Assert UBound(Tokens) = 4
    
    With This
        .WorkbookFullName = Tokens(0)
        .WorkbookPath = Tokens(1)
        .WorkbookName = Tokens(2)
        .WorksheetName = Tokens(3)
        .ListObjectName = Tokens(4)
    End With
End Sub

Public Function TryMatch() As MatchType
    If This.ListObject Is Nothing Then Exit Function
    
    If This.ListObject.Name = This.ListObjectName Then
        TryMatch = LISTOBJECT_NAME
        Exit Function
    ElseIf This.ListObject.Parent.Name = This.WorksheetName Then
        If This.ListObject.Parent.ListObjects.Count = 1 Then
            TryMatch = WORKSHEET_NAME
            Exit Function
        End If
    End If
End Function

Public Function TryResolve() As MatchType
    If TryResolveByListObjectName Then
        TryResolve = MatchType.LISTOBJECT_NAME
        Exit Function
    ElseIf TryResolveByWorksheetName Then
        TryResolve = MatchType.WORKSHEET_NAME
        Exit Function
    ElseIf TryResolveByWorkbookName Then
        TryResolve = MatchType.WORKBOOK_NAME
        Exit Function
    End If
    
    TryResolve = NO_MATCH
End Function

Private Function TryResolveByListObjectName() As Boolean
    Dim Workbook As Workbook
    If Not TryGetWorkbook(This.WorkbookName, Workbook) Then Exit Function
    
    Dim Worksheet As Worksheet
    If Not WorksheetHelpers.TryGetWorksheet(This.WorksheetName, Workbook, Worksheet) Then Exit Function
    
    Dim ListObject As ListObject
    If Not TryGetListObjectInWorksheet(This.ListObjectName, Worksheet, ListObject) Then Exit Function
    
    Set This.ListObject = ListObject
    TryResolveByListObjectName = True
End Function

Private Function TryResolveByWorksheetName() As Boolean
    Dim Workbook As Workbook
    If Not TryGetWorkbook(This.WorkbookName, Workbook) Then Exit Function
    
    Dim Worksheet As Worksheet
    If Not WorksheetHelpers.TryGetWorksheet(This.WorksheetName, Workbook, Worksheet) Then Exit Function
        
    If Worksheet.ListObjects.Count <> 1 Then Exit Function
    
    Set This.ListObject = Worksheet.ListObjects.Item(1)
    TryResolveByWorksheetName = True
End Function

Private Function TryResolveByWorkbookName() As Boolean
    Dim ListObjects As Collection
    Set ListObjects = New Collection
    
    Dim Workbook As Workbook
    If Not TryGetWorkbook(This.WorkbookName, Workbook) Then Exit Function
    
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        Dim ListObject As ListObject
        For Each ListObject In Worksheet.ListObjects
            ListObjects.Add ListObject
        Next ListObject
    Next Worksheet
    
    If ListObjects.Count = 1 Then
        Set This.ListObject = ListObjects.Item(1)
        TryResolveByWorkbookName = True
    End If
End Function
