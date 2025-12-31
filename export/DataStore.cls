VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "RemoteDataStore.Models"
Option Explicit

Private Const DEFAULT_REPO_PATH As String = "C:\Users\User\Repos\Private\excel-dev\excel-but-git\"
Private Const DEFAULT_REPO_NAME As String = "repo-datastore.xlsx"
Private Const SETTINGS_FILENAME As String = "excel-but-git.ini"
Private Const SETTINGS_UUID As String = "{ff144951-2f4a-4072-8016-6ffb7a338ae1}"
Private Const SETTING_REPOPATH As String = "RepoPath"
Private Const SETTING_REPONAME As String = "RepoName"
Private Const MAGIC_PHRASE_A1 As String = "DATASTORE"

Private Type TState
    RepoPath As String
    RepoName As String
    Workbook As Workbook
    Remote As Remote
End Type
Private This As TState

Public Property Get Remote() As Remote
    Set Remote = This.Remote
End Property

Public Sub Load()
    GetDataStoreFullPath
    GetOrOpenDataStore
    CreateRemote
End Sub

Public Sub Unload()
    If This.Workbook Is Nothing Then Exit Sub
    This.Workbook.Close SaveChanges:=True
    Set This.Workbook = Nothing
End Sub

Private Sub GetDataStoreFullPath()
    Static UserSettings As ISettings
    If Not UserSettings Is Nothing Then Exit Sub
    
    Set UserSettings = MyDocSettings.Create(SETTINGS_UUID, SETTINGS_FILENAME)
    
    With UserSettings
        This.RepoPath = .GetSetting(SETTING_REPOPATH)
        This.RepoName = .GetSetting(SETTING_REPONAME)
    
        If This.RepoPath = vbNullString Then
            This.RepoPath = DEFAULT_REPO_PATH
            .SetSetting SETTING_REPOPATH, DEFAULT_REPO_PATH
        End If
        
        If This.RepoName = vbNullString Then
            This.RepoName = DEFAULT_REPO_NAME
            .SetSetting SETTING_REPONAME, DEFAULT_REPO_NAME
        End If
    End With
End Sub

Private Sub GetOrOpenDataStore()
    If Not TryGetWorkbook(This.RepoName, This.Workbook) Then
        TryOpenWorkbook
    End If
End Sub

Private Sub TryOpenWorkbook()
    Dim WorkbookFilename As String
    WorkbookFilename = This.RepoPath & "\" & This.RepoName
    
    On Error GoTo ErrorCouldNotOpenWorkbook
    Application.Workbooks.Open Filename:=WorkbookFilename, Notify:=False
    On Error GoTo 0
    
    Set This.Workbook = Application.Workbooks.Item(This.RepoName)
    
    This.Workbook.Windows.Item(1).Visible = False
    Exit Sub
    
ErrorCouldNotOpenWorkbook:
    MsgBox MSG_CANT_OPEN_DATASTORE_WB, vbExclamation + vbOKOnly, APP_TITLE
End Sub

Private Sub CreateRemote()
    If Not IsDatastoreWorkbookValid(This.Workbook) Then
        MsgBox MSG_DATASTORE_WB_INVALID, vbExclamation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Set This.Remote = New Remote
    This.Remote.Load This.Workbook
End Sub

Private Function IsDatastoreWorkbookValid(ByVal Workbook As Workbook) As Boolean
    If Workbook Is Nothing Then Exit Function
    IsDatastoreWorkbookValid = This.Workbook.Worksheets.Item(1).Range("A1").Value2 = MAGIC_PHRASE_A1
End Function
