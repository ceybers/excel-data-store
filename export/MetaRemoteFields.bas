Attribute VB_Name = "MetaRemoteFields"
'@Folder("RemoteDataStore.Models.Plural")
Option Explicit

Public Const META_FIELD_COUNT As Long = 2

Public Function Unmapped() As RemoteField
    Static Result As RemoteField
    
    If Result Is Nothing Then
        Set Result = New RemoteField
        With Result
            .ID = FIELD_ID_UNMAPPED
            .Caption = "(Not mapped)"
        End With
    End If
    
    Set Unmapped = Result
End Function

Public Function AddNew() As RemoteField
    Static Result As RemoteField
    
    If Result Is Nothing Then
        Set Result = New RemoteField
        With Result
            .ID = FIELD_ID_ADDNEW
            .Caption = "(Add new field)"
        End With
    End If
    
    Set AddNew = Result
End Function
