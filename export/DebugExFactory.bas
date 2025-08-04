Attribute VB_Name = "DebugExFactory"
'@Folder("Version4.Factories")
Option Explicit

Public Function Log() As IDebugEx
    Static mLog As IDebugEx
    If mLog Is Nothing Then
        Set mLog = DebugEx.Create
        'mLog.AddProvider ImmediateLoggingProvider.Create
    End If
    Set Log = mLog
End Function

