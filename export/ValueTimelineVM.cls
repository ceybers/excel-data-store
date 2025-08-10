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

Private Type TValueTimelineVM
    'KeyValue As String
    'KeyID As String
    'FieldValue As String
    'FieldID As String
    Values As RemoteValuesVM
End Type
Private This As TValueTimelineVM

Public Property Get Values() As RemoteValuesVM
    Set Values = This.Values
End Property

Public Sub Load(ByVal Values As Variant, ByVal Remote As Remote)
    Set This.Values = New RemoteValuesVM
    This.Values.Load Values, Remote
End Sub
