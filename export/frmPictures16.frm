VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPictures16 
   Caption         =   "frmPictures16"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmPictures16.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPictures16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@IgnoreModule HungarianNotation, MemberNotOnInterface
'@Folder "Version4.Resources"
Option Explicit

Public Function GetImageList() As ImageList
    Static Result As ImageList
    If Not Result Is Nothing Then
        Set GetImageList = Result
        Exit Function
    End If
    
    Set Result = New ImageList
    With Result
        .ImageWidth = 16
        .ImageHeight = 16
    End With
    
    Dim i As Long
    For i = 0 To Me.Controls.Count - 1
        If Me.Controls.Item(i).Tag = "IMG" Then
            Dim ThisLabel As MSForms.Label
            Set ThisLabel = Me.Controls.Item(i)
            Result.ListImages.Add Key:=ThisLabel.Name, Picture:=ThisLabel.Picture
        End If
    Next i
    
    Set GetImageList = Result
End Function

