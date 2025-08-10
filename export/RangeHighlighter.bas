Attribute VB_Name = "RangeHighlighter"
'@Folder("Version4.Queries")
Option Explicit

Public Enum HighlightColor
    hcCompareOnly = 0
    hcBeforePull
    hcAfterPull
    hcBeforePush
    hcAfterPush
    hcHeaders
    hcKey
End Enum

Private Const COND_FMT_FORMULA As String = "=TRUE+N(""9dd8b78c-2b33-4313-8dcf-6c1870ee0ef4"")"

Public Sub HighlightRange(ByVal Range As Range, ByVal HighlightColor As HighlightColor)
    Dim FormatCondition  As FormatCondition
    Set FormatCondition = Range.FormatConditions.Add(Type:=xlExpression, Formula1:=COND_FMT_FORMULA)
    With FormatCondition
        .Interior.Color = GetColor(HighlightColor)
        If HighlightColor = hcHeaders Or HighlightColor = hcKey Then .Font.Color = RGB(0, 0, 0)
    End With
End Sub

Public Sub RemoveHighlights(ByVal ListObject As ListObject)
    Dim FormatCondition As FormatCondition
    Do While TryGetFormatConditionByFormula(COND_FMT_FORMULA, ListObject.Range, FormatCondition)
        FormatCondition.Delete
    Loop
End Sub

Private Function TryGetFormatConditionByFormula(ByVal Formula As String, ByVal Range As Range, ByRef OutFormatCondition As FormatCondition) As Boolean
    Dim ThisFormatCondition As Object
    For Each ThisFormatCondition In Range.FormatConditions
        If TypeOf ThisFormatCondition Is FormatCondition Then
            Dim ThisFormatConditionB As FormatCondition
            Set ThisFormatConditionB = ThisFormatCondition
            If ThisFormatConditionB.Formula1 = Formula Then
                Set OutFormatCondition = ThisFormatConditionB
                TryGetFormatConditionByFormula = True
                Exit Function
            End If
        End If
    Next ThisFormatCondition
End Function

Private Function GetColor(ByVal ColorEnum As HighlightColor) As Long
    Select Case ColorEnum
        Case hcCompareOnly
            GetColor = 6750105 ' Highlighter green
        Case hcBeforePull
            GetColor = 3243501 ' Orange
        Case hcAfterPull
            GetColor = 11389944 ' Light Orange
        Case hcBeforePush
            GetColor = 15773696 ' Blue
        Case hcAfterPush
            GetColor = 15652797 ' Light blue
        Case hcHeaders
            GetColor = RGB(255, 255, 0) ' Yellow
        Case hcKey
            GetColor = RGB(255, 0, 0) ' Red
    End Select
End Function
