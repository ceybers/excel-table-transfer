Attribute VB_Name = "Highlighting"
'@Folder "MVVM.Model.Formatting"
Option Explicit

Private Const MAGIC_FORMULA As String = "=OR(TRUE,""HighlightMapped;b92d7b59-e7ec-4db0-a7c6-5a6ad86ceac2"")"
Private Const MAPPED_COLOR As Long = 10092492 '#CCFF99

Public Sub ApplyHighlighting(ByVal Range As Range, Optional ByVal Color As Long = MAPPED_COLOR)
    Dim FormatConditions As FormatConditions
    Set FormatConditions = Range.FormatConditions
    
    FormatConditions.Add Type:=xlExpression, Formula1:=MAGIC_FORMULA
    
    With FormatConditions.Item(FormatConditions.Count)
        .Interior.Color = Color
    End With
End Sub

Public Sub RemoveExistingHighlighting(ByVal Worksheet As Worksheet)
    Dim FormatConditions As FormatConditions
    Set FormatConditions = Worksheet.Cells.FormatConditions
    
    Dim FormatConditionToDelete As FormatCondition
    Do While TryFindFormatCondition(FormatConditions, FormatConditionToDelete)
        FormatConditionToDelete.Delete
    Loop
End Sub

Private Function TryFindFormatCondition(ByVal FormatConditions As FormatConditions, ByRef OutFormatCondition As FormatCondition) As Boolean
    Dim Item As Object
    For Each Item In FormatConditions
        If TypeOf Item Is FormatCondition Then
            Dim FormatCondition As FormatCondition
            Set FormatCondition = Item
            If FormatCondition.Formula1 = MAGIC_FORMULA Then
                Set OutFormatCondition = FormatCondition
                TryFindFormatCondition = True
                Exit Function
            End If
        End If
    Next Item
End Function
