Attribute VB_Name = "Highlighting"
'@Folder "MVVM.Model.Formatting"
Option Explicit

Public Sub ApplyHighlighting(ByVal Range As Range, Optional ByVal Color As Long = COLOR_DEFAULT_HIGHLIGHT)
    Dim FormatConditions As FormatConditions
    Set FormatConditions = Range.FormatConditions
    
    FormatConditions.Add Type:=xlExpression, Formula1:=MAGIC_FORMULA_HIGHLIGHTING
    
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
            If FormatCondition.Formula1 = MAGIC_FORMULA_HIGHLIGHTING Then
                Set OutFormatCondition = FormatCondition
                TryFindFormatCondition = True
                Exit Function
            End If
        End If
    Next Item
End Function
