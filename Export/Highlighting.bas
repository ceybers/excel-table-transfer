Attribute VB_Name = "Highlighting"
'@Folder "Model.Formatting"
Option Explicit

Private Const MAGIC_FORMULA As String = "=""HighlightMapped""=""HighlightMapped"""
Private Const MAPPED_COLOR As Long = 10092492 '#CCFF99

Public Sub ApplyHighlighting(ByVal Range As Range, ByVal DoHighlight As Boolean)
    RemoveExistingHighlighting Range.parent
    
    If DoHighlight = False Then
        Exit Sub
    End If
    
    Dim FormatConditions As FormatConditions
    Set FormatConditions = Range.FormatConditions
    
    FormatConditions.Add Type:=xlExpression, Formula1:=MAGIC_FORMULA
    
    With FormatConditions.Item(FormatConditions.Count)
        .Interior.Color = MAPPED_COLOR
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
