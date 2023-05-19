Attribute VB_Name = "Highlighting"
'@Folder "MVVM.Common"
Option Explicit

Private Const MAGIC_FORMULA As String = "=""HighlightMapped""=""HighlightMapped"""
Private Const MAPPED_COLOR As Long = 10092492

Public Sub ApplyHighlighting(ByVal rng As Range, ByVal DoHighlight As Boolean)
    Dim fcs As FormatConditions
    Set fcs = Selection.parent.Cells.FormatConditions
    
    Dim ob As Object
    Dim fc As FormatCondition
    Dim fcToDelete As FormatCondition
    For Each ob In fcs
        If TypeOf ob Is FormatCondition Then
            Set fc = ob
            If fc.Formula1 = MAGIC_FORMULA Then
                Set fcToDelete = fc
            End If
        End If
    Next
    
    If Not fcToDelete Is Nothing Then
        fcToDelete.Delete
    End If
    
    If DoHighlight = False Then
        Exit Sub
    End If
    
    Set fcs = rng.FormatConditions
    fcs.Add Type:=xlExpression, Formula1:=MAGIC_FORMULA
    Set fc = fcs(fcs.Count)
    With fc
        .Interior.Color = MAPPED_COLOR
    End With
End Sub
