Attribute VB_Name = "RangeHelpers"
'@Folder "Helpers.Objects"
Option Explicit

Private Sub TestAppendRange()
    Dim runningRange As Range
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    Dim sel As Range
    
    Set sel = ws.Range("A2")
    Debug.Print sel.Value2
    AppendRange sel, runningRange

    Set sel = ws.Range("b6")
    Debug.Print sel.Value2
    AppendRange sel, runningRange

    Set sel = ThisWorkbook.Worksheets(2).Range("b7")
    Debug.Print sel.Value2
    AppendRange sel, runningRange
    
    Debug.Print runningRange.Address
End Sub

Public Sub AppendRange(ByVal rangeToAppend As Range, ByRef unionRange As Range)
    If rangeToAppend Is Nothing Then Exit Sub
    
    If unionRange Is Nothing Then
        Set unionRange = rangeToAppend
        Exit Sub
    End If
    
    If Not rangeToAppend.parent Is unionRange.parent Then Exit Sub
    
    Set unionRange = Application.Union(unionRange, rangeToAppend)
End Sub

