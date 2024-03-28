Attribute VB_Name = "TestColumnVisibleToArray"
'@IgnoreModule
'@Folder "Tests.Helpers"
Option Explicit
Option Private Module

Public Sub Test()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim rng As Range
    Dim v As Variant
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets.Item(1)
    Set lo = ws.ListObjects.Item(1)
    Set lc = lo.ListColumns.Item(1)
    Set rng = lc.DataBodyRange                   '.SpecialCells(xlCellTypeVisible)
    v = VisibleRangeToArray(rng)
    
    Debug.Print UBound(v, 1)
    Debug.Print UBound(v, 2)
    
    Debug.Assert False
End Sub

