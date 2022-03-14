Attribute VB_Name = "TestArrayToFilteredRange"
'@Folder("TableTransferTool")
Option Explicit

Private Sub Test()
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim rng As Range
    Dim fltRng As Range
    
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(2)
    Set lc = lo.ListColumns(4)
    Set rng = lc.DataBodyRange
    Set fltRng = rng.SpecialCells(xlCellTypeVisible)
    
    'Debug.Print "rng:"
    'PrintRange rng
    
    'Debug.Print "fltRng:"
    'PrintRange fltRng
    
    Dim mask As Variant
    mask = rng.Value2
    SetArrayElementsToEmpty mask
    mask(4, 1) = "dd"
    mask(6, 1) = "ff"
    mask(7, 1) = "gg"
    
    'Debug.Print "mask:"
    'PrintArray mask
    
    'DoWork rng, mask
End Sub


