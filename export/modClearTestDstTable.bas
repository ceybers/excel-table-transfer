Attribute VB_Name = "modClearTestDstTable"
'@Folder("TestTableProperties")
Option Explicit

'@EntryPoint "DoClearTestDstTable"
Public Sub DoClearTestDstTable()
    Dim Range As Range
    Set Range = Worksheets(1).ListObjects("TestDstTable").DataBodyRange
    Set Range = Range.Offset(0, 1)
    Set Range = Range.Resize(, Range.Columns.Count - 1)
    Range.Clear
End Sub
