Attribute VB_Name = "modTest"
'@Folder("TableTransfer")
Option Explicit
Option Private Module

Public Sub TestListColumnFormula()
    Dim lc As ListColumn
    Set lc = ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(4)
    Debug.Print lc.DataBodyRange.FormulaArray
    Stop
End Sub
