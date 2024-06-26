VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.ValueColumns"
Option Explicit

Public Name As String
Public IsSelectable As Boolean

Public IsKeyColumn As Boolean
Public IsMapped As Boolean
Public IsFormula As Boolean
Public IsEmpty As Boolean

Public DataType As Long

Public HasErrors As Boolean
Public HasBoolean As Boolean
Public HasNumbers As Boolean
Public HasText As Boolean

Public Sub Load(ByVal ListColumn As ListColumn)
    Name = ListColumn.Name
    
    If IsNull(ListColumn.DataBodyRange.HasFormula) Then
        IsFormula = True ' At least one cell has a formula
    Else
        IsFormula = ListColumn.DataBodyRange.HasFormula ' True is all has formula, False if none have formula
    End If
    
    IsEmpty = IsListColumnEmpty(ListColumn)
    
    HasErrors = HasSpecialCells(ListColumn.DataBodyRange, xlCellTypeConstants, xlErrors)
    HasBoolean = HasSpecialCells(ListColumn.DataBodyRange, xlCellTypeConstants, xlLogical)
    HasNumbers = HasSpecialCells(ListColumn.DataBodyRange, xlCellTypeConstants, xlNumbers)
    HasText = HasSpecialCells(ListColumn.DataBodyRange, xlCellTypeConstants, xlTextValues)
End Sub

Private Function IsListColumnEmpty(ByVal ListColumn As ListColumn) As Boolean
    Dim EmptyCellRange As Range
    
    On Error Resume Next
    Set EmptyCellRange = ListColumn.DataBodyRange.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    
    If Not EmptyCellRange Is Nothing Then
        IsListColumnEmpty = (EmptyCellRange.Count = ListColumn.DataBodyRange.Count)
    End If
End Function

