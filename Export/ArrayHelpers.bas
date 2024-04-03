Attribute VB_Name = "ArrayHelpers"
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
'@Folder "Helpers.Objects"
Option Explicit

Private Const ERR_MSG_NOT_SINGLE_COLUMN As String = "DisjointRangeToArray only works with range with a column count of 1"
Private Const ERR_MSG_NO_VISIBLE_CELLS As String = "No visible cells in Destination Range"

' Used when Transferring and `HasFlag(This.Flags, DestinationFilteredOnly)`
' SourceArray must be of shape (1 To n, 1 To 1).
' DestinationRange must be exactly 1 column wide and n columns tall.
' Values from SourceArray will only be placed into cells in DestinationRange that are visible.
' Hidden cells in the Range will not be affected.
Public Sub ArrayToFilteredRange(ByVal SourceArray As Variant, ByVal DestinationRange As Range)
    If DestinationRange.Columns.Count <> 1 Then
        Err.Raise vbObjectError + 2, , ERR_MSG_NOT_SINGLE_COLUMN
    End If
    
    Dim FilteredRange As Range
    On Error Resume Next
        Set FilteredRange = DestinationRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If FilteredRange Is Nothing Then
        ' Exit Sub
        Err.Raise vbObjectError + 2, , ERR_MSG_NO_VISIBLE_CELLS
    End If

    Dim FirstRow As Long
    FirstRow = DestinationRange.Rows.Item(1).Row
    
    Dim Area As Range
    For Each Area In FilteredRange.Areas
        Dim TopRow As Long
        TopRow = Area.Rows.Item(1).Row
        
        Dim BottomRow As Long
        BottomRow = Area.Rows.Item(Area.Rows.Count).Row
        
        Dim AreaHeight As Long
        AreaHeight = BottomRow - TopRow + 1
        
        Dim ValueVariant As Variant
        ValueVariant = Area.Value2
        
        If AreaHeight = 1 Then
            ValueVariant = SourceArray(1 + TopRow - FirstRow, 1)
        Else
            Dim i As Long
            For i = 1 To AreaHeight
                ValueVariant(i, 1) = SourceArray(i + TopRow - FirstRow, 1)
            Next
        End If
        
        Area.Value2 = ValueVariant
    Next Area
End Sub

