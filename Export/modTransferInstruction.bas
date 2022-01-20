Attribute VB_Name = "modTransferInstruction"
'@Folder "TransferInstructions"
Option Explicit

Public Type TransferInstruction
    Source As ListObject
    Destination As ListObject
    SourceKey As ListColumn
    DestinationKey As ListColumn
    Columns As Variant
    ClearDestinationColumns As Boolean
    PasteIntoBlankCellsOnly As Boolean
    CopyNonBlankCellsOnly As Boolean
End Type

Public Sub RunTransferInstruction(ti As TransferInstruction)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim srcKeys As Variant
    Dim dstKeys As Variant
    Dim keys As Variant
    Dim mapping As Variant
    
    'srcKeys = ti.SourceKey.DataBodyRange.Value
    'dstKeys = ti.DestinationKey.DataBodyRange.Value

    
    ' Try to get filtered keys only
    srcKeys = VisibleRangeToArray(ti.SourceKey.DataBodyRange)
    dstKeys = VisibleRangeToArray(ti.DestinationKey.DataBodyRange)
    
    keys = ArrayIntersect(srcKeys, dstKeys)
    
    ReDim mapping(1 To ArrayLength(keys), 1 To 2)
    
    Dim i As Integer
    For i = 1 To ArrayLength(keys)
        mapping(i, 1) = ArrayFind(keys(i, 1), srcKeys)
        mapping(i, 2) = ArrayFind(keys(i, 1), dstKeys)
    Next i
    
    Dim srcValues As Variant
    Dim dstValues As Variant
    
    Dim j As Integer, k As Integer
    
    If ti.ClearDestinationColumns Then
        For j = 1 To UBound(ti.Columns, 1)
            ti.Destination.ListColumns(ti.Columns(j, 2)).DataBodyRange.Value = vbNullString
        Next j
    End If
    
    Dim thisVarType As VbVarType
    For j = 1 To UBound(ti.Columns, 1)
        srcValues = ti.Source.ListColumns(ti.Columns(j, 1)).DataBodyRange.Value
        dstValues = ti.Destination.ListColumns(ti.Columns(j, 2)).DataBodyRange.Value
        For k = 1 To UBound(mapping, 1)
            thisVarType = VarType(dstValues(mapping(k, 2), 1))
            If (ti.PasteIntoBlankCellsOnly = False) Or IsEmpty(dstValues(mapping(k, 2), 1)) Then
                If (ti.CopyNonBlankCellsOnly = False) Or Not (IsEmpty(srcValues(mapping(k, 1), 1))) Then
                    dstValues(mapping(k, 2), 1) = srcValues(mapping(k, 1), 1)
                End If
            End If
        Next k
        ti.Destination.ListColumns(ti.Columns(j, 2)).DataBodyRange.Value = dstValues
        DoEvents
    Next j
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ti.Destination.parent.Activate
End Sub

Private Function IsEmpty(v As Variant) As Boolean
    If VarType(v) = vbEmpty Then IsEmpty = True
    If VarType(v) = vbString And Len(v) = 0 Then IsEmpty = True
End Function
