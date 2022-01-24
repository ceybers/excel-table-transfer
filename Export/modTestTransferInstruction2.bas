Attribute VB_Name = "modTestTransferInstruction2"
'@Folder "TransferInstructions2"
Option Explicit

Public Sub TestTransferInstruction2()
    
End Sub

Public Sub TestCompareKeyColumns()
    Dim compare As clsCompareKeyColumns
    Dim MapResult As Variant
    
    Set compare = clsCompareKeyColumns.Create(GetLHS, GetRHS)
    MapResult = compare.Map
    
    DoTransfer MapResult, GetSrc, GetDst
    Debug.Print "OK"
End Sub

Public Sub DoTransfer(ByRef Map As Variant, ByRef Source As ListColumn, ByRef Destination As ListColumn)
    Dim i As Integer
    Dim arrLHS As Variant
    Dim arrRHS As Variant
    Dim arrLHSOffset As Long
    Dim arrRHSOffset As Long
    Dim newValue As Variant
    
    arrLHS = Source.DataBodyRange.Value2
    arrRHS = Destination.DataBodyRange.Value2
    arrLHSOffset = Source.DataBodyRange.Row
    arrRHSOffset = Destination.DataBodyRange.Row
    
    For i = LBound(Map) To UBound(Map)
        If Map(i) > -1 Then
            newValue = arrLHS(Map(i) - arrLHSOffset + 1, 1)
            Select Case VarType(newValue)
                Case vbString
                    If newValue <> vbNullString Then
                        arrRHS(i - arrRHSOffset + 1, 1) = newValue
                    End If
                Case Else
                    arrRHS(i - arrRHSOffset + 1, 1) = newValue
            End Select
        Else
            arrRHS(i - arrRHSOffset + 1, 1) = "Unmapped"
        End If
    Next i
    
    Destination.DataBodyRange.Value2 = arrRHS
End Sub

Private Function GetLHS() As clsKeyColumn
    Set GetLHS = clsKeyColumn.FromRange(ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1).DataBodyRange)
End Function

Private Function GetRHS() As clsKeyColumn
    Set GetRHS = clsKeyColumn.FromRange(ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(1).DataBodyRange)
End Function

Private Function GetSrc() As ListColumn
    Set GetSrc = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(2)
End Function

Private Function GetDst() As ListColumn
    Set GetDst = ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(2)
End Function
