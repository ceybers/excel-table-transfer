Attribute VB_Name = "ZZ_TransferHistorySerializer"
'@Folder("TransferHistory")
Option Explicit

Private Const TRANSFER_SERIALIZED_OBJECT_ROW_COUNT As Integer = 8

Public Sub Test()
    Dim result As Variant
    Set result = LoadTransferInstructionsFromWorksheet(ThisWorkbook.Worksheets("CAETransferTableHistory"))
    
    Dim ti As TransferInstructionUnref
    Set ti = result(1)
    
    Debug.Print ti.Flags
End Sub

Public Function LoadTransferInstructionsFromWorksheet(ByVal ws As Worksheet) As Variant
    Dim TransferInstructions
    Set TransferInstructions = New Collection
    
    Dim rng As Range
    Set rng = ws.UsedRange
    
    Dim arr As Variant
    arr = rng.Value2
    
    If IsEmpty(arr) Then
        Set LoadTransferInstructionsFromWorksheet = Nothing
        Exit Function
    End If
    
    Dim c As Long
    c = UBound(arr, 1)
    
   
    Dim att As String
    Dim val As String
    
    Dim lhs As String
    Dim rhs As String
    
    Dim curArr As Variant
    Dim ti As TransferInstructionUnref
    Dim i As Long
    For i = 1 To c
        curArr = Split(arr(i, 1), " ")
        
        Debug.Print UBound(curArr, 1)
        Select Case UBound(curArr, 1)
            Case 0
                If curArr(0) = "TRANSFER" Then
                    Set ti = New TransferInstructionUnref
                ElseIf curArr(0) = "END" Then
                    TransferInstructions.Add ti
                    Set ti = Nothing
                End If
                
            Case 1
                att = Split(curArr(1), ",")(0)
                val = Split(curArr(1), ",")(1)
                
                Select Case att
                    Case "SRC"
                        ti.Source = val
                    Case "DST"
                        ti.Destination = val
                    Case "SRCKEY"
                        ti.SourceKey = val
                    Case "DSTKEY"
                        ti.DestinationKey = val
                    Case "FLAGS"
                        ti.Flags = val
                End Select
                
            Case 2
                lhs = Split(curArr(2), ",")(0)
                rhs = Split(curArr(2), ",")(1)
                ' This will fail if columns have commas in them
        End Select
    Next i
    
    Set LoadTransferInstructionsFromWorksheet = TransferInstructions
End Function

Public Function ZZ_LoadTransferInstructionsFromWorksheet(ByVal ws As Worksheet) As Variant
    Dim rng As Range
    Set rng = ws.UsedRange
    
    Dim arr As Variant
    arr = ws.UsedRange.Value2
    
    If IsEmpty(arr) Then ' TODO check if vbarray+vbvariant and not single cell
        Set ZZ_LoadTransferInstructionsFromWorksheet = Nothing
        Exit Function
    End If
    
    Dim i As Long
    Dim curStr As String
    Dim curArr As Variant
    
    Dim att As String
    Dim val As String
    
    Dim lhs As String
    Dim rhs As String
    
    Dim ti As TransferInstruction
    Dim tis As Collection
    
    Set tis = New Collection
    
    For i = 1 To UBound(arr, 1)
        curStr = arr(i, 1)
        curArr = Split(curStr, " ")
        'Debug.Print UBound(curArr, 1)
        Select Case UBound(curArr, 1)
            Case 0
                If curArr(0) = "TRANSFER" Then
                    Set ti = New TransferInstruction
                ElseIf curArr(0) = "END" Then
                    tis.Add ti
                    Set ti = Nothing
                End If
            Case 1
                att = Split(curArr(1), ",")(0)
                val = Split(curArr(1), ",")(1)
                Select Case att
                    Case "SRC"
                        TryGetTableFromText val, ti.Source
                    Case "DST"
                        TryGetTableFromText val, ti.Destination
                    Case "SRCKEY"
                        TryGetListColumnFromText val, ti.Source, ti.SourceKey
                    Case "DSTKEY"
                        TryGetListColumnFromText val, ti.Destination, ti.DestinationKey
                    Case "FLAGS"
                        ti.Flags = val
                End Select
            Case 2
                lhs = Split(curArr(2), ",")(0)
                rhs = Split(curArr(2), ",")(1)
                
                If Not ti.Source Is Nothing And Not ti.Destination Is Nothing Then
                    ti.ValuePairs.Add ColumnPair.Create(ti.Source.ListColumns(lhs), ti.Destination.ListColumns(rhs))
                End If
        End Select
    Next i
    
    Set ZZ_LoadTransferInstructionsFromWorksheet = tis
End Function

'@Description "Save a Collection of TransferInstruction to a worksheet"
Public Sub SaveTransferInstructionsFromWorksheet(ByVal coll As Collection, ByVal ws As Worksheet)
Attribute SaveTransferInstructionsFromWorksheet.VB_Description = "Save a Collection of TransferInstruction to a worksheet"
    Dim rows As Long
    Dim rng As Range
    Dim tgtArr As Variant
    Dim i As Long
    Dim offset As Long
    
    Debug.Assert Not coll Is Nothing
    Debug.Assert coll.Count > 0
    
    rows = coll.Count * TRANSFER_SERIALIZED_OBJECT_ROW_COUNT
    
    ws.UsedRange.Clear
    
    PasteArrayIntoWorksheet SerializeTransferInstruction(coll(1)), ws
    offset = 1
    offset = offset + UBound(SerializeTransferInstruction(coll(1)), 1)
    
    For i = 2 To coll.Count
        PasteArrayIntoWorksheet SerializeTransferInstruction(coll(i)), ws, offset, 1
        offset = offset + UBound(SerializeTransferInstruction(coll(i)), 1)
    Next i
End Sub

Private Function SerializeTransferInstruction(ByVal Transfer As TransferInstruction) As Variant
    Dim c As Long
    Dim i As Long
    Dim result As Variant
    
    c = 8 + Transfer.ValuePairs.Count
    
    ReDim result(1 To c, 1 To 1)
    
    result(1, 1) = "TRANSFER"
    result(2, 1) = " SRC," & Transfer.Source.Range.Address(external:=True)
    result(3, 1) = " SRCKEY," & Transfer.SourceKey.Name
    result(4, 1) = " DST," & Transfer.Destination.Range.Address(external:=True)
    result(5, 1) = " DSTKEY," & Transfer.DestinationKey.Name
    result(6, 1) = " FLAGS," & Transfer.Flags
    result(7, 1) = " VALUES," & Transfer.ValuePairs.Count
    
    For i = 1 To Transfer.ValuePairs.Count
        result(7 + i, 1) = "  " & Transfer.ValuePairs(i).ToString
    Next i
    
    result(c, 1) = "END"
    
    SerializeTransferInstruction = result
End Function
