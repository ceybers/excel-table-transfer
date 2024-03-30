Attribute VB_Name = "TransferInstructionSerializer"
'@Folder("Model2.TransferInstruction2")
Option Explicit

Private Const ASC_UNIT_SEPARATOR As Long = 134 '31
Private Const ASC_RECORD_SEPARATOR As Long = 135 '30

Public Function Serialize(ByVal Transfer As TransferInstruction2) As String
    Dim Result(0 To 6) As Variant
    
    Result(0) = SerializeTable(Transfer.Source.Table)
    Result(1) = Transfer.Source.KeyColumnName
    Result(2) = SerializeTable(Transfer.Destination.Table)
    Result(3) = Transfer.Destination.KeyColumnName
    Result(4) = Transfer.Source.ValueColumnCount
    Result(5) = Join(Transfer.Source.ValueColumns, Chr$(ASC_UNIT_SEPARATOR)) ' Variant/Variant(0 to 1); Variant/Object/ListColumn
    Result(6) = Join(Transfer.Destination.ValueColumns, Chr$(ASC_UNIT_SEPARATOR))
    
    Serialize = Join(Result, Chr$(ASC_RECORD_SEPARATOR))
End Function

Public Function TryDeserialize(ByVal SerialString As String, ByRef OutTransfer As TransferInstruction2) As Boolean
    Set OutTransfer = Deserialize(SerialString)
    TryDeserialize = OutTransfer.Source.IsValid And OutTransfer.Destination.IsValid
End Function

Public Function Deserialize(ByVal SerialString As String) As TransferInstruction2
    Dim Transfer As TransferInstruction2
    Set Transfer = New TransferInstruction2
    
    'SerialString = "C:\Users\User\Repos\Public\excel-table-transfer\Development.xlsm†Development.xlsm†Sheet1†Table1‡KeyA‡C:\Users\User\Repos\Public\excel-table-transfer\Development.xlsm†Development.xlsm†Sheet1†Table2‡KeyB‡2‡data2†data3‡data2†data3"
    
    Dim SerialSplit As Variant ' Variant/String(0 to 6)
    SerialSplit = Split(SerialString, Chr$(ASC_RECORD_SEPARATOR))
    
    If UBound(SerialSplit) = 6 Then
        TryRehydrateTableTransfer SerialSplit(0), SerialSplit(1), SerialSplit(5), Transfer.Source
        TryRehydrateTableTransfer SerialSplit(2), SerialSplit(3), SerialSplit(6), Transfer.Destination
    End If
    
    Set Deserialize = Transfer
End Function

Public Sub TryRehydrateTableTransfer(ByVal SerialTable As Variant, ByVal KeyColumnName As Variant, _
ByVal SerialValues As Variant, ByVal TransferTable As TransferTable)
    Dim AllTables As Collection
    Set AllTables = ListObjectHelpers.GetAllTablesInApplication
    
    Dim SerialTableSplit As Variant
    SerialTableSplit = Split(SerialTable, Chr$(ASC_UNIT_SEPARATOR))
    
    Dim Workbook As Workbook
    Dim ListObject As ListObject
    If TryGetWorkbook(SerialTableSplit(1), Workbook) Then
        If TryGetListObjectFromWorkbook(Workbook, SerialTableSplit(3), ListObject) Then
            Set TransferTable.Table = ListObject
        End If
    End If
    
    If TransferTable.Table Is Nothing Then
        If TryGetListObjectFromCollection(AllTables, SerialTableSplit(3), ListObject) Then
            Set TransferTable.Table = ListObject
        End If
    End If
    
    If Not TransferTable.Table Is Nothing Then
        TransferTable.Load _
            Table:=TransferTable.Table, _
            mKeyColumnName:=KeyColumnName, _
            ValueColumnNames:=Split(SerialValues, Chr$(ASC_UNIT_SEPARATOR))
    End If
End Sub

Private Function SerializeTable(ByVal ListObject As ListObject) As String
    Dim Result(0 To 3) As Variant
    With ListObject
        Result(0) = .Parent.Parent.FullName ' Workbook
        Result(1) = .Parent.Parent.Name ' Workbook
        Result(2) = .Parent.Name ' Worksheet
        Result(3) = .Name ' ListObject
    End With
    
    SerializeTable = Join(Result, Chr$(ASC_UNIT_SEPARATOR))
End Function


