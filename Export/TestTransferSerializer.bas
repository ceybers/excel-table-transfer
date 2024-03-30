Attribute VB_Name = "TestTransferSerializer"
'@Folder("MVVM2.AppContext")
Option Explicit

Private Transfer As TransferInstruction2

Private Sub TestTransferSerializer()
    TransferInstructionSerializer.Serialize Transfer
    
    Dim SerialString As String
    
    SerialString = "C:\Users\User\Repos\Public\excel-table-transfer\Development.xlsm†Development.xlsm†Sheet1†Table1‡KeyA‡C:\Users\User\Repos\Public\excel-table-transfer\Development.xlsm†Development.xlsm†Sheet1†Table2‡KeyB‡2‡data2†data3‡data2†data3"
    If TransferInstructionSerializer.TryDeserialize(SerialString, Transfer) Then
        Debug.Print "A"
    Else
        Debug.Print "B"
    End If
    
    SerialString = "C:\Users\User\Repos\Public\excel-table-transfer\Development.xlsm†Development.xlsm†Sheet1†Table1‡KeyAz‡C:\Users\User\Repos\Public\excel-table-transfer\Development.xlsm†Development.xlsm†Sheet1†Table2‡KeyB‡2‡data2†data3‡data2†data3"
    If TransferInstructionSerializer.TryDeserialize(SerialString, Transfer) Then
        Debug.Print "C"
    Else
        Debug.Print "D"
    End If
    
    SerialString = "C:\Users\User\Repos\Public\excel-table-transfer\Development.xlsm†Development.xlsm†Sheet1†Table1‡KeyA‡C:\Users\User\Repos\Public\excel-table-transfer\Development.xlsm†Development.xlsm†Sheet1†Table2‡KeyB‡2‡data2z†data3‡data2†data3"
    If TransferInstructionSerializer.TryDeserialize(SerialString, Transfer) Then
        Debug.Print "E"
    Else
        Debug.Print "F"
    End If
End Sub
