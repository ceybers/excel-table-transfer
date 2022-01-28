Attribute VB_Name = "TestTransferHistory"
'@Folder("TransferHistory")
Option Explicit

Public Sub Test()
    Dim vm As TransferHistoryViewModel
    Set vm = New TransferHistoryViewModel
    'Set vm.ActiveTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim view As IView
    Set view = New TransferHistoryView
    If view.ShowDialog(vm) Then
        Debug.Assert Not vm.SelectedInstruction Is Nothing
        'Debug.Print "Selected Instruction: " & vm.SelectedInstruction.Name
        modMain.DoTransferTable vm.SelectedInstruction
    Else
        'Debug.Print "FAIL"
    End If
End Sub
