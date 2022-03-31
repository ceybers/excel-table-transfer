Attribute VB_Name = "TestTransferHistory"
'@Folder("TransferHistory")
Option Explicit

Public Sub ATest()
    Dim tiUr As TransferInstructionUnref
    TransferHistorySerializer.TryLoad tiUr
    
    Dim ti As TransferInstruction
    Set ti = New TransferInstruction
    
    'Set ti.Source = ThisWorkbook.Worksheets(1).ListObjects(1)
    'Set ti.Destination = ThisWorkbook.Worksheets(1).ListObjects(2)
    Set ti.UnRef = tiUr
    
    ti.LoadFlags
    Debug.Print "ti.TryLoadTables "; ti.TryLoadTables
    Debug.Print "ti.TryLoadKeyColumns "; ti.TryLoadKeyColumns
    Debug.Print "ti.TryLoadValuePairs "; ti.TryLoadValuePairs
    
    Debug.Print "ti.Transfer "; ti.Transfer
    
    'Set rng = ActiveWorkbook.Worksheets("CAETransferTableHistory").Range("L20")
    'ti.SaveToRange rng
    
    ' TransferHistorySerializer.TrySave ti
    'Stop
End Sub

Public Sub Test()
    Dim vm As TransferHistoryViewModel
    Set vm = New TransferHistoryViewModel
    'Set vm.ActiveTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim view As IView
    Set view = New TransferHistoryView
    If view.ShowDialog(vm) Then
        Debug.Assert Not vm.SelectedInstruction Is Nothing
        'Debug.Print "Selected Instruction: " & vm.SelectedInstruction.Name
        'modMain.DoTransferTable vm.SelectedInstruction
    Else
        'Debug.Print "FAIL"
    End If
End Sub
