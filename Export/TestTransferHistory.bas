Attribute VB_Name = "TestTransferHistory"
'@Folder("TransferHistory")
Option Explicit

Public Sub ATest()
    Dim ti As TransferInstruction
    Set ti = TestTransferInstruction.GetTestTransferInstruction
    
    Debug.Print ti.IsValid
    
    Dim c As Long
    c = 12 + ti.ValuePairs.Count

    Dim arr() As String
    ReDim arr(1 To c, 1 To 4)
    
    arr(1, 1) = "TRANSFER"
    arr(2, 2) = "NAME"
    arr(3, 2) = "TIME"
    arr(4, 2) = "PATH"
    arr(5, 2) = "FN"
    arr(6, 2) = "SHEET"
    arr(7, 2) = "RNG"
    arr(8, 2) = "TBL"
    arr(9, 2) = "KEYS"
    arr(10, 2) = "FLAGS"
    arr(11, 2) = "PAIRS"
    arr(c, 1) = "END"
    
    arr(2, 3) = ti.Name
    arr(3, 3) = Now()
    'arr(3,3) = ti.Source.Parent.
    
    arr(4, 3) = ti.Source.parent.parent.path
    arr(5, 3) = ti.Source.parent.parent.Name
    arr(6, 3) = ti.Source.parent.Name
    arr(7, 3) = ti.Source.Range.Address
    arr(8, 3) = ti.Source.Name
    arr(9, 3) = ti.SourceKey.Name
    
    arr(4, 4) = ti.Destination.parent.parent.path
    arr(5, 4) = ti.Destination.parent.parent.Name
    arr(6, 4) = ti.Destination.parent.Name
    arr(7, 4) = ti.Destination.Range.Address
    arr(8, 4) = ti.Destination.Name
    arr(9, 4) = ti.DestinationKey.Name
    
    arr(10, 3) = ti.Flags
    arr(11, 3) = ti.ValuePairs.Count
    
    Dim n As Long
    n = 12
    Dim cp As ColumnPair
    For Each cp In ti.ValuePairs
        arr(n, 3) = cp.LHS.Name
        arr(n, 4) = cp.RHS.Name
        n = n + 1
    Next cp
    
    Dim rng As Range
    Set rng = ThisWorkbook.Worksheets("CAETransferTableHistory").Range("L1").Resize(c, 4)
    
    rng.Value2 = arr
    rng.Select
End Sub

Public Sub BTest()
    Dim rng As Range
    Set rng = ActiveWorkbook.Worksheets("CAETransferTableHistory").Range("L1")
    
    Debug.Print rng.offset(9, 2).Address
    
    Dim c As Long
    c = 12 + rng.offset(10, 2).Value2
    
    Dim arr As Variant
    arr = rng.Resize(c, 4).Value2
    
    Debug.Assert arr(1, 1) = "TRANSFER"
    Debug.Assert arr(c, 1) = "END"
    
    Dim ti As TransferInstructionUnref
    Set ti = New TransferInstructionUnref
    
    Debug.Print arr(2, 3)
    
    ti.SourceFilename = arr(5, 3)
    ti.DestinationFilename = arr(5, 4)
    ti.Source = arr(8, 3)
    ti.Destination = arr(8, 4)
    ti.SourceKey = arr(9, 3)
    ti.DestinationKey = arr(9, 4)
    ti.Flags = arr(10, 3)
    
    Dim vpArr As Variant
    vpArr = rng.offset(11, 2).Resize(c - 12, 2).Value2
    ti.ValuePairs = vpArr
    
    Dim refTI As TransferInstruction
    Set refTI = ti.AsReferenced
    
    Stop
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
