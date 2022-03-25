Attribute VB_Name = "ZZZ_modMain"
'@IgnoreModule EmptyIfBlock
'@Folder "TableTransfer"
Option Explicit
Option Private Module

Public Sub ZZZ_TransferTable()
    Dim transfer As TransferInstruction
    Set transfer = New TransferInstruction
    DoTransferTable transfer
End Sub

Public Sub ZZZ_TransferTableFromHistory()
    Dim vm As TransferHistoryViewModel
    Set vm = New TransferHistoryViewModel
    
    Dim view As IView
    Set view = New TransferHistoryView
    If view.ShowDialog(vm) Then
        DoTransferTable vm.SelectedInstruction
    End If
End Sub

Private Sub DoTransferTable(transfer As TransferInstruction)
    If transfer.IsValid = True Then GoTo AlreadyValid
    If Selection.ListObject Is Nothing Then GoTo NoTableSelected
    
    Dim firstTable As ListObject
    Dim secondTable As ListObject
    
    Set firstTable = Selection.ListObject
    
    Dim IsSource As Boolean
    Dim IsDestination As Boolean
    If TryGetSourceOrDestination(IsSource, IsDestination) Then
        If IsSource Then
            Set transfer.Source = firstTable
        Else
            Set transfer.Destination = firstTable
        End If
    Else
        Exit Sub
    End If
        
    If TryGetSecondTable(firstTable, secondTable) Then
        If IsSource Then
            Set transfer.Destination = secondTable
        Else
            Set transfer.Source = secondTable
        End If
    Else
        Exit Sub
    End If

    'Set Transfer.SourceKey = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    'Set Transfer.DestinationKey = ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(1)
NoTableSelected:
    If GetKeyColumns(transfer) Then
        ' continue
    Else
        Exit Sub
    End If
    
    ' TODO Move default Flags somewhere better
    ' Transfer.Flags = AddFlag(Transfer.Flags, ClearDestinationFirst)
    'Transfer.Flags = AddFlag(Transfer.Flags, ReplaceEmptyOnly)
    'Transfer.Flags = AddFlag(Transfer.Flags, TransferBlanks)
    'Transfer.Flags = AddFlag(Transfer.Flags, SourceFilteredOnly)
    ' Transfer.Flags = AddFlag(Transfer.Flags, DestinationFilteredOnly)
AlreadyValid:
    If SetValueMapping(transfer) Then
        ' continue
    Else
        Exit Sub
    End If
    
    'Dim lhs As ListObject
    'Dim rhs As ListObject
    'Set lhs = ThisWorkbook.Worksheets(1).ListObjects(1)
    'Set rhs = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    'Transfer.ValuePairs.Add ColumnPair.Create(lhs.ListColumns(2), rhs.ListColumns(2))
    'Transfer.ValuePairs.Add ColumnPair.Create(lhs.ListColumns(3), rhs.ListColumns(3))
    'Transfer.ValuePairs.Add ColumnPair.Create(lhs.ListColumns(4), rhs.ListColumns(4))
     
    transfer.transfer
    
    If HasFlag(transfer.Flags, SaveToHistory) Then
        Dim history As TransferHistoryViewModel
        Set history = New TransferHistoryViewModel
        If history.HasHistory = False Then
            history.Create
        End If
        history.Refresh
        history.Add transfer
        history.Save
    End If
End Sub

Private Function TryGetSecondTable(ByVal SelectedTable As ListObject, ByRef OutTable As ListObject) As Boolean
    Dim vm As SelectTableViewModel
    Set vm = New SelectTableViewModel
    Set vm.ActiveTable = SelectedTable
    
    Dim frm As IView
    Set frm = New SelectTableView
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set OutTable = vm.SelectedTable
            TryGetSecondTable = True
        End If
    End If
End Function

Private Function TryGetSourceOrDestination(ByRef IsSource As Boolean, ByRef IsDestination As Boolean) As Boolean
    Dim vm As SourceOrDestinationViewModel
    Set vm = New SourceOrDestinationViewModel
    Set vm.ListObject = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim view As IView
    Set view = New SourceOrDestinationView
    If view.ShowDialog(vm) Then
        IsSource = vm.IsSource
        IsDestination = vm.IsDestination
        TryGetSourceOrDestination = True
    Else
        TryGetSourceOrDestination = False
    End If
End Function

Private Function GetKeyColumns(ByVal transfer As TransferInstruction) As Boolean
    Dim vm As KeyMapperViewModel
    Set vm = New KeyMapperViewModel
    Set vm.LHSTable = transfer.Source
    Set vm.RHSTable = transfer.Destination
    
    Dim frm As IView
    Set frm = New KeyMapperView
    
    If frm.ShowDialog(vm) Then
        If vm.IsValid Then
            Set transfer.Source = vm.LHSTable
            Set transfer.Destination = vm.RHSTable
            Set transfer.SourceKey = vm.LHSKeyColumn
            Set transfer.DestinationKey = vm.RHSKeyColumn
            GetKeyColumns = True
        Else
            MsgBox "Invalid VM"
        End If
    End If
End Function

Private Function SetValueMapping(ByVal transfer As TransferInstruction) As Boolean
    Dim vm As ValueMapperViewModel
    Set vm = New ValueMapperViewModel
    Set vm.LHS = transfer.Source
    Set vm.rhs = transfer.Destination
    Set vm.KeyColumnLHS = transfer.SourceKey
    Set vm.KeyColumnRHS = transfer.DestinationKey
    vm.Flags = transfer.Flags
    
    vm.LoadFromTransferInstruction transfer
    
    Dim frm As IView
    Set frm = New ValueMapperView
    
    If frm.ShowDialog(vm) Then
        Set transfer.ValuePairs = vm.checked
        transfer.Flags = vm.Flags
        SetValueMapping = True
    End If
End Function
