Attribute VB_Name = "modMain"
'@Folder "TableTransfer"
Option Explicit

Public Sub TransferTable()
    Dim Transfer As TransferInstruction
    Set Transfer = New TransferInstruction
    DoTransferTable Transfer
End Sub

Public Sub TransferTableFromHistory()
    Dim vm As TransferHistoryViewModel
    Set vm = New TransferHistoryViewModel
    
    Dim view As IView
    Set view = New TransferHistoryView
    If view.ShowDialog(vm) Then
        DoTransferTable vm.SelectedInstruction
    End If
End Sub

Public Sub DoTransferTable(Transfer As TransferInstruction)
    If Transfer.IsValid = True Then GoTo AlreadyValid
    If Selection.ListObject Is Nothing Then GoTo NoTableSelected
    
    Dim firstTable As ListObject
    Dim secondTable As ListObject
    
    Set firstTable = Selection.ListObject
    
    Dim IsSource As Boolean
    Dim IsDestination As Boolean
    If TryGetSourceOrDestination(IsSource, IsDestination) Then
        If IsSource Then
            Set Transfer.Source = firstTable
        Else
            Set Transfer.Destination = firstTable
        End If
    Else
        Exit Sub
    End If
        
    If TryGetSecondTable(firstTable, secondTable) Then
        If IsSource Then
            Set Transfer.Destination = secondTable
        Else
            Set Transfer.Source = secondTable
        End If
    Else
        Exit Sub
    End If

    'Set Transfer.SourceKey = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    'Set Transfer.DestinationKey = ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(1)
NoTableSelected:
    If GetKeyColumns(Transfer) Then
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
    If SetValueMapping(Transfer) Then
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
     
    Transfer.Transfer
    
    If HasFlag(Transfer.Flags, saveToHistory) Then
        Dim history As TransferHistoryViewModel
        Set history = New TransferHistoryViewModel
        If history.HasHistory = False Then
            history.Create
        End If
        history.Refresh
        history.Add Transfer
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

Private Function GetKeyColumns(ByVal Transfer As TransferInstruction) As Boolean
    Dim vm As KeyMapperViewModel
    Set vm = New KeyMapperViewModel
    Set vm.LHSTable = Transfer.Source
    Set vm.RHSTable = Transfer.Destination
    
    Dim frm As IView
    Set frm = New KeyMapperView
    
    If frm.ShowDialog(vm) Then
        If vm.IsValid Then
            Set Transfer.Source = vm.LHSTable
            Set Transfer.Destination = vm.RHSTable
            Set Transfer.SourceKey = vm.LHSKeyColumn
            Set Transfer.DestinationKey = vm.RHSKeyColumn
            GetKeyColumns = True
        Else
            MsgBox "Invalid VM"
        End If
    End If
End Function

Private Function SetValueMapping(ByVal Transfer As TransferInstruction) As Boolean
    Dim vm As ValueMapperViewModel
    Set vm = New ValueMapperViewModel
    Set vm.lhs = Transfer.Source
    Set vm.RHS = Transfer.Destination
    Set vm.KeyColumnLHS = Transfer.SourceKey
    Set vm.KeyColumnRHS = Transfer.DestinationKey
    vm.Flags = Transfer.Flags
    
    vm.LoadFromTransferInstruction Transfer
    
    Dim frm As IView
    Set frm = New ValueMapperView
    
    If frm.ShowDialog(vm) Then
        Set Transfer.ValuePairs = vm.checked
        Transfer.Flags = vm.Flags
        SetValueMapping = True
    End If
End Function
