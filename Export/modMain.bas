Attribute VB_Name = "modMain"
'@Folder "TableTransfer"
Option Explicit

Public Sub TransferTable()
    Dim Transfer As TransferInstruction
    Set Transfer = New TransferInstruction
    
    If Selection.ListObject Is Nothing Then Exit Sub
    Set Transfer.Destination = Selection.ListObject
    
    Set Transfer.Source = ThisWorkbook.Worksheets(1).ListObjects(1)
    'If GetSourceTable(Transfer) Then
    '    ' continue
    'Else
    '    Exit Sub
    'End If

    Set Transfer.SourceKey = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    Set Transfer.DestinationKey = ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(1)
    
    'If GetKeyColumns(Transfer) Then
    '    ' continue
    'Else
        'Exit Sub
    'End If
    
    ' TODO Move default Flags somewhere better
    Transfer.Flags = AddFlag(Transfer.Flags, ClearDestinationFirst)
    'Transfer.Flags = AddFlag(Transfer.Flags, ReplaceEmptyOnly)
    'Transfer.Flags = AddFlag(Transfer.Flags, TransferBlanks)
    'Transfer.Flags = AddFlag(Transfer.Flags, SourceFilteredOnly)
    Transfer.Flags = AddFlag(Transfer.Flags, DestinationFilteredOnly)
    
    If SetValueMapping(Transfer) Then
        ' continue
    Else
        Exit Sub
    End If
    
    
     
    Transfer.Transfer
End Sub

Private Function GetSourceTable(ByVal Transfer As TransferInstruction) As Boolean
    Dim vm As SelectTableViewModel
    Set vm = New SelectTableViewModel
    Set vm.ActiveTable = Transfer.Destination
    
    Dim frm As IView
    Set frm = New SelectTableView
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set Transfer.Source = vm.SelectedTable
            GetSourceTable = True
        End If
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
    
    Dim frm As IView
    Set frm = New ValueMapperView
    
    If frm.ShowDialog(vm) Then
        Set Transfer.ValuePairs = vm.checked
        Transfer.Flags = vm.Flags
        SetValueMapping = True
    End If
End Function
