Attribute VB_Name = "modTableTransferTool2"
'@Folder("TransferTableTool2")
Option Explicit

Public Sub TestTableTransferTool2()
    Dim Transfer As clsTransferTableTool2
    Set Transfer = New clsTransferTableTool2
    
    If Selection.ListObject Is Nothing Then Exit Sub
    Set Transfer.Destination = Selection.ListObject
    
    Set Transfer.Source = ThisWorkbook.Worksheets(1).ListObjects(1)
    'If GetSourceTable(Transfer) Then
    '    ' continue
    'Else
    '    Exit Sub
    'End If

    'Set Transfer.SourceKey = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    'Set Transfer.DestinationKey = ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(1)
    
    If GetKeyColumns(Transfer) Then
        ' continue
    Else
        Exit Sub
    End If
    
    If SetValueMapping(Transfer) Then
        ' continue
    Else
        Exit Sub
    End If
    
    DoTransfer Transfer
End Sub

Private Function GetSourceTable(ByVal Transfer As clsTransferTableTool2) As Boolean
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

Private Function GetKeyColumns(ByVal Transfer As clsTransferTableTool2) As Boolean
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

Private Function SetValueMapping(ByVal Transfer As clsTransferTableTool2) As Boolean
    Dim vm As ValueMapperViewModel
    Set vm = New ValueMapperViewModel
    Set vm.LHS = Transfer.Source
    Set vm.RHS = Transfer.Destination
    Set vm.KeyColumnLHS = Transfer.SourceKey
    Set vm.KeyColumnRHS = Transfer.DestinationKey
    
    Dim frm As IView
    Set frm = New ValueMapperView
    
    If frm.ShowDialog(vm) Then
        Set Transfer.ValuePairs = vm.checked
        SetValueMapping = True
    End If
End Function

Private Sub DoTransfer(ByVal Transfer As clsTransferTableTool2)
    Transfer.Transfer
End Sub
