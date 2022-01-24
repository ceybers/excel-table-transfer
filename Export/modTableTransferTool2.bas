Attribute VB_Name = "modTableTransferTool2"
'@Folder("TransferTableTool2")
Option Explicit

Public Sub TestTableTransferTool2()
    Dim Transfer As clsTransferTableTool2
    Set Transfer = New clsTransferTableTool2
    
    If Selection.ListObject Is Nothing Then Exit Sub
    Set Transfer.Destination = Selection.ListObject
    
    'GetSourceTable transfer
    Set Transfer.Source = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    'GetKeyColumns transfer
    Set Transfer.SourceKey = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    Set Transfer.DestinationKey = ThisWorkbook.Worksheets(1).ListObjects(2).ListColumns(1)
    
    SetValueMapping Transfer
    
    DoTransfer Transfer
End Sub

Private Sub GetSourceTable(ByRef Transfer As clsTransferTableTool2)
    Dim vm As clsSelectTableViewModel
    Set vm = New clsSelectTableViewModel
    Set vm.ActiveTable = Transfer.Destination
    
    Dim frm As IView
    Set frm = New frmSelectTable
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set Transfer.Source = vm.SelectedTable
        End If
    End If
End Sub

Private Sub GetKeyColumns(ByRef Transfer As clsTransferTableTool2)
    Dim vm As clsKeyMapperViewModel
    Set vm = New clsKeyMapperViewModel
    Set vm.LHSTable = Transfer.Source
    Set vm.RHSTable = Transfer.Destination
    
    Dim frm As IView
    Set frm = New frmKeyMapper2
    
    If frm.ShowDialog(vm) Then
        If vm.IsValid Then
            Set Transfer.SourceKey = vm.LHSKeyColumn
            Set Transfer.DestinationKey = vm.RHSKeyColumn
        Else
            MsgBox "Invalid VM"
        End If
    End If
End Sub

Private Sub SetValueMapping(ByRef Transfer As clsTransferTableTool2)
    Dim vm As clsValueMapperViewModel
    Set vm = New clsValueMapperViewModel
    Set vm.lhs = Transfer.Source
    Set vm.RHS = Transfer.Destination
    Set vm.KeyColumnLHS = Transfer.SourceKey
    Set vm.KeyColumnRHS = Transfer.DestinationKey
    
    Dim frm As IView
    Set frm = New frmValueMapper2
    
    If frm.ShowDialog(vm) Then
        Set Transfer.ValuePairs = vm.Checked
    End If
End Sub

Private Sub DoTransfer(ByRef Transfer As clsTransferTableTool2)
    Transfer.Transfer
End Sub

