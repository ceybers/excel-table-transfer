Attribute VB_Name = "modSelectTable"
'@Folder("SelectTable")
Option Explicit

Public Function SelectTable() As ListObject
    Dim vm As clsSelectTableViewModel
    Set vm = New clsSelectTableViewModel
    
    Dim frm As IView
    Set frm = frmSelectTable
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set SelectTable = vm.SelectedTable
        End If
    End If
End Function

Public Function TrySelectTable(ByRef result As ListObject) As Boolean
    Dim vm As clsSelectTableViewModel
    Set vm = New clsSelectTableViewModel
    
    Dim frm As IView
    Set frm = frmSelectTable
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set result = vm.SelectedTable
            TrySelectTable = True
        End If
    End If
End Function
