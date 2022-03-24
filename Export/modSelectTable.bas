Attribute VB_Name = "modSelectTable"
'@Folder("SelectTable")
Option Explicit

Public Function TrySelectTable(ByRef result As ListObject) As Boolean
    ' TODO This is being used by key select dialog to choose new tables
    'Err.Raise 5, , "tryselecttable deprec"
    Dim vm As SelectTableViewModel
    Set vm = New SelectTableViewModel
    
    Dim frm As IView
    Set frm = SelectTableView
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set result = vm.SelectedTable
            TrySelectTable = True
        End If
    End If
End Function
