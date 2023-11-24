Attribute VB_Name = "modSelectTable"
'@Folder("SelectTable")
Option Explicit

Public Function TrySelectTable(Optional ByRef frm As IView, Optional ByRef vm As SelectTableViewModel) As Boolean
    If frm Is Nothing Then
        Set frm = New SelectTableView
    End If
    
    If vm Is Nothing Then
        Set vm = New SelectTableViewModel
    End If
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            TrySelectTable = True
        End If
    End If
End Function

