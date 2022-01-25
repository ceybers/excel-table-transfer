Attribute VB_Name = "TestSelectTable"
'@Folder("SelectTable")
Option Explicit
Option Private Module

Public Function Test() As ListObject
    Dim vm As SelectTableViewModel
    Set vm = New SelectTableViewModel
    Set vm.ActiveTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim view As IView
    Set view = New SelectTableView
    If view.ShowDialog(vm) Then
        Debug.Print vm.SelectedTable.Name
    Else
        Debug.Print "No table selected"
    End If
End Function

Public Function TrySelectTable(ByRef Result As ListObject) As Boolean
    ' TODO This is being used by key select dialog to choose new tables
    'Err.Raise 5, , "tryselecttable deprec"
    Dim vm As SelectTableViewModel
    Set vm = New SelectTableViewModel
    
    Dim frm As IView
    Set frm = SelectTableView
    
    If frm.ShowDialog(vm) Then
        If Not vm.SelectedTable Is Nothing Then
            Set Result = vm.SelectedTable
            TrySelectTable = True
        End If
    End If
End Function
