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
