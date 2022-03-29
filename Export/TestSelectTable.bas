Attribute VB_Name = "TestSelectTable"
'@Folder("SelectTable")
Option Explicit
Option Private Module

Public Sub Test()
    Dim vm As SelectTableViewModel
    Set vm = New SelectTableViewModel
    Set vm.ActiveTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim view As IView
    Set view = New SelectTableView
    
    If TrySelectTable(view, vm) Then
        Debug.Print "TrySelectTable result: TRUE"
        Debug.Print " vm.SelectedTable: "; vm.SelectedTable
        Debug.Print " vm.ActiveTable: "; vm.ActiveTable
        Debug.Print " vm.AutoSelected: "; vm.AutoSelected
    Else
        Debug.Print "TrySelectTable result: FALSE"
    End If
End Sub
