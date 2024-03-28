Attribute VB_Name = "TestSelectTable"
'@Folder "Tests.MVVM"
Option Explicit
Option Private Module

Public Sub Test()
    Dim vm As SelectTableViewModel
    Set vm = New SelectTableViewModel
    Set vm.ActiveTable = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim View As IView
    Set View = New SelectTableView
    
    If TrySelectTable(View, vm) Then
        Debug.Print "TrySelectTable result: TRUE"
        Debug.Print " vm.SelectedTable: "; vm.SelectedTable
        Debug.Print " vm.ActiveTable: "; vm.ActiveTable
        Debug.Print " vm.AutoSelected: "; vm.AutoSelected
    Else
        Debug.Print "TrySelectTable result: FALSE"
    End If
End Sub

