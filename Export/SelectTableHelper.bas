Attribute VB_Name = "SelectTableHelper"
'@IgnoreModule
'@Folder "MVVM.Miscellaneous"
Option Explicit

Public Function TrySelectTable(ByVal ActiveTable As ListObject, ByRef OutListObject As ListObject) As Boolean
    Dim View As IView2
    Set View = New SelectTableView
    
    Dim ViewModel As SelectTableViewModel
    Set ViewModel = New SelectTableViewModel

    If Not ActiveTable Is Nothing Then
        Set ViewModel.ActiveTable = ActiveTable
    End If
    
    If View.ShowDialog(ViewModel) = vrNext Then
        Set OutListObject = ViewModel.SelectedTable
        TrySelectTable = True
    End If
End Function

