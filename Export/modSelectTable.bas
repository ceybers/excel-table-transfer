Attribute VB_Name = "modSelectTable"
'@IgnoreModule
'@Folder "MVVM"
Option Explicit

' 2024/03/28 used when picking 2nd table
Public Function TrySelectTable(Optional ByRef View As IView2, Optional ByRef ViewModel As SelectTableViewModel) As Boolean
    If View Is Nothing Then
        Set View = New SelectTableView
    End If
    
    If ViewModel Is Nothing Then
        Set ViewModel = New SelectTableViewModel
    End If
    
    If View.ShowDialog(ViewModel) = vrNext Then
        Debug.Assert Not ViewModel.SelectedTable Is Nothing
        TrySelectTable = True
    End If
End Function

