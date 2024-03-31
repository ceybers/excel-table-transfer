Attribute VB_Name = "TestTablePicker"
'@Folder("MVVM2")
Option Explicit

Public Sub AAA()
    Dim ViewModel As TablePickerViewModel
    Set ViewModel = New TablePickerViewModel
    ViewModel.Load
    
    Dim View As IView2
    Set View = New TablePickerView
    
    Dim Result As ViewResult
    Result = View.ShowDialog(ViewModel)
    
End Sub
