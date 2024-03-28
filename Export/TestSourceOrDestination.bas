Attribute VB_Name = "TestSourceOrDestination"
'@Folder "Tests.MVVM"
Option Explicit
Option Private Module

Public Sub Test()
    Dim vm As SourceOrDestinationViewModel
    Set vm = New SourceOrDestinationViewModel
    Set vm.ListObject = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    
    Dim View As IView
    Set View = New SourceOrDestinationView
    If View.ShowDialog(vm) Then
        Debug.Print vm.IsSource; vm.IsDestination
    Else
        Debug.Print "No option selected"
    End If
End Sub

