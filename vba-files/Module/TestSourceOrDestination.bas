Attribute VB_Name = "TestSourceOrDestination"
'@Folder "MVVM.SourceOrDestination"
Option Explicit
Option Private Module

Public Sub Test()
    Dim vm As SourceOrDestinationViewModel
    Set vm = New SourceOrDestinationViewModel
    Set vm.ListObject = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim view As IView
    Set view = New SourceOrDestinationView
    If view.ShowDialog(vm) Then
        Debug.Print vm.IsSource; vm.IsDestination
    Else
        Debug.Print "No option selected"
    End If
End Sub
