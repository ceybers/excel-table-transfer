Attribute VB_Name = "TestSourceOrDestination"
'@Folder "SourceOrDestination"
Option Explicit
Option Private Module

Public Function Test() As ListObject
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
End Function
