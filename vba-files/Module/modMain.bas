Attribute VB_Name = "modMain"
Option Explicit

'@Description "AAATest"
Public Sub AAATest()
Attribute AAATest.VB_Description = "AAATest"
    Dim vm As TablePropViewModel
    Set vm = New TablePropViewModel
    vm.Load ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim view As iview
    Set view = New TablePropView
    
    If view.ShowDialog(vm) Then
        Debug.Print "true"
    Else
        Debug.Print "false"
    End If
End Sub

