Attribute VB_Name = "RunTablePropView"
'@Folder "MVVM.TableProps"
Option Explicit

'@Description "DoRun"
Public Sub DoRun()
Attribute DoRun.VB_Description = "AAATest"
    Dim vm As TablePropViewModel
    Set vm = New TablePropViewModel
    vm.Load ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim view As TablePropView
    Set view = New TablePropView
    
    If view.ShowDialog(vm) Then
        Debug.Print "view.ShowDialog(vm) = True"
    Else
        Debug.Print "view.ShowDialog(vm) = False"
    End If
End Sub
