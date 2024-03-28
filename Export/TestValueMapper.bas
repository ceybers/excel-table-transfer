Attribute VB_Name = "TestValueMapper"
'@IgnoreModule
'@Folder "Tests.MVVM"
Option Explicit
Option Private Module

'@ExcelHotkey p
Public Sub Test()
Attribute Test.VB_ProcData.VB_Invoke_Func = "p\n14"
    Dim vm As ValueMapperViewModel
    Dim View As IView
    
    Set vm = New ValueMapperViewModel
    
    Set View = New ValueMapperView
    
    ' TODO Fix
    Dim vview As ValueMapperView
    Set vview = ValueMapperView
    'vview.DEBUG_EVENTS = True
    
    Set vm.LHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    Set vm.RHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
    
    Set vm.KeyColumnLHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Item(1)
    Set vm.KeyColumnRHS = ThisWorkbook.Worksheets.Item(2).ListObjects.Item(1).ListColumns.Item(1)
    
    If View.ShowDialog(vm) Then
        'Debug.Print "ShowDialog true"
    Else
        'Debug.Print "ShowDialog false"
    End If

End Sub
