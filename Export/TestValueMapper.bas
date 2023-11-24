Attribute VB_Name = "TestValueMapper"
'@IgnoreModule EmptyIfBlock
'@Folder "ValueMapper"
Option Explicit
Option Private Module

'@ExcelHotkey p
Public Sub Test()
    Dim vm As ValueMapperViewModel
    Dim view As IView
    
    Set vm = New ValueMapperViewModel
    
    Set view = New ValueMapperView
    
    ' TODO Fix
    Dim vview As ValueMapperView
    Set vview = ValueMapperView
    'vview.DEBUG_EVENTS = True
    
    Set vm.lhs = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set vm.rhs = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    Set vm.KeyColumnLHS = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    Set vm.KeyColumnRHS = ThisWorkbook.Worksheets(2).ListObjects(1).ListColumns(1)
    
    If view.ShowDialog(vm) Then
        'Debug.Print "ShowDialog true"
    Else
        'Debug.Print "ShowDialog false"
    End If

End Sub

