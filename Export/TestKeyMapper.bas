Attribute VB_Name = "TestKeyMapper"
'@Folder("KeyMapper")
Option Explicit
Option Private Module

Public Sub Test()
    Dim vm As KeyMapperViewModel
    Dim view As IView
    
    Set vm = New KeyMapperViewModel
    
    Set view = New KeyMapperView
    
    ' TODO Fix
    Dim vview As KeyMapperView
    Set vview = KeyMapperView
    vview.DEBUG_EVENTS = True
    
    Set vm.LHSTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set vm.RHSTable = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    If view.ShowDialog(vm) Then
        Debug.Print "ShowDialog true"
    Else
        Debug.Print "ShowDialog false"
    End If
End Sub

