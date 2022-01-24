Attribute VB_Name = "modTestKeyMapper2"
'@Folder("KeyMapper")
Option Explicit

Public Sub TestKeyMapper2()
    Dim vm As clsKeyMapperViewModel
    Dim view As IView
    
    Set vm = New clsKeyMapperViewModel
    
    Set view = New frmKeyMapper2
    
    ' TODO Fix
    Dim vview As frmKeyMapper2
    Set vview = frmKeyMapper2
    vview.DEBUG_EVENTS = True
    
    Set vm.LHSTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    'Set vm.RHSTable = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    If view.ShowDialog(vm) Then
        MsgBox "Hi"
        'Debug.Print "ShowDialog true"
    Else
        'Debug.Print "ShowDialog false"
    End If
End Sub

