Attribute VB_Name = "modTestValueMapper2"
'@Folder("ValueMapper2")
Option Explicit

Public Sub TestValueMapper2()
Attribute TestValueMapper2.VB_ProcData.VB_Invoke_Func = "p\n14"
    Dim vm As clsValueMapperViewModel
    Dim view As IView
    
    Set vm = New clsValueMapperViewModel
    
    Set view = New frmValueMapper2
    
    ' TODO Fix
    Dim vview As frmValueMapper2
    Set vview = frmValueMapper2
    'vview.DEBUG_EVENTS = True
    
    Set vm.lhs = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set vm.RHS = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    Set vm.KeyColumnLHS = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    Set vm.KeyColumnRHS = ThisWorkbook.Worksheets(2).ListObjects(1).ListColumns(1)
    
    If view.ShowDialog(vm) Then
        'Debug.Print "ShowDialog true"
    Else
        'Debug.Print "ShowDialog false"
    End If

End Sub
