Attribute VB_Name = "modTestKeyMapper2"
'@Folder("KeyMapper")
Option Explicit

Public Sub TestKeyMapper2()
    Dim vm As clsKeyMapperViewModel
    Dim view As frmKeyMapper2
    
    Set vm = New clsKeyMapperViewModel
    Set view = New frmKeyMapper2
    Set vm.LHSTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set vm.RHSTable = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    If view.ShowDialog(vm) Then
        'Debug.Print "ShowDialog true"
    Else
        'Debug.Print "ShowDialog false"
    End If
End Sub
