Attribute VB_Name = "modTestColumnQuality"
'@Folder("ColumnQuality")
Option Explicit

Public Sub AAATest()
    Dim vm As clsColumnQualityViewModel
    Dim view As IView
    
    Set vm = New clsColumnQualityViewModel
    Set view = New frmColumnQuality
    
    Set vm.ListColumn = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    
    If view.ShowDialog(vm) Then
        'noop
    End If
End Sub
