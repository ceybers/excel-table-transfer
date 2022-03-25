Attribute VB_Name = "TestColumnQuality"
'@IgnoreModule EmptyIfBlock
'@Folder("ColumnQuality")
Option Explicit
Option Private Module

Public Sub test()
    Dim vm As ColumnQualityViewModel
    Dim view As IView
    
    Set vm = New ColumnQualityViewModel
    Set view = New ColumnQualityView
    
    Set vm.ListColumn = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    
    If view.ShowDialog(vm) Then
        'noop
    End If
End Sub
