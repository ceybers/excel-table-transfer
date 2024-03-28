Attribute VB_Name = "TestColumnQuality"
'@IgnoreModule EmptyIfBlock
'@Folder "Tests.MVVM"
Option Explicit
Option Private Module

'@Description "Displays a TestColumnQuality MVVM UserForm"
Public Sub TestColumnQualityView()
    Dim ViewModel As ColumnQualityViewModel
    Set ViewModel = New ColumnQualityViewModel
    Set ViewModel.ListColumn = ThisWorkbook.Worksheets(1).ListObjects(1).ListColumns(1)
    
    Dim View As IView
    Set View = New ColumnQualityView

    If View.ShowDialog(ViewModel) Then
        Debug.Print "ColumnQualityView.ShowDialog returned TRUE"
    Else
        Debug.Print "ColumnQualityView.ShowDialog returned FALSE"
    End If
End Sub

