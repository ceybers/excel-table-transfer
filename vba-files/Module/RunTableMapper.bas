Attribute VB_Name = "RunTableMapper"
'@Folder "MVVM.TableMapper"
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@Description "DoRunTableMapper"
Public Sub DoRunTableMapper()
Attribute DoRunTableMapper.VB_Description = "DoRun"
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim VM As TableMapperViewModel
    Set VM = New TableMapperViewModel
    VM.Load ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)

    Dim View As IView
    Set View = TableMapperView.Create(ctx, VM)
    
    With View
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "TableMapper.ShowDialog(vm) returned True"
            Debug.Print "TableMapper result:"
            'Debug.Print " Src: "; VM.Source.Name
            'Debug.Print " Dst: "; VM.Destination.Name
            Debug.Print vbNullString
        Else
            If DO_DEBUG Then Debug.Print "TableMapper.ShowDialog(vm) returned False"
        End If
    End With
End Sub
