Attribute VB_Name = "RunValueMapper"
'@Folder("MVVM.ValueMapper")
Option Explicit
Private Const DO_DEBUG As Boolean = False

'@Description "DoRunValueMapper"
Public Sub DoRunValueMapper()
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim VM As ValueMapperViewModel
    Set VM = New ValueMapperViewModel
    VM.Load _
        SrcTable:=ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1), _
        DstTable:=ThisWorkbook.Worksheets.Item(3).ListObjects.Item(1)

    Dim View As IView
    Set View = ValueMapperView.Create(ctx, VM)
    
    With View
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "ValueMapper.ShowDialog(vm) returned True"
            Debug.Print "ValueMapper results:"
            Debug.Print vbNullString
        Else
            If DO_DEBUG Then Debug.Print "ValueMapper.ShowDialog(vm) returned False"
        End If
    End With
End Sub
