Attribute VB_Name = "RunKeyMapper"
'@Folder("MVVM.KeyMapper")
Option Explicit
Private Const DO_DEBUG As Boolean = False

'@Description "DoRunKeyMapper"
Public Sub DoRunKeyMapper()
Attribute DoRunKeyMapper.VB_Description = "DoRunKeyMapper"
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim VM As KeyMapperViewModel
    Set VM = New KeyMapperViewModel
    VM.Load _
        Context:=ctx, _
        SrcTable:=ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1), _
        DstTable:=ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)

    Dim View As IView
    Set View = KeyMapperView.Create(ctx, VM)
    
    With View
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "KeyMapper.ShowDialog(vm) returned True"
            Debug.Print "KeyMapper result:"
            Debug.Print " Src: "; VM.SrcKeyColumnVM.SelectedAsText
            Debug.Print " Dst: "; VM.DstKeyColumnVM.SelectedAsText
            Debug.Print vbNullString
        Else
            If DO_DEBUG Then Debug.Print "KeyMapper.ShowDialog(vm) returned False"
        End If
    End With
End Sub
