Attribute VB_Name = "RunDirectionPickerView"
'@Folder("MVVM.DirectionPicker")
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@Description "DoRunDirectionPicker"
Public Sub DoRunDirectionPicker()
Attribute DoRunDirectionPicker.VB_Description = "DoRunDirectionPicker"
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim VM As DirectionPickerViewModel
    Set VM = New DirectionPickerViewModel
    VM.Load ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)

    Dim View As IView
    Set View = DirectionPickerView.Create(ctx, VM)
    
    With View
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "TablePropView.ShowDialog(vm) returned True"
            'VM.Commit
            Debug.Print "Result was = "; VM.Result
        Else
            If DO_DEBUG Then Debug.Print "TablePropView.ShowDialog(vm) returned False"
        End If
    End With
End Sub
