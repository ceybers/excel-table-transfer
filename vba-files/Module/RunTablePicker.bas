Attribute VB_Name = "RunTablePicker"
'@Folder("MVVM.TablePicker")
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@Description "DoRun"
Public Sub DoRunTablePicker()
Attribute DoRunTablePicker.VB_Description = "DoRun"
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim VM As TablePickerViewModel
    Set VM = New TablePickerViewModel
    VM.Load ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)

    Dim View As IView
    Set View = TablePickerView.Create(ctx, VM)
    
    With View
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "TablePicker.ShowDialog(vm) returned True"
            Debug.Print "TablePicker result:"
            Debug.Print " Src: "; VM.Source.Name
            Debug.Print " Src: "; VM.Destination.Name
            Debug.Print vbNullString
        Else
            If DO_DEBUG Then Debug.Print "TablePicker.ShowDialog(vm) returned False"
        End If
    End With
End Sub
