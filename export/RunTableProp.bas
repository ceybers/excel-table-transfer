Attribute VB_Name = "RunTableProp"
'@Folder "MVVM.TableProps"
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@Description "DoRunTableProp"
Public Sub DoRunTableProp()
Attribute DoRunTableProp.VB_Description = "DoRunTableProp"
    Dim Context As AppContext
    Set Context = New AppContext
    Context.LoadSettings ThisWorkbook
    
    Dim TablePropVM As TablePropViewModel
    Set TablePropVM = New TablePropViewModel
    TablePropVM.Load _
        Context:=Context, _
        ListObject:=ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)

    Dim TablePropV As IView
    Set TablePropV = TablePropView.Create(Context, TablePropVM)
    
    With TablePropV
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "TablePropView.ShowDialog(vm) returned True"
            TablePropVM.Commit
        Else
            If DO_DEBUG Then Debug.Print "TablePropView.ShowDialog(vm) returned False"
        End If
    End With
End Sub
