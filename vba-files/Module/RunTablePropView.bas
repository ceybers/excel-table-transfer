Attribute VB_Name = "RunTablePropView"
'@Folder "MVVM.TableProps"
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@Description "DoRun"
Public Sub DoRun()
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim vm As TablePropViewModel
    Set vm = New TablePropViewModel
    vm.Load ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)

    Dim view As IView
    Set view = TablePropView.Create(ctx, vm)
    
    With view
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "TablePropView.ShowDialog(vm) returned True"
            vm.Commit
        Else
            If DO_DEBUG Then Debug.Print "TablePropView.ShowDialog(vm) returned False"
        End If
    End With
End Sub

Public Sub ResetPersistentStorage()
    Dim SettingsModel As XMLSettingsModel
    Set SettingsModel = XMLSettingsModel.Create(ThisWorkbook, "TableTransferTool")
    'XMLSettingsModel.Reset
    SettingsModel.Reset
End Sub
